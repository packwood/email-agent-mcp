import { afterEach, describe, expect, it, vi } from 'vitest';
import { MatonGraphApiClient } from './maton.js';
import { MAX_READ_ATTEMPTS, readRetryDelayMs } from './throttle.js';

const KEY = 'maton-key';
const CONNECTION = 'conn-123';

function throttled(headers: Record<string, string> = {}): Response {
  return new Response('{"error":"throttled"}', { status: 429, headers });
}

function ok(body: unknown = { value: [] }): Response {
  return new Response(JSON.stringify(body), { status: 200, headers: { 'Content-Type': 'application/json' } });
}

afterEach(() => {
  vi.restoreAllMocks();
});

describe('provider-microsoft/throttle delay policy', () => {
  it('prefers a numeric Retry-After, in seconds', () => {
    expect(readRetryDelayMs(throttled({ 'retry-after': '2' }), 0)).toBe(2000);
  });

  it('accepts an HTTP-date Retry-After', () => {
    const delay = readRetryDelayMs(throttled({ 'retry-after': new Date(Date.now() + 3000).toUTCString() }), 0);
    // Whole-second resolution in the header, so allow a small band.
    expect(delay).toBeGreaterThan(1500);
    expect(delay).toBeLessThanOrEqual(3000);
  });

  it('caps an absurd Retry-After so a briefing cannot stall indefinitely', () => {
    expect(readRetryDelayMs(throttled({ 'retry-after': '86400' }), 0)).toBe(30_000);
  });

  it('never returns a negative delay for a Retry-After date in the past', () => {
    expect(readRetryDelayMs(throttled({ 'retry-after': new Date(Date.now() - 60_000).toUTCString() }), 0)).toBe(0);
  });

  it('falls back to bounded jittered backoff when the header is absent or junk', () => {
    for (const response of [throttled(), throttled({ 'retry-after': 'soon' })]) {
      for (let attempt = 0; attempt < 6; attempt += 1) {
        const delay = readRetryDelayMs(response, attempt);
        expect(delay).toBeGreaterThanOrEqual(0);
        expect(delay).toBeLessThanOrEqual(8_000);
      }
    }
  });
});

describe('provider-microsoft/MatonGraphApiClient throttling', () => {
  it('retries a throttled read and returns the eventual success', async () => {
    const fetchMock = vi.spyOn(globalThis, 'fetch' as never)
      .mockResolvedValueOnce(throttled({ 'retry-after': '0' }) as never)
      .mockResolvedValueOnce(throttled({ 'retry-after': '0' }) as never)
      .mockResolvedValueOnce(ok({ value: [{ id: 'm-1' }] }) as never);

    const client = new MatonGraphApiClient(KEY, CONNECTION);
    await expect(client.get('/me/messages')).resolves.toMatchObject({ value: [{ id: 'm-1' }] });
    expect(fetchMock).toHaveBeenCalledTimes(3);
  });

  it('gives up after the read attempt budget and surfaces the 429', async () => {
    // Fresh Response per call: real fetch never hands back the same object twice.
    const fetchMock = vi.spyOn(globalThis, 'fetch' as never)
      .mockImplementation((() => Promise.resolve(throttled({ 'retry-after': '0' }))) as never);

    const client = new MatonGraphApiClient(KEY, CONNECTION);
    await expect(client.get('/me/messages')).rejects.toMatchObject({ status: 429 });
    expect(fetchMock).toHaveBeenCalledTimes(MAX_READ_ATTEMPTS);
  });

  it('does NOT retry a throttled write', async () => {
    // A throttled write may or may not have been applied. Re-sending mail is
    // worse than surfacing the error, so writes get exactly one attempt.
    const fetchMock = vi.spyOn(globalThis, 'fetch' as never)
      .mockImplementation((() => Promise.resolve(throttled({ 'retry-after': '0' }))) as never);

    const client = new MatonGraphApiClient(KEY, CONNECTION);
    await expect(client.post('/me/sendMail', { message: {} })).rejects.toMatchObject({ status: 429 });
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });

  it('arms a fresh deadline per attempt rather than reusing one signal', async () => {
    const signals: Array<AbortSignal | undefined> = [];
    vi.spyOn(globalThis, 'fetch' as never).mockImplementation(((_url: string, init: RequestInit) => {
      signals.push(init.signal ?? undefined);
      return Promise.resolve(signals.length < 2 ? throttled({ 'retry-after': '0' }) : ok());
    }) as never);

    const client = new MatonGraphApiClient(KEY, CONNECTION);
    await client.get('/me/messages');
    expect(signals).toHaveLength(2);
    // Reusing one already-armed signal across retries would abort the retry the
    // moment the first attempt's deadline elapsed.
    expect(signals[0]).toBeDefined();
    expect(signals[1]).toBeDefined();
    expect(signals[0]).not.toBe(signals[1]);
  });

  it('does not retry a non-throttling error', async () => {
    const fetchMock = vi.spyOn(globalThis, 'fetch' as never)
      .mockResolvedValue(new Response('nope', { status: 404 }) as never);

    const client = new MatonGraphApiClient(KEY, CONNECTION);
    await expect(client.get('/me/messages')).rejects.toMatchObject({ status: 404 });
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });
});
