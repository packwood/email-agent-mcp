import { mkdtemp, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterEach, describe, expect, it, vi } from 'vitest';
import {
  MatonGraphApiClient,
  loadMatonOutlookConnections,
  matonGraphUrl,
  resolveMatonOutlookConnection,
} from './maton.js';

afterEach(() => {
  vi.unstubAllGlobals();
});

describe('Maton Outlook connection isolation', () => {
  it('loads only exact ACTIVE Outlook accounts and resolves case-insensitively', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'maton-connections-'));
    const path = join(dir, 'connections.json');
    await writeFile(path, JSON.stringify({
      connections: [
        { app: 'outlook', account: 'User@Example.com', connection_id: 'conn-1', status: 'ACTIVE' },
        { app: 'google-mail', account: 'user@example.com', connection_id: 'wrong', status: 'ACTIVE' },
        { app: 'outlook', account: null, connection_id: 'ignored', status: null },
      ],
    }));

    const connections = await loadMatonOutlookConnections(path);
    expect(connections.size).toBe(1);
    expect(resolveMatonOutlookConnection(connections, ' user@example.COM ')).toEqual({
      account: 'user@example.com',
      connectionId: 'conn-1',
    });
    expect(() => resolveMatonOutlookConnection(connections, 'other@example.com'))
      .toThrow('No ACTIVE Maton Outlook connection');
  });

  it('fails closed on duplicate or inactive Outlook account records', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'maton-connections-'));
    const duplicatePath = join(dir, 'duplicate.json');
    const inactivePath = join(dir, 'inactive.json');
    await writeFile(duplicatePath, JSON.stringify({
      connections: [
        { app: 'outlook', account: 'user@example.com', connection_id: 'a', status: 'ACTIVE' },
        { app: 'outlook', account: 'USER@example.com', connection_id: 'b', status: 'ACTIVE' },
      ],
    }));
    await writeFile(inactivePath, JSON.stringify({
      connections: [
        { app: 'outlook', account: 'user@example.com', connection_id: 'a', status: 'EXPIRED' },
      ],
    }));

    await expect(loadMatonOutlookConnections(duplicatePath)).rejects.toThrow('Duplicate');
    await expect(loadMatonOutlookConnections(inactivePath)).rejects.toThrow('not ACTIVE');
  });

  it('ignores an unrelated inactive Outlook record when expected accounts are scoped', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'maton-connections-'));
    const path = join(dir, 'connections.json');
    await writeFile(path, JSON.stringify({
      connections: [
        { app: 'outlook', account: 'wanted@example.com', connection_id: 'wanted', status: 'ACTIVE' },
        { app: 'outlook', account: 'old@example.com', connection_id: 'old', status: 'EXPIRED' },
      ],
    }));

    const connections = await loadMatonOutlookConnections(path, ['wanted@example.com']);
    expect([...connections.keys()]).toEqual(['wanted@example.com']);
  });
});

describe('Maton Graph URL boundary', () => {
  it('rewrites relative and Microsoft Graph v1 URLs without changing path or query', () => {
    expect(matonGraphUrl('/me/messages?$top=1')).toBe(
      'https://gateway.maton.ai/outlook/v1.0/me/messages?$top=1',
    );
    expect(matonGraphUrl('https://graph.microsoft.com/v1.0/me/messages?$skiptoken=abc')).toBe(
      'https://gateway.maton.ai/outlook/v1.0/me/messages?$skiptoken=abc',
    );
    expect(matonGraphUrl('https://gateway.maton.ai/outlook/v1.0/me/messages')).toBe(
      'https://gateway.maton.ai/outlook/v1.0/me/messages',
    );
  });

  it.each([
    'http://graph.microsoft.com/v1.0/me/messages',
    'https://graph.microsoft.com.evil.example/v1.0/me/messages',
    'https://gateway.maton.ai.evil.example/outlook/v1.0/me/messages',
    'https://gateway.maton.ai/other/v1.0/me/messages',
    'https://user:pass@gateway.maton.ai/outlook/v1.0/me/messages',
  ])('rejects an untrusted URL: %s', url => {
    expect(() => matonGraphUrl(url)).toThrow('Untrusted Microsoft Graph URL');
  });
});

describe('MatonGraphApiClient', () => {
  it('sends credentials only to the Maton gateway and parses JSON', async () => {
    const fetchMock = vi.fn(async () => new Response(JSON.stringify({ value: [{ id: '1' }] }), {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
    }));
    vi.stubGlobal('fetch', fetchMock);
    const client = new MatonGraphApiClient('secret-key', 'connection-1', 1000);

    await expect(client.get('/me/messages')).resolves.toEqual({ value: [{ id: '1' }] });
    expect(fetchMock).toHaveBeenCalledTimes(1);
    const [url, init] = fetchMock.mock.calls[0]!;
    expect(url).toBe('https://gateway.maton.ai/outlook/v1.0/me/messages');
    expect(init.headers).toMatchObject({
      Authorization: 'Bearer secret-key',
      'Maton-Connection': 'connection-1',
    });
  });

  it('does not retry a failed mutation and does not put secrets in its error', async () => {
    const fetchMock = vi.fn(async () => new Response('upstream rejected request', { status: 503 }));
    vi.stubGlobal('fetch', fetchMock);
    const client = new MatonGraphApiClient('secret-key', 'connection-1', 1000);

    await expect(client.post('/me/sendMail', { message: {} })).rejects.toThrow(
      'Graph API error 503: upstream rejected request',
    );
    expect(fetchMock).toHaveBeenCalledTimes(1);
    try {
      await client.post('/me/sendMail', { message: {} });
    } catch (err) {
      expect(String(err)).not.toContain('secret-key');
      expect(String(err)).not.toContain('connection-1');
    }
  });
});
