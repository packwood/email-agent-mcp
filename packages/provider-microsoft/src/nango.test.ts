import { afterAll, afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import {
  NangoGrantExpiredError,
  NangoIdentityCheckError,
  NangoIdentityMismatchError,
  NangoInvalidResponseError,
  NangoNotConfiguredError,
  NangoNotFoundError,
  NangoRefusedError,
  NangoTokenSource,
  NangoTransportError,
  NangoUnreachableError,
  parseNangoOutlookConnections,
  resolveNangoHost,
  type NangoFetch,
  type NangoTokenSourceOptions,
} from './nango.js';

const ACCOUNT = 'joshua@example.com';
const SECRET = 'nsk_7f3a9c_secret';
const TOKEN = 'tok_7f3a9c_access';
const CONNECTION_ID = 'conn+7f3a9c@id';
const PROVIDER_KEY = 'microsoft';
const OTHER_ADDRESS = 'someone-else@contoso.example';
const LEAK = 'LEAK7f3a9cBODY';
const SECRETS = [SECRET, TOKEN, CONNECTION_ID, LEAK, OTHER_ADDRESS];

const clock = { now: Date.parse('2026-09-24T12:00:00.000Z') };
const consoleLines: string[] = [];
const seenUrls: string[] = [];

function flatten(value: unknown): string {
  if (typeof value === 'string') return value;
  if (value instanceof Error) return `${value.name}\n${value.message}\n${value.stack ?? ''}`;
  try {
    return JSON.stringify(value);
  } catch {
    return String(value);
  }
}

beforeEach(() => {
  clock.now = Date.parse('2026-09-24T12:00:00.000Z');
  for (const method of ['log', 'error', 'warn'] as const) {
    vi.spyOn(console, method).mockImplementation((...args: unknown[]) => {
      consoleLines.push(args.map(flatten).join(' '));
    });
  }
});

afterEach(() => {
  vi.restoreAllMocks();
});

afterAll(() => {
  const logged = consoleLines.join('\n');
  for (const secret of SECRETS) {
    expect(logged).not.toContain(secret);
  }
  expect(seenUrls.filter(url => url.includes('/proxy'))).toEqual([]);
});

function json(body: unknown, status = 200): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: { 'Content-Type': 'application/json' },
  });
}

function poisoned(status: number, body: string = JSON.stringify({
  access_token: TOKEN,
  connection_id: CONNECTION_ID,
  secret: SECRET,
  leak: LEAK,
  mail: OTHER_ADDRESS,
})): Response {
  return new Response(body, { status, headers: { 'Content-Type': 'application/json' } });
}

function nangoBody(expiresAt?: string | null): Record<string, unknown> {
  const credentials: Record<string, unknown> = { access_token: TOKEN };
  if (expiresAt !== null) {
    credentials.expires_at = expiresAt ?? new Date(clock.now + 60 * 60 * 1000).toISOString();
  }
  return { credentials, connection_id: CONNECTION_ID, leak: LEAK };
}

function graphIdentity(fields: { mail?: unknown; userPrincipalName?: unknown }): Response {
  return json({ leak: LEAK, ...fields });
}

function createSource(
  overrides: Partial<NangoTokenSourceOptions> = {},
  fetchImpl?: NangoFetch,
): NangoTokenSource {
  const inner = fetchImpl ?? overrides.fetchImpl;
  const tracking: NangoFetch = async (url, init) => {
    seenUrls.push(url);
    if (!inner) {
      if (url.startsWith('https://graph.microsoft.com/')) return graphIdentity({ mail: ACCOUNT });
      return json(nangoBody());
    }
    return inner(url, init);
  };
  return new NangoTokenSource({
    account: ACCOUNT,
    providerConfigKey: PROVIDER_KEY,
    connectionId: CONNECTION_ID,
    secretKey: SECRET,
    host: 'https://nango.example.com',
    now: () => clock.now,
    ...overrides,
    fetchImpl: tracking,
  });
}

function expectSecretFree(err: unknown): void {
  const error = err as Error & { code?: string };
  const blob = `${error.message}\n${String(err)}\n${error.stack ?? ''}`;
  for (const secret of [SECRET, TOKEN, CONNECTION_ID, LEAK]) {
    expect(blob).not.toContain(secret);
  }
  expect(error.message).toContain(ACCOUNT);
}

async function rejectionFrom(run: () => Promise<unknown>): Promise<NangoTransportError> {
  try {
    await run();
  } catch (err) {
    expectSecretFree(err);
    return err as NangoTransportError;
  }
  throw new Error('expected a rejection');
}

describe('provider-microsoft/Nango Outlook connections', () => {
  const connectionId = 'LeakMeConnectionId99';

  it('parses a valid list and normalises the account', () => {
    const parsed = parseNangoOutlookConnections(JSON.stringify([
      {
        account: ' Joshua@Example.com ',
        providerConfigKey: 'microsoft',
        connectionId: 'conn~ok.1',
      },
      {
        account: 'other@example.com',
        providerConfigKey: 'a',
        connectionId: 'AZaz09._~:@+-',
      },
    ]));

    expect(parsed.size).toBe(2);
    expect(parsed.get('joshua@example.com')).toEqual({
      account: 'joshua@example.com',
      providerConfigKey: 'microsoft',
      connectionId: 'conn~ok.1',
    });
    expect(parsed.get('other@example.com')?.providerConfigKey).toBe('a');
    expect(parseNangoOutlookConnections(undefined).size).toBe(0);
    expect(parseNangoOutlookConnections('').size).toBe(0);
  });

  it('rejects a duplicate account without echoing the connection id', () => {
    const raw = JSON.stringify([
      { account: 'Joshua@Example.com', providerConfigKey: 'microsoft', connectionId },
      { account: ' joshua@example.com ', providerConfigKey: 'microsoft', connectionId: 'other-conn-id' },
    ]);
    expect(() => parseNangoOutlookConnections(raw)).toThrow(
      'Duplicate Nango Outlook connection for joshua@example.com',
    );
    try {
      parseNangoOutlookConnections(raw);
    } catch (err) {
      expect(String(err)).not.toContain(connectionId);
      expect(String(err)).not.toContain('other-conn-id');
      expect(String(err)).not.toContain(raw);
    }
  });

  it.each([
    ['bad connection id', { account: ACCOUNT, providerConfigKey: 'microsoft', connectionId: 'bad id/secret' }],
    ['bad provider key', { account: ACCOUNT, providerConfigKey: 'Microsoft', connectionId }],
    ['extra field', { account: ACCOUNT, providerConfigKey: 'microsoft', connectionId, region: 'us' }],
    ['missing @', { account: 'not-an-email', providerConfigKey: 'microsoft', connectionId }],
    ['not an object', 'mailbox'],
  ])('rejects %s without echoing the connection id or the raw JSON', (_label, item) => {
    const raw = JSON.stringify([item]);
    expect(() => parseNangoOutlookConnections(raw)).toThrow('Invalid Nango Outlook connections configuration');
    try {
      parseNangoOutlookConnections(raw);
    } catch (err) {
      expect(String(err)).not.toContain(connectionId);
      expect(String(err)).not.toContain('bad id/secret');
      expect(String(err)).not.toContain(raw);
    }
  });

  it('rejects invalid JSON without echoing the document', () => {
    const raw = `[{"account":"${ACCOUNT}","providerConfigKey":"microsoft","connectionId":"${connectionId}"`;
    expect(() => parseNangoOutlookConnections(raw)).toThrow('Invalid Nango Outlook connections configuration');
    try {
      parseNangoOutlookConnections(raw);
    } catch (err) {
      expect(String(err)).not.toContain(connectionId);
      expect(String(err)).not.toContain(raw);
    }
  });
});

describe('provider-microsoft/Nango host', () => {
  it('defaults when omitted and accepts a custom https origin', () => {
    expect(resolveNangoHost(undefined)).toBe('https://api.nango.dev');
    expect(resolveNangoHost('')).toBe('https://api.nango.dev');
    expect(resolveNangoHost('  https://nango.example.com/  ')).toBe('https://nango.example.com');
    expect(resolveNangoHost('https://nango.example.com')).toBe('https://nango.example.com');
  });

  it.each([
    'http://nango.example.com',
    'http://nango.example.com/proxy',
    'https://user:pass@nango.example.com',
    'https://user:pass@nango.example.com/connection',
    'https://nango.example.com/proxy',
    'https://nango.example.com/v1',
    'https://nango.example.com?x=1',
    'https://nango.example.com/#frag',
    'not a url',
  ])('rejects %s without echoing it', raw => {
    expect(() => resolveNangoHost(raw)).toThrow('Invalid Nango host');
    try {
      resolveNangoHost(raw);
    } catch (err) {
      expect(String(err)).not.toContain(raw);
      expect(String(err)).not.toContain('pass');
    }
  });
});

describe('provider-microsoft/Nango token source', () => {
  it('requests the connection URL with the secret and returns the token', async () => {
    const calls: Array<{ url: string; init?: RequestInit }> = [];
    const source = createSource({}, async (url, init) => {
      calls.push({ url, init });
      if (url.startsWith('https://graph.microsoft.com/')) return graphIdentity({ mail: ` ${ACCOUNT.toUpperCase()} ` });
      return json(nangoBody());
    });

    await expect(source.getAccessToken()).resolves.toBe(TOKEN);
    expect(calls[0]?.url).toBe(
      `https://nango.example.com/connection/${encodeURIComponent(CONNECTION_ID)}?provider_config_key=${encodeURIComponent(PROVIDER_KEY)}`,
    );
    expect(calls[0]?.url).not.toContain('/proxy');
    expect(calls[0]?.url).not.toContain(SECRET);
    expect(calls[0]?.init?.method).toBe('GET');
    expect(calls[0]?.init?.redirect).toBe('error');
    expect(calls[0]?.init?.headers).toMatchObject({ Authorization: `Bearer ${SECRET}` });
    expect(calls[1]?.url).toBe('https://graph.microsoft.com/v1.0/me?$select=mail,userPrincipalName');
    expect(calls[1]?.init?.headers).toMatchObject({ Authorization: `Bearer ${TOKEN}` });
    expect(source.getTokenHealthWarning()).toBeUndefined();
  });

  it('serves a second caller from cache', async () => {
    let nangoCalls = 0;
    const source = createSource({}, async url => {
      if (url.startsWith('https://graph.microsoft.com/')) return graphIdentity({ mail: ACCOUNT });
      nangoCalls += 1;
      return json(nangoBody());
    });

    await expect(source.getAccessToken()).resolves.toBe(TOKEN);
    await expect(source.getAccessToken()).resolves.toBe(TOKEN);
    expect(nangoCalls).toBe(1);
  });

  it('refetches once the 5-minute expiry window is reached', async () => {
    let nangoCalls = 0;
    const expiresAt = new Date(clock.now + 10 * 60 * 1000).toISOString();
    const source = createSource({}, async url => {
      if (url.startsWith('https://graph.microsoft.com/')) return graphIdentity({ mail: ACCOUNT });
      nangoCalls += 1;
      return json(nangoBody(expiresAt));
    });

    await source.getAccessToken();
    clock.now += 5 * 60 * 1000 - 1;
    await source.getAccessToken();
    expect(nangoCalls).toBe(1);
    clock.now += 1;
    await source.getAccessToken();
    expect(nangoCalls).toBe(2);
  });

  it('does not cache a token already inside the 5-minute window', async () => {
    let nangoCalls = 0;
    const expiresAt = new Date(clock.now + 4 * 60 * 1000).toISOString();
    const source = createSource({}, async url => {
      if (url.startsWith('https://graph.microsoft.com/')) return graphIdentity({ mail: ACCOUNT });
      nangoCalls += 1;
      return json(nangoBody(expiresAt));
    });

    await expect(source.getAccessToken()).resolves.toBe(TOKEN);
    await expect(source.getAccessToken()).resolves.toBe(TOKEN);
    expect(nangoCalls).toBe(2);
  });

  it('caches for 60 seconds when expires_at is absent or unparseable', async () => {
    for (const expiresAt of [null, 'not-a-date'] as const) {
      let nangoCalls = 0;
      clock.now = Date.parse('2026-09-24T12:00:00.000Z');
      const source = createSource({}, async url => {
        if (url.startsWith('https://graph.microsoft.com/')) return graphIdentity({ mail: ACCOUNT });
        nangoCalls += 1;
        return json(nangoBody(expiresAt));
      });

      await source.getAccessToken();
      clock.now += 60_000 - 1;
      await source.getAccessToken();
      expect(nangoCalls).toBe(1);
      clock.now += 1;
      await source.getAccessToken();
      expect(nangoCalls).toBe(2);
    }
  });

  it('shares one in-flight Nango request between concurrent callers', async () => {
    let nangoCalls = 0;
    let release!: (response: Response) => void;
    const source = createSource({}, async url => {
      if (url.startsWith('https://graph.microsoft.com/')) return graphIdentity({ userPrincipalName: ACCOUNT });
      nangoCalls += 1;
      return new Promise<Response>(resolve => {
        release = resolve;
      });
    });

    const first = source.getAccessToken();
    const second = source.getAccessToken();
    await Promise.resolve();
    expect(nangoCalls).toBe(1);
    release(json(nangoBody()));
    await expect(first).resolves.toBe(TOKEN);
    await expect(second).resolves.toBe(TOKEN);
    expect(nangoCalls).toBe(1);
  });

  it.each([
    [401, NangoRefusedError, 'NANGO_REFUSED'],
    [403, NangoRefusedError, 'NANGO_REFUSED'],
    [400, NangoRefusedError, 'NANGO_REFUSED'],
    [404, NangoNotFoundError, 'NANGO_CONNECTION_NOT_FOUND'],
    [424, NangoGrantExpiredError, 'NANGO_GRANT_EXPIRED'],
    [500, NangoUnreachableError, 'NANGO_UNREACHABLE'],
    [503, NangoUnreachableError, 'NANGO_UNREACHABLE'],
  ] as const)('maps HTTP %s to %s', async (status, ctor, code) => {
    const source = createSource({}, async () => poisoned(status));
    const err = await rejectionFrom(() => source.getAccessToken());
    expect(err).toBeInstanceOf(ctor);
    expect(err).toBeInstanceOf(NangoTransportError);
    expect(err.code).toBe(code);
    expect(err.message).toContain(`(HTTP ${status})`);
    if (status === 424) {
      expect(err.message).toBe(
        `Nango could not refresh the Microsoft grant for ${ACCOUNT} (HTTP 424); reconnect this mailbox in Nango`,
      );
      expect(source.getTokenHealthWarning()).toBe(err.message);
    } else {
      expect(source.getTokenHealthWarning()).toBeUndefined();
    }
  });

  it('maps a network throw and a timeout to NangoUnreachableError', async () => {
    const network = createSource({}, async () => {
      throw new Error(`connect ECONNREFUSED ${SECRET} ${CONNECTION_ID} ${TOKEN} ${LEAK} /connection/${CONNECTION_ID}`);
    });
    const networkErr = await rejectionFrom(() => network.getAccessToken());
    expect(networkErr).toBeInstanceOf(NangoUnreachableError);
    expect(networkErr.code).toBe('NANGO_UNREACHABLE');
    expect(networkErr.message).toBe(`Nango is unreachable for ${ACCOUNT}`);

    const timeout = createSource({ timeoutMs: 20 }, (_url, init) => new Promise((_resolve, reject) => {
      const signal = init?.signal;
      const fail = () => {
        const reason = signal?.reason;
        reject(reason instanceof Error
          ? reason
          : Object.assign(new Error(`timeout ${SECRET} ${CONNECTION_ID} ${TOKEN}`), { name: 'TimeoutError' }));
      };
      if (!signal) {
        fail();
        return;
      }
      if (signal.aborted) fail();
      else signal.addEventListener('abort', fail, { once: true });
    }));
    const timeoutErr = await rejectionFrom(() => timeout.getAccessToken());
    expect(timeoutErr).toBeInstanceOf(NangoUnreachableError);
    expect(timeoutErr.code).toBe('NANGO_UNREACHABLE');
    expect(timeoutErr.message).toBe(`Nango timed out for ${ACCOUNT}`);
  });

  it('maps a 200 without a usable token, and non-JSON, to NangoInvalidResponseError', async () => {
    const missing = createSource({}, async () => poisoned(200, JSON.stringify({
      credentials: { refresh_token: TOKEN },
      connection_id: CONNECTION_ID,
      leak: LEAK,
    })));
    const missingErr = await rejectionFrom(() => missing.getAccessToken());
    expect(missingErr).toBeInstanceOf(NangoInvalidResponseError);
    expect(missingErr.code).toBe('NANGO_INVALID_RESPONSE');
    expect(missingErr.message).toContain('(HTTP 200)');

    const garbage = createSource({}, async () => poisoned(200, `not-json ${TOKEN} ${SECRET} ${CONNECTION_ID} ${LEAK}`));
    const garbageErr = await rejectionFrom(() => garbage.getAccessToken());
    expect(garbageErr).toBeInstanceOf(NangoInvalidResponseError);
    expect(garbageErr.code).toBe('NANGO_INVALID_RESPONSE');
  });

  it('passes identity on mail or userPrincipalName, case-insensitively', async () => {
    const byMail = createSource({}, async url => {
      if (url.startsWith('https://graph.microsoft.com/')) {
        return graphIdentity({ mail: 'Joshua@Example.com', userPrincipalName: OTHER_ADDRESS });
      }
      return json(nangoBody());
    });
    await expect(byMail.getAccessToken()).resolves.toBe(TOKEN);

    const byUpn = createSource({}, async url => {
      if (url.startsWith('https://graph.microsoft.com/')) {
        return graphIdentity({ mail: null, userPrincipalName: 'Joshua@Example.COM' });
      }
      return json(nangoBody());
    });
    await expect(byUpn.getAccessToken()).resolves.toBe(TOKEN);
  });

  it('rejects an identity mismatch and checks again on the next call', async () => {
    let graphCalls = 0;
    let nangoCalls = 0;
    const source = createSource({}, async url => {
      if (url.startsWith('https://graph.microsoft.com/')) {
        graphCalls += 1;
        return graphIdentity({ mail: OTHER_ADDRESS, userPrincipalName: OTHER_ADDRESS });
      }
      nangoCalls += 1;
      return json(nangoBody());
    });

    const first = await rejectionFrom(() => source.getAccessToken());
    expect(first).toBeInstanceOf(NangoIdentityMismatchError);
    expect(first.code).toBe('NANGO_IDENTITY_MISMATCH');
    expect(first.message).toBe(`Nango connection does not belong to ${ACCOUNT}`);
    expect(first.message).not.toContain(OTHER_ADDRESS);

    const second = await rejectionFrom(() => source.getAccessToken());
    expect(second).toBeInstanceOf(NangoIdentityMismatchError);
    expect(graphCalls).toBe(2);
    expect(nangoCalls).toBe(2);
  });

  it('runs the identity check once per instance after it passes', async () => {
    let graphCalls = 0;
    let nangoCalls = 0;
    const expiresAt = new Date(clock.now + 10 * 60 * 1000).toISOString();
    const source = createSource({}, async url => {
      if (url.startsWith('https://graph.microsoft.com/')) {
        graphCalls += 1;
        return graphIdentity({ userPrincipalName: 'Joshua@Example.com' });
      }
      nangoCalls += 1;
      return json(nangoBody(expiresAt));
    });

    await source.getAccessToken();
    await source.getAccessToken();
    clock.now += 5 * 60 * 1000;
    await source.getAccessToken();
    expect(nangoCalls).toBe(2);
    expect(graphCalls).toBe(1);
  });

  it('maps a non-2xx Graph identity response to NangoIdentityCheckError', async () => {
    const source = createSource({}, async url => {
      if (url.startsWith('https://graph.microsoft.com/')) return poisoned(401);
      return json(nangoBody());
    });
    const err = await rejectionFrom(() => source.getAccessToken());
    expect(err).toBeInstanceOf(NangoIdentityCheckError);
    expect(err.code).toBe('NANGO_IDENTITY_CHECK_FAILED');
    expect(err.message).toBe(`Microsoft Graph identity check failed for ${ACCOUNT} (HTTP 401)`);
  });

  it('throws NangoNotConfiguredError before any request when the secret is missing', async () => {
    const fetchImpl = vi.fn();
    const source = createSource({ secretKey: '' }, fetchImpl);
    const err = await rejectionFrom(() => source.getAccessToken());
    expect(err).toBeInstanceOf(NangoNotConfiguredError);
    expect(err.code).toBe('NANGO_NOT_CONFIGURED');
    expect(fetchImpl).not.toHaveBeenCalled();
    await expect(source.forceRefresh()).resolves.toBe(false);
    expect(fetchImpl).not.toHaveBeenCalled();
  });

  it('forceRefresh sends force_refresh=true and returns false instead of throwing', async () => {
    const urls: string[] = [];
    const success = createSource({}, async url => {
      urls.push(url);
      if (url.startsWith('https://graph.microsoft.com/')) return graphIdentity({ mail: ACCOUNT });
      return json(nangoBody());
    });
    await success.getAccessToken();
    const afterCache = urls.length;
    await expect(success.forceRefresh()).resolves.toBe(true);
    const refreshed = urls.slice(afterCache);
    expect(refreshed.some(url => url.includes('force_refresh=true'))).toBe(true);
    expect(refreshed.some(url => url.includes('/proxy'))).toBe(false);

    const failure = createSource({}, async url => {
      urls.push(url);
      if (url.includes('force_refresh=true')) return poisoned(503, `down ${SECRET} ${CONNECTION_ID} ${TOKEN} ${LEAK}`);
      if (url.startsWith('https://graph.microsoft.com/')) return graphIdentity({ mail: ACCOUNT });
      return json(nangoBody());
    });
    const beforeFailure = urls.length;
    await expect(failure.forceRefresh()).resolves.toBe(false);
    const failed = urls.slice(beforeFailure);
    expect(failed.some(url => url.includes('force_refresh=true'))).toBe(true);
    expect(failed.some(url => url.includes('/proxy'))).toBe(false);
  });

  it('never requests a URL containing /proxy', async () => {
    const urls: string[] = [];
    const source = createSource({}, async url => {
      urls.push(url);
      if (url.includes('/proxy')) return poisoned(500);
      if (url.startsWith('https://graph.microsoft.com/')) return graphIdentity({ mail: ACCOUNT });
      return json(nangoBody());
    });
    await source.getAccessToken();
    await source.forceRefresh();
    expect(urls.length).toBeGreaterThan(0);
    expect(urls.some(url => url.includes('/proxy'))).toBe(false);
  });
});
