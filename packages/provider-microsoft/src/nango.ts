// Nango holds and refreshes the Microsoft grant. Mail itself still goes to
// graph.microsoft.com. Never call Nango's /proxy endpoint, and never put the
// access token, connection id, secret key, or a Nango response body into an
// error or a log — those bodies can echo the request.

const DEFAULT_NANGO_HOST = 'https://api.nango.dev';
const EXPIRY_SKEW_MS = 5 * 60 * 1000;
const FALLBACK_TTL_MS = 60 * 1000;
const DEFAULT_TIMEOUT_MS = 15_000;
const GRAPH_IDENTITY_URL = 'https://graph.microsoft.com/v1.0/me?$select=mail,userPrincipalName';

const PROVIDER_CONFIG_KEY_PATTERN = /^[a-z0-9][a-z0-9._-]{0,127}$/;
const CONNECTION_FIELDS = ['account', 'providerConfigKey', 'connectionId'] as const;

/**
 * Nango connection ids are opaque. '/' '%' '=' and spaces are valid; the id is
 * encodeURIComponent-encoded in the request path. Reject control characters
 * and whitespace-only values. Never echo the id in an error.
 */
function isAcceptedConnectionId(connectionId: string): boolean {
  if (connectionId.length < 1 || connectionId.length > 255) return false;
  if (/[\u0000-\u001F\u007F]/.test(connectionId)) return false;
  return connectionId.trim() !== '';
}

export interface NangoOutlookConnection {
  account: string;
  providerConfigKey: string;
  connectionId: string;
}

export type NangoFetch = (url: string, init?: RequestInit) => Promise<Response>;

export interface NangoTokenSourceOptions {
  account: string;
  providerConfigKey: string;
  connectionId: string;
  secretKey: string;
  host: string;
  fetchImpl?: NangoFetch;
  now?: () => number;
  timeoutMs?: number;
}

export class NangoTransportError extends Error {
  readonly code: string;
  readonly account: string;

  constructor(code: string, account: string, message: string) {
    super(message);
    this.name = new.target.name;
    this.code = code;
    this.account = account;
  }
}

export class NangoNotConfiguredError extends NangoTransportError {
  constructor(account: string) {
    super('NANGO_NOT_CONFIGURED', account, `Nango is not configured for ${account} (missing secret key)`);
  }
}

export class NangoRefusedError extends NangoTransportError {
  constructor(account: string, status: number) {
    super('NANGO_REFUSED', account, `Nango refused the token request for ${account} (HTTP ${status})`);
  }
}

export class NangoNotFoundError extends NangoTransportError {
  constructor(account: string, status: number) {
    super(
      'NANGO_CONNECTION_NOT_FOUND',
      account,
      `Nango connection was not found for ${account} (HTTP ${status})`,
    );
  }
}

export class NangoGrantExpiredError extends NangoTransportError {
  constructor(account: string) {
    super(
      'NANGO_GRANT_EXPIRED',
      account,
      `Nango could not refresh the Microsoft grant for ${account} (HTTP 424); reconnect this mailbox in Nango`,
    );
  }
}

export class NangoUnreachableError extends NangoTransportError {
  constructor(account: string, status?: number, timedOut = false) {
    const statusText = status === undefined ? '' : ` (HTTP ${status})`;
    const what = timedOut ? 'timed out' : 'is unreachable';
    super('NANGO_UNREACHABLE', account, `Nango ${what} for ${account}${statusText}`);
  }
}

export class NangoInvalidResponseError extends NangoTransportError {
  constructor(account: string, status: number) {
    super(
      'NANGO_INVALID_RESPONSE',
      account,
      `Nango returned an invalid response for ${account} (HTTP ${status})`,
    );
  }
}

export class NangoIdentityMismatchError extends NangoTransportError {
  constructor(account: string) {
    super(
      'NANGO_IDENTITY_MISMATCH',
      account,
      `Nango connection does not belong to ${account}`,
    );
  }
}

export class NangoIdentityCheckError extends NangoTransportError {
  /** Graph returned a response whose body could not be read or parsed. */
  static unreadable(account: string, status: number): NangoIdentityCheckError {
    return new NangoIdentityCheckError(account, status, false, true);
  }

  constructor(account: string, status?: number, timedOut = false, unreadable = false) {
    // No HTTP status on a network failure or timeout: there was no Graph response.
    // A 2xx body that cannot be read or parsed still came from Graph, not Nango.
    const message = unreadable
      ? `Microsoft Graph identity check returned an unreadable response for ${account} (HTTP ${status})`
      : status === undefined
        ? timedOut
          ? `Microsoft Graph identity check timed out for ${account}`
          : `Microsoft Graph identity check could not reach Microsoft Graph for ${account}`
        : `Microsoft Graph identity check failed for ${account} (HTTP ${status})`;
    super('NANGO_IDENTITY_CHECK_FAILED', account, message);
  }
}

function invalidConfig(): never {
  throw new Error('Invalid Nango Outlook connections configuration');
}

function normalizeAccount(value: string): string {
  return value.trim().toLowerCase();
}

function readConnection(item: unknown): NangoOutlookConnection {
  if (item === null || typeof item !== 'object' || Array.isArray(item)) invalidConfig();
  const record = item as Record<string, unknown>;
  const keys = Object.keys(record);
  if (
    keys.length !== CONNECTION_FIELDS.length ||
    CONNECTION_FIELDS.some(field => !Object.hasOwn(record, field))
  ) {
    invalidConfig();
  }
  const { account, providerConfigKey, connectionId } = record;
  if (typeof account !== 'string' || typeof providerConfigKey !== 'string' || typeof connectionId !== 'string') {
    invalidConfig();
  }
  const normalised = normalizeAccount(account);
  if (!normalised.includes('@')) invalidConfig();
  if (!PROVIDER_CONFIG_KEY_PATTERN.test(providerConfigKey)) invalidConfig();
  if (!isAcceptedConnectionId(connectionId)) invalidConfig();
  return { account: normalised, providerConfigKey, connectionId };
}

/**
 * Parse `EMAIL_AGENT_MCP_NANGO_OUTLOOK_CONNECTIONS`. Empty or omitted input is
 * no Nango mailboxes. Any invalid document throws, and the message names neither
 * a connection id nor the raw JSON.
 */
export function parseNangoOutlookConnections(raw: string | undefined): Map<string, NangoOutlookConnection> {
  if (raw === undefined || raw === '') return new Map();
  let parsed: unknown;
  try {
    parsed = JSON.parse(raw);
  } catch {
    invalidConfig();
  }
  if (!Array.isArray(parsed)) invalidConfig();

  const result = new Map<string, NangoOutlookConnection>();
  for (const item of parsed) {
    const connection = readConnection(item);
    if (result.has(connection.account)) {
      throw new Error(`Duplicate Nango Outlook connection for ${connection.account}`);
    }
    result.set(connection.account, connection);
  }
  return result;
}

/** Validated Nango origin with no trailing slash. Omitted input uses the public API. */
export function resolveNangoHost(raw: string | undefined): string {
  const trimmed = raw?.trim() ?? '';
  if (trimmed === '') return DEFAULT_NANGO_HOST;
  let url: URL;
  try {
    url = new URL(trimmed);
  } catch {
    throw new Error('Invalid Nango host');
  }
  if (
    url.protocol !== 'https:' ||
    url.username !== '' ||
    url.password !== '' ||
    url.pathname !== '/' ||
    url.search !== '' ||
    url.hash !== ''
  ) {
    throw new Error('Invalid Nango host');
  }
  return url.origin;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === 'object' && !Array.isArray(value);
}

function readAccessToken(body: unknown): string | undefined {
  if (!isRecord(body) || !isRecord(body.credentials)) return undefined;
  const token = body.credentials.access_token;
  if (typeof token !== 'string' || token.trim() === '') return undefined;
  return token;
}

function isAbortFailure(err: unknown): boolean {
  if (!err || typeof err !== 'object') return false;
  const name = (err as { name?: unknown }).name;
  return name === 'TimeoutError' || name === 'AbortError';
}

function addressMatches(value: unknown, account: string): boolean {
  return typeof value === 'string' && value.trim().toLowerCase() === account.trim().toLowerCase();
}

interface CachedToken {
  token: string;
  until: number;
}

export class NangoTokenSource {
  private readonly account: string;
  private readonly providerConfigKey: string;
  private readonly connectionId: string;
  private readonly secretKey: string;
  private readonly host: string;
  private readonly fetchImpl: NangoFetch;
  private readonly now: () => number;
  private readonly timeoutMs: number;

  private cached: CachedToken | null = null;
  private inFlight: Promise<string> | null = null;
  private grantExpiredWarning: string | undefined;
  private generation = 0;
  private requestSerial = 0;

  constructor(options: NangoTokenSourceOptions) {
    this.account = options.account;
    this.providerConfigKey = options.providerConfigKey;
    this.connectionId = options.connectionId;
    this.secretKey = options.secretKey;
    this.host = options.host.replace(/\/+$/, '');
    this.fetchImpl = options.fetchImpl ?? ((url, init) => fetch(url, init));
    this.now = options.now ?? (() => Date.now());
    this.timeoutMs = options.timeoutMs ?? DEFAULT_TIMEOUT_MS;
  }

  async getAccessToken(): Promise<string> {
    const cached = this.freshCachedToken();
    if (cached !== undefined) return cached;
    if (this.inFlight) return this.inFlight;

    let settle!: (token: string) => void;
    let fail!: (err: unknown) => void;
    const flight = new Promise<string>((resolve, reject) => {
      settle = resolve;
      fail = reject;
    });
    this.inFlight = flight;
    void this.obtainToken(false).then(settle, fail).finally(() => {
      if (this.inFlight === flight) this.inFlight = null;
    });
    return flight;
  }

  /** Drop the in-memory token. The next token fetched from Nango is identity-checked again. */
  invalidate(): void {
    this.cached = null;
    this.generation += 1;
  }

  /**
   * Drop the cache and ask Nango to refresh the grant. Returns false on any
   * transport failure. The failure is not logged: the error can carry nothing
   * from the Nango body, and we do not print it either.
   */
  async forceRefresh(): Promise<boolean> {
    this.invalidate();
    try {
      await this.obtainToken(true);
      return true;
    } catch (err) {
      if (err instanceof NangoTransportError) return false;
      throw err;
    }
  }

  /**
   * Drop the cache and ask Nango to refresh after Graph returned 401.
   * Returns false only when Nango itself is unreachable or timed out, so the
   * caller can still surface that 401. Every other transport error is rethrown:
   * Graph did not apply the request, and a dead grant, a missing connection,
   * or a failed identity check must not be hidden behind a generic Graph 401.
   * Nothing from the Nango body is logged.
   */
  async refreshAfterAuthError(): Promise<boolean> {
    this.invalidate();
    try {
      await this.obtainToken(true);
      return true;
    } catch (err) {
      if (err instanceof NangoUnreachableError) return false;
      throw err;
    }
  }

  getTokenHealthWarning(): string | undefined {
    return this.grantExpiredWarning;
  }

  private freshCachedToken(): string | undefined {
    if (!this.cached) return undefined;
    if (this.now() < this.cached.until) return this.cached.token;
    this.cached = null;
    return undefined;
  }

  private async obtainToken(force: boolean): Promise<string> {
    const serial = ++this.requestSerial;
    const generation = this.generation;
    try {
      if (this.secretKey === '') throw new NangoNotConfiguredError(this.account);
      const fetched = await this.fetchNangoToken(force);
      await this.ensureIdentity(fetched.token);
      if (
        serial === this.requestSerial &&
        generation === this.generation &&
        fetched.cacheUntil !== null
      ) {
        this.cached = { token: fetched.token, until: fetched.cacheUntil };
      }
      if (serial === this.requestSerial) this.grantExpiredWarning = undefined;
      return fetched.token;
    } catch (err) {
      if (serial === this.requestSerial) {
        this.grantExpiredWarning = err instanceof NangoGrantExpiredError ? err.message : undefined;
      }
      throw err;
    }
  }

  private connectionUrl(force: boolean): string {
    const id = encodeURIComponent(this.connectionId);
    const key = encodeURIComponent(this.providerConfigKey);
    const refresh = force ? '&force_refresh=true' : '';
    return `${this.host}/connection/${id}?provider_config_key=${key}${refresh}`;
  }

  private async fetchNangoToken(force: boolean): Promise<{ token: string; cacheUntil: number | null }> {
    const response = await this.request(this.connectionUrl(force), `Bearer ${this.secretKey}`, 'nango');
    if (!response.ok || response.status !== 200) {
      await discardBody(response);
      throw this.statusError(response.status);
    }

    const body = await this.readJson(response);
    const token = readAccessToken(body);
    if (!token) throw new NangoInvalidResponseError(this.account, response.status);
    const expiresAt = isRecord(body) && isRecord(body.credentials) ? body.credentials.expires_at : undefined;
    return { token, cacheUntil: this.cacheUntilFor(expiresAt) };
  }

  private statusError(status: number): NangoTransportError {
    if (status === 424) return new NangoGrantExpiredError(this.account);
    if (status === 404) return new NangoNotFoundError(this.account, status);
    if (status >= 500) return new NangoUnreachableError(this.account, status);
    if (status >= 400 && status < 500) return new NangoRefusedError(this.account, status);
    return new NangoInvalidResponseError(this.account, status);
  }

  private cacheUntilFor(expiresAt: unknown): number | null {
    const now = this.now();
    if (typeof expiresAt === 'string') {
      const parsed = Date.parse(expiresAt);
      if (Number.isFinite(parsed)) {
        const until = parsed - EXPIRY_SKEW_MS;
        // Inside the 5-minute window the token is still usable once, but the
        // next call must refetch rather than serve a near-expiry cache hit.
        return until > now ? until : null;
      }
    }
    return now + FALLBACK_TTL_MS;
  }

  /**
   * Confirm a token just fetched from Nango, before it is cached or returned.
   * A connection can be re-authorised to a different Microsoft account between
   * fetches, so a check that passed for an earlier token does not cover this
   * one. Cache hits never reach here.
   */
  private async ensureIdentity(token: string): Promise<void> {
    const response = await this.request(GRAPH_IDENTITY_URL, `Bearer ${token}`, 'identity');
    if (!response.ok) {
      await discardBody(response);
      throw new NangoIdentityCheckError(this.account, response.status);
    }
    const body = await this.readJson(response, 'identity');
    const mail = isRecord(body) ? body.mail : undefined;
    const userPrincipalName = isRecord(body) ? body.userPrincipalName : undefined;
    if (!addressMatches(mail, this.account) && !addressMatches(userPrincipalName, this.account)) {
      throw new NangoIdentityMismatchError(this.account);
    }
  }

  private async request(
    url: string,
    authorization: string,
    failure: 'nango' | 'identity',
  ): Promise<Response> {
    try {
      return await this.fetchImpl(url, {
        method: 'GET',
        headers: { Authorization: authorization },
        signal: AbortSignal.timeout(this.timeoutMs),
        redirect: 'error',
      });
    } catch (err) {
      // Replace the client error. Its message often includes the request URL
      // (and therefore the connection id) or, on a timeout, nothing we need.
      // Blame the service this call was talking to: a Graph /me failure is not
      // "Nango is unreachable".
      const timedOut = isAbortFailure(err);
      if (failure === 'identity') {
        throw new NangoIdentityCheckError(this.account, undefined, timedOut);
      }
      throw new NangoUnreachableError(this.account, undefined, timedOut);
    }
  }

  private async readJson(response: Response, failure: 'nango' | 'identity' = 'nango'): Promise<unknown> {
    const invalid = (): NangoTransportError => failure === 'identity'
      ? NangoIdentityCheckError.unreadable(this.account, response.status)
      : new NangoInvalidResponseError(this.account, response.status);
    let text: string;
    try {
      text = await response.text();
    } catch {
      throw invalid();
    }
    try {
      return JSON.parse(text) as unknown;
    } catch {
      throw invalid();
    }
  }
}

async function discardBody(response: Response): Promise<void> {
  try {
    await response.body?.cancel();
  } catch {
    // The status is the diagnostic. The body is never surfaced.
  }
}
