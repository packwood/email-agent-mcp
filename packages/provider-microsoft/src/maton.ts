import { readFile } from 'node:fs/promises';
import { GraphApiError, type GraphApiClient } from './email-graph-provider.js';

const GRAPH_ROOT = 'https://graph.microsoft.com/v1.0';
const MATON_ROOT = 'https://api.maton.ai/outlook/v1.0';
const DEFAULT_DEADLINE_MS = 90_000;

interface MatonConnectionRecord {
  app?: unknown;
  account?: unknown;
  connection_id?: unknown;
  status?: unknown;
}

interface MatonConnectionsFile {
  connections?: unknown;
}

export interface MatonOutlookConnection {
  account: string;
  connectionId: string;
}

function normalizeAccount(value: string): string {
  return value.trim().toLowerCase();
}

function assertConnectionId(value: unknown, account: string): string {
  if (
    typeof value !== 'string' ||
    value.length < 1 ||
    value.length > 256 ||
    /[\u0000-\u001f\u007f]/.test(value)
  ) {
    throw new Error(`Invalid Maton Outlook connection id for ${account}`);
  }
  return value;
}

export async function loadMatonOutlookConnections(
  path: string,
  expectedAccounts?: Iterable<string>,
): Promise<Map<string, MatonOutlookConnection>> {
  let parsed: MatonConnectionsFile;
  try {
    parsed = JSON.parse(await readFile(path, 'utf8')) as MatonConnectionsFile;
  } catch (err) {
    throw new Error(
      `Unable to load Maton connections file: ${err instanceof Error ? err.message : String(err)}`,
    );
  }
  if (!Array.isArray(parsed.connections)) {
    throw new Error('Invalid Maton connections file: "connections" must be an array');
  }

  const result = new Map<string, MatonOutlookConnection>();
  const expected = expectedAccounts
    ? new Set([...expectedAccounts].map(normalizeAccount))
    : null;
  for (const raw of parsed.connections as MatonConnectionRecord[]) {
    if (raw.app !== 'outlook' || typeof raw.account !== 'string' || !raw.account.trim()) continue;
    const account = normalizeAccount(raw.account);
    if (expected && !expected.has(account)) continue;
    if (result.has(account)) {
      throw new Error(`Duplicate Maton Outlook connection for ${account}`);
    }
    if (raw.status !== 'ACTIVE') {
      throw new Error(`Maton Outlook connection for ${account} is not ACTIVE`);
    }
    result.set(account, {
      account,
      connectionId: assertConnectionId(raw.connection_id, account),
    });
  }
  return result;
}

export function resolveMatonOutlookConnection(
  connections: Map<string, MatonOutlookConnection>,
  emailAddress: string,
): MatonOutlookConnection {
  const account = normalizeAccount(emailAddress);
  const connection = connections.get(account);
  if (!connection) {
    throw new Error(`No ACTIVE Maton Outlook connection for ${account}`);
  }
  return connection;
}

export function matonGraphUrl(url: string): string {
  const absolute = url.startsWith('/') ? `${GRAPH_ROOT}${url}` : url;
  let parsed: URL;
  try {
    parsed = new URL(absolute);
  } catch {
    throw new Error('Untrusted Microsoft Graph URL');
  }
  if (
    parsed.protocol !== 'https:' ||
    parsed.port !== '' ||
    parsed.username !== '' ||
    parsed.password !== ''
  ) {
    throw new Error('Untrusted Microsoft Graph URL');
  }

  if (parsed.hostname === 'graph.microsoft.com' && parsed.pathname.startsWith('/v1.0/')) {
    return `${MATON_ROOT}${parsed.pathname.slice('/v1.0'.length)}${parsed.search}`;
  }
  if (
    parsed.hostname === 'api.maton.ai' &&
    (parsed.pathname === '/outlook/v1.0' || parsed.pathname.startsWith('/outlook/v1.0/'))
  ) {
    return parsed.toString();
  }
  throw new Error('Untrusted Microsoft Graph URL');
}

export class MatonGraphApiClient implements GraphApiClient {
  constructor(
    private readonly apiKey: string,
    private readonly connectionId: string,
    private readonly deadlineMs = DEFAULT_DEADLINE_MS,
  ) {
    if (!apiKey) throw new Error('MATON_API_KEY is required for Microsoft transport');
    assertConnectionId(connectionId, 'configured mailbox');
    if (!Number.isSafeInteger(deadlineMs) || deadlineMs < 1 || deadlineMs > 300_000) {
      throw new Error('Invalid Maton request deadline');
    }
  }

  private async request(url: string, method: string, body?: unknown): Promise<Response> {
    const headers: Record<string, string> = {
      Authorization: `Bearer ${this.apiKey}`,
      'Maton-Connection': this.connectionId,
    };
    const init: RequestInit = {
      method,
      headers,
      signal: AbortSignal.timeout(this.deadlineMs),
    };
    if (body !== undefined) {
      headers['Content-Type'] = 'application/json';
      init.body = JSON.stringify(body);
    }
    const response = await fetch(matonGraphUrl(url), init);
    if (!response.ok) {
      throw new GraphApiError(response.status, await response.text());
    }
    return response;
  }

  async get(url: string): Promise<{ value?: unknown[]; [key: string]: unknown }> {
    return this.request(url, 'GET').then(
      response => response.json() as Promise<{ value?: unknown[]; [key: string]: unknown }>,
    );
  }

  async post(url: string, body?: unknown): Promise<{ id?: string; [key: string]: unknown }> {
    const response = await this.request(url, 'POST', body);
    if (response.status === 202 || response.status === 204) return {};
    const text = await response.text();
    return text ? JSON.parse(text) as { id?: string; [key: string]: unknown } : {};
  }

  async patch(url: string, body: unknown): Promise<void> {
    await this.request(url, 'PATCH', body);
  }

  async delete(url: string): Promise<void> {
    await this.request(url, 'DELETE');
  }
}
