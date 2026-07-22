import type { EmailMessage } from '@usejunior/email-core';

export const SEARCH_SCOPES = ['all-visible', 'any', 'participants', 'subject', 'body'] as const;
export type SearchScope = typeof SEARCH_SCOPES[number];

export type MatchClassification =
  | 'verified-header'
  | 'verified-visible-body'
  | 'unexplained-provider-hit';

export interface DeterministicSearchInput {
  query: string;
  searchScope: SearchScope;
  verificationTerm?: string;
  receivedAfter?: string;
  receivedBefore?: string;
}

export interface MatchEvidence {
  matchClassification: MatchClassification;
  matchedFields: string[];
}

const DAY_MS = 24 * 60 * 60 * 1000;
const SNIPPET_LIMIT = 500;

function quoteGmailLiteral(query: string): string {
  return `"${query.replace(/\\/g, '\\\\').replace(/"/g, '\\"')}"`;
}

function kqlLiteral(query: string): string {
  return query
    .replace(/[()"']/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function utcDate(value: number): string {
  return new Date(value).toISOString().slice(0, 10);
}

export function buildProviderSearchQuery(
  providerType: 'microsoft' | 'gmail',
  input: DeterministicSearchInput,
): string {
  const query = input.query.trim();
  let effective: string;

  if (providerType === 'gmail') {
    const literal = quoteGmailLiteral(query);
    switch (input.searchScope) {
      case 'all-visible':
        effective = literal;
        break;
      case 'participants':
        effective = `{from:${literal} to:${literal} cc:${literal} bcc:${literal}}`;
        break;
      case 'subject':
        effective = `subject:${literal}`;
        break;
      case 'body':
        // Gmail has no body-only search operator. Fetch phrase candidates and
        // enforce visible-body scope locally before returning them.
        effective = literal;
        break;
      default:
        effective = query;
    }
    if (input.receivedAfter) {
      effective += ` after:${Math.max(0, Math.floor(Date.parse(input.receivedAfter) / 1000) - 1)}`;
    }
    if (input.receivedBefore) {
      effective += ` before:${Math.ceil(Date.parse(input.receivedBefore) / 1000) + 1}`;
    }
    return effective.trim();
  }

  const literal = kqlLiteral(query);
  switch (input.searchScope) {
    case 'all-visible':
      effective = `(participants:(${literal}) OR subject:(${literal}) OR body:(${literal}))`;
      break;
    case 'participants':
      effective = `participants:(${literal})`;
      break;
    case 'subject':
      effective = `subject:(${literal})`;
      break;
    case 'body':
      effective = `body:(${literal})`;
      break;
    default:
      effective = query;
  }

  // Graph KQL date comparisons are day-granular in common Exchange search
  // deployments. Pad by one UTC day, then enforce the exact timestamp locally.
  const dateTerms: string[] = [];
  if (input.receivedAfter) {
    dateTerms.push(`received>=${utcDate(Date.parse(input.receivedAfter) - DAY_MS)}`);
  }
  if (input.receivedBefore) {
    dateTerms.push(`received<=${utcDate(Date.parse(input.receivedBefore) + DAY_MS)}`);
  }
  return dateTerms.length > 0 ? `(${effective}) AND ${dateTerms.join(' AND ')}` : effective;
}

export function canonicalEffectiveQuery(input: DeterministicSearchInput): string {
  const parts = [`scope=${input.searchScope}`, `query=${JSON.stringify(input.query)}`];
  if (input.verificationTerm) parts.push(`verification_term=${JSON.stringify(input.verificationTerm)}`);
  if (input.receivedAfter) parts.push(`received_after=${input.receivedAfter}`);
  if (input.receivedBefore) parts.push(`received_before=${input.receivedBefore}`);
  return parts.join('; ');
}

function normalize(value: string | undefined): string {
  return (value ?? '').normalize('NFKC').toLocaleLowerCase('en-US');
}

function addressText(address: { email: string; name?: string }): string {
  return address.name ? `${address.name} <${address.email}>` : address.email;
}

export function visibleBody(message: EmailMessage): string {
  if (message.body) return message.body;
  if (!message.bodyHtml) return '';
  return message.bodyHtml
    .replace(/<style\b[^>]*>[\s\S]*?<\/style>/gi, ' ')
    .replace(/<script\b[^>]*>[\s\S]*?<\/script>/gi, ' ')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'")
    .replace(/\s+/g, ' ')
    .trim();
}

export function boundedSnippet(message: EmailMessage): string | undefined {
  const text = (message.snippet ?? visibleBody(message)).replace(/\s+/g, ' ').trim();
  if (!text) return undefined;
  return text.length <= SNIPPET_LIMIT ? text : `${text.slice(0, SNIPPET_LIMIT - 1)}…`;
}

export function classifyMatch(message: EmailMessage, query: string): MatchEvidence {
  const needle = normalize(query.trim().replace(/^"|"$/g, ''));
  if (!needle) return { matchClassification: 'unexplained-provider-hit', matchedFields: [] };

  const fields: Array<[string, string]> = [
    ['subject', message.subject],
    ['from', addressText(message.from)],
    ...message.to.map(address => ['to', addressText(address)] as [string, string]),
    ...(message.cc ?? []).map(address => ['cc', addressText(address)] as [string, string]),
    ...(message.bcc ?? []).map(address => ['bcc', addressText(address)] as [string, string]),
  ];
  const matchedFields = [...new Set(fields
    .filter(([, value]) => normalize(value).includes(needle))
    .map(([field]) => field))];
  if (normalize(visibleBody(message)).includes(needle)) matchedFields.push('body');
  if (matchedFields.length > 0) {
    return {
      matchClassification: matchedFields.some(field => field !== 'body')
        ? 'verified-header'
        : 'verified-visible-body',
      matchedFields,
    };
  }
  return { matchClassification: 'unexplained-provider-hit', matchedFields: [] };
}

export function matchesScope(message: EmailMessage, scope: SearchScope, query: string): boolean {
  if (scope === 'any') return true;
  const evidence = classifyMatch(message, query);
  if (scope === 'all-visible') return evidence.matchedFields.length > 0;
  if (scope === 'participants') {
    return evidence.matchedFields.some(field => ['from', 'to', 'cc', 'bcc'].includes(field));
  }
  if (scope === 'subject') return evidence.matchedFields.includes('subject');
  return evidence.matchedFields.includes('body');
}

export function inReceivedWindow(
  message: EmailMessage,
  receivedAfter?: string,
  receivedBefore?: string,
): boolean {
  const received = Date.parse(message.receivedAt);
  if (!Number.isFinite(received)) return false;
  if (receivedAfter && received < Date.parse(receivedAfter)) return false;
  if (receivedBefore && received >= Date.parse(receivedBefore)) return false;
  return true;
}

export function compareSearchMessages(a: EmailMessage, b: EmailMessage): number {
  const timeDifference = Date.parse(b.receivedAt) - Date.parse(a.receivedAt);
  if (timeDifference !== 0) return timeDifference;
  const mailboxDifference = (a.mailbox ?? '').localeCompare(b.mailbox ?? '');
  return mailboxDifference !== 0 ? mailboxDifference : a.id.localeCompare(b.id);
}

export function formatAddress(address: { email: string; name?: string }): string {
  return addressText(address);
}
