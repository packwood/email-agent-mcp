import { describe, expect, it } from 'vitest';
import type { EmailMessage } from '@usejunior/email-core';
import {
  boundedSnippet,
  buildProviderSearchQuery,
  canonicalEffectiveQuery,
  classifyMatch,
  compareSearchMessages,
  inReceivedWindow,
  matchesScope,
} from './search-contract.js';

function message(overrides: Partial<EmailMessage> = {}): EmailMessage {
  return {
    id: 'message-1',
    subject: 'Nala Equities discussion',
    from: { email: 'sender@example.com', name: 'Sender' },
    to: [{ email: 'agustin@nalaequities.com', name: 'Agustin' }],
    cc: [],
    bcc: [],
    receivedAt: '2026-07-22T15:00:00.000Z',
    isRead: false,
    hasAttachments: false,
    body: 'Visible body text',
    ...overrides,
  };
}

describe('deterministic search contract', () => {
  it('builds provider-native scope and padded date queries', () => {
    const input = {
      query: 'Nala Equities',
      searchScope: 'participants' as const,
      receivedAfter: '2026-07-15T17:37:00.000Z',
      receivedBefore: '2026-07-22T17:37:00.000Z',
    };
    expect(buildProviderSearchQuery('microsoft', input)).toBe(
      '(participants:(Nala Equities)) AND received>=2026-07-14 AND received<=2026-07-23',
    );
    expect(buildProviderSearchQuery('gmail', input)).toContain(
      '{from:"Nala Equities" to:"Nala Equities" cc:"Nala Equities" bcc:"Nala Equities"}',
    );
    expect(buildProviderSearchQuery('microsoft', {
      query: 'Nala',
      searchScope: 'all-visible',
    })).toBe('(participants:(Nala) OR subject:(Nala) OR body:(Nala))');
    expect(canonicalEffectiveQuery(input)).toBe(
      'scope=participants; query="Nala Equities"; received_after=2026-07-15T17:37:00.000Z; received_before=2026-07-22T17:37:00.000Z',
    );
  });

  it('classifies independently observable header and body evidence', () => {
    expect(classifyMatch(message(), 'Nala')).toEqual({
      matchClassification: 'verified-header',
      matchedFields: ['subject', 'to'],
    });
    expect(classifyMatch(message({ subject: 'Update', to: [], body: 'Discuss Nala tomorrow' }), 'Nala')).toEqual({
      matchClassification: 'verified-visible-body',
      matchedFields: ['body'],
    });
    expect(classifyMatch(message({ body: 'Discuss Nala tomorrow' }), 'Nala')).toEqual({
      matchClassification: 'verified-header',
      matchedFields: ['subject', 'to', 'body'],
    });
    expect(classifyMatch(message({ subject: 'Update', to: [], body: 'No visible term' }), 'Nala')).toEqual({
      matchClassification: 'unexplained-provider-hit',
      matchedFields: [],
    });
  });

  it('enforces scope locally instead of inferring from thread membership', () => {
    const participant = message({ subject: 'Update', body: 'No visible term' });
    expect(matchesScope(participant, 'participants', 'Nala')).toBe(true);
    expect(matchesScope(participant, 'subject', 'Nala')).toBe(false);
    expect(matchesScope(participant, 'body', 'Nala')).toBe(false);
    expect(matchesScope(participant, 'all-visible', 'Nala')).toBe(true);
    expect(matchesScope(participant, 'any', 'Nala')).toBe(true);
  });

  it('uses an inclusive lower and exclusive upper received-time window', () => {
    const atLower = message({ receivedAt: '2026-07-15T17:37:00.000Z' });
    const atUpper = message({ receivedAt: '2026-07-22T17:37:00.000Z' });
    expect(inReceivedWindow(atLower, '2026-07-15T17:37:00.000Z', '2026-07-22T17:37:00.000Z')).toBe(true);
    expect(inReceivedWindow(atUpper, '2026-07-15T17:37:00.000Z', '2026-07-22T17:37:00.000Z')).toBe(false);
  });

  it('orders equal timestamps by stable message ID', () => {
    const messages = [message({ id: 'b' }), message({ id: 'a' })];
    expect(messages.sort(compareSearchMessages).map(item => item.id)).toEqual(['a', 'b']);
  });

  it('uses mailbox before ID as the global tie-breaker', () => {
    const messages = [
      message({ id: 'a', mailbox: 'z@example.com' }),
      message({ id: 'z', mailbox: 'a@example.com' }),
    ];
    expect(messages.sort(compareSearchMessages).map(item => item.mailbox)).toEqual([
      'a@example.com',
      'z@example.com',
    ]);
  });

  it('bounds visible snippets and strips HTML-only markup', () => {
    const html = message({ body: undefined, bodyHtml: `<p>${'Nala '.repeat(150)}</p><script>hidden</script>` });
    const snippet = boundedSnippet(html)!;
    expect(snippet.length).toBeLessThanOrEqual(500);
    expect(snippet).not.toContain('<p>');
    expect(snippet).not.toContain('hidden');
  });
});
