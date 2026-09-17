import { describe, it, expect, beforeEach, vi } from 'vitest';
import { writeFile, mkdir } from 'node:fs/promises';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import { MockEmailProvider } from '../testing/mock-provider.js';
import { createForwardDraftAction } from './forward.js';
import type { ActionContext } from './registry.js';

const VALID_MSG_ID = 'abc123def456ghi789jkl012';

let provider: MockEmailProvider;
let ctx: ActionContext;
let testDir: string;

beforeEach(async () => {
  provider = new MockEmailProvider();
  provider.addMessage({
    id: VALID_MSG_ID,
    subject: 'Quarterly Report',
    from: { email: 'alice@corp.com', name: 'Alice' },
    to: [{ email: 'me@company.com' }],
    threadId: 'thread-source',
    conversationId: 'thread-source',
    receivedAt: '2026-04-25T16:08:46Z',
    isRead: true,
    hasAttachments: false,
    body: 'Original body',
  });
  testDir = join(tmpdir(), `forward-test-${Date.now()}`);
  await mkdir(testDir, { recursive: true });
  ctx = {
    provider,
    mailboxName: 'work',
    allMailboxes: [
      { name: 'work', emailAddress: 'me@company.com', provider, providerType: 'microsoft', isDefault: true, status: 'connected' },
    ],
    sendAllowlist: { entries: ['*@blocked.com'] },
    safeDir: testDir,
  };
});

describe('email-write/Create Forward Draft', () => {
  it('Scenario: Create forward draft of allowed message', async () => {
    const result = await createForwardDraftAction.run(ctx, {
      message_id: VALID_MSG_ID,
      to: ['bob@example.com'],
      comment: 'Please review',
    });

    expect(result.success).toBe(true);
    expect(result.draftId).toBeDefined();
    const draft = provider.getDrafts().get(result.draftId!)!;
    expect(draft.to).toEqual([{ email: 'bob@example.com' }]);
    expect(draft.subject).toBe('Fwd: Quarterly Report');
    expect(draft.threadId).toBe('thread-source');
    expect(draft.body).toContain('Please review');
    expect(provider.getSentMessages()).toHaveLength(0);
  });

  it('Scenario: Forward drafts bypass the send allowlist', async () => {
    const result = await createForwardDraftAction.run(ctx, {
      message_id: VALID_MSG_ID,
      to: ['outsider@not-allowed.com'],
      comment: 'fyi',
    });

    expect(result.success).toBe(true);
    expect(result.draftId).toBeDefined();
    expect(provider.getSentMessages()).toHaveLength(0);
  });

  it('Scenario: Mailbox required with multiple accounts', async () => {
    const multi: ActionContext = {
      ...ctx,
      allMailboxes: [
        { name: 'work', emailAddress: 'me@company.com', provider, providerType: 'microsoft', isDefault: true, status: 'connected' },
        { name: 'personal', emailAddress: 'me@home.com', provider, providerType: 'microsoft', isDefault: false, status: 'connected' },
      ],
    };
    const result = await createForwardDraftAction.run(multi, {
      message_id: VALID_MSG_ID,
      to: ['bob@example.com'],
    });
    expect(result.success).toBe(false);
    expect(result.error?.code).toBe('MAILBOX_REQUIRED');
  });

  it('rejects an implausible message_id before any provider call', async () => {
    const spy = vi.spyOn(provider, 'createForwardDraft');
    const result = await createForwardDraftAction.run(ctx, {
      message_id: 'short',
      to: ['bob@example.com'],
    });
    expect(result.success).toBe(false);
    expect(result.error?.code).toBe('INVALID_MESSAGE_ID');
    expect(spy).not.toHaveBeenCalled();
  });

  it('returns NOT_SUPPORTED when the provider cannot create forward drafts', async () => {
    (provider as { createForwardDraft?: unknown }).createForwardDraft = undefined;
    const result = await createForwardDraftAction.run(ctx, {
      message_id: VALID_MSG_ID,
      to: ['bob@example.com'],
    });
    expect(result.success).toBe(false);
    expect(result.error?.code).toBe('NOT_SUPPORTED');
  });

  it('returns INVALID_ADDRESS for a malformed to recipient', async () => {
    const result = await createForwardDraftAction.run(ctx, {
      message_id: VALID_MSG_ID,
      to: ['not an email'],
    });
    expect(result.success).toBe(false);
    expect(result.error?.code).toBe('INVALID_ADDRESS');
    expect(result.error?.message).toContain('to[0]');
  });

  it('parses name-address strings on to and cc', async () => {
    const result = await createForwardDraftAction.run(ctx, {
      message_id: VALID_MSG_ID,
      to: ['Bob Smith <bob@example.com>'],
      cc: ['"Carol Jones" <carol@example.com>'],
      comment: 'heads up',
    });
    expect(result.success).toBe(true);
    const draft = provider.getDrafts().get(result.draftId!)!;
    expect(draft.to).toEqual([{ name: 'Bob Smith', email: 'bob@example.com' }]);
    expect(draft.cc).toEqual([{ name: 'Carol Jones', email: 'carol@example.com' }]);
  });

  it('carries caller attachments on the forward draft', async () => {
    const pdf = Buffer.from('%PDF-1.4\n1 0 obj<<>>endobj\n', 'utf-8');
    await writeFile(join(testDir, 'note.pdf'), pdf);
    const result = await createForwardDraftAction.run(ctx, {
      message_id: VALID_MSG_ID,
      to: ['bob@example.com'],
      attachments: [{ path: 'note.pdf' }],
    });
    expect(result.success).toBe(true);
    const draft = provider.getDrafts().get(result.draftId!)!;
    expect(draft.attachments).toHaveLength(1);
    expect(draft.attachments![0]!.filename).toBe('note.pdf');
  });

  it('does not send the forward', async () => {
    const sendSpy = vi.spyOn(provider, 'sendMessage');
    const sendDraftSpy = vi.spyOn(provider, 'sendDraft');
    const result = await createForwardDraftAction.run(ctx, {
      message_id: VALID_MSG_ID,
      to: ['bob@example.com'],
      comment: 'fyi',
    });
    expect(result.success).toBe(true);
    expect(sendSpy).not.toHaveBeenCalled();
    expect(sendDraftSpy).not.toHaveBeenCalled();
    expect(provider.getSentMessages()).toHaveLength(0);
  });
});
