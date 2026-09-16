// create_forward_draft — compose a forward draft in the source thread, never send
import { z } from 'zod';
import type { EmailAction } from './registry.js';
import { isPlausibleMessageId } from '../security/reply-validation.js';
import { renderEmailBody } from '../content/body-renderer.js';
import {
  checkMailboxRequired,
  handleProviderError,
  parseRecipients,
  buildDraftPreview,
  resolveAttachments,
  AttachmentInputSchema,
  DraftPreviewSchema,
  PreviewErrorSchema,
} from './compose-helpers.js';

const CreateForwardDraftInput = z.object({
  mailbox: z.string().optional(),
  message_id: z.string(),
  to: z.array(z.string()).min(1)
    .describe('Forward recipients. Parsed with the same name-address grammar as create_draft to/cc.'),
  cc: z.array(z.string()).optional(),
  comment: z.string().optional()
    .describe('Optional comment inserted above the quoted original. Empty or omitted still creates a forward draft.'),
  format: z.enum(['markdown', 'html', 'text']).optional()
    .describe("Comment format. 'markdown' (default) renders via GFM; 'html' is passthrough; 'text' is plain text."),
  force_black: z.boolean().optional()
    .describe('Wrap rendered HTML in a force-black div so Outlook dark mode does not hide the text. Default true.'),
  attachments: z.array(AttachmentInputSchema).optional()
    .describe('Additional files to attach alongside the original message attachments.'),
  include_quoted: z.boolean().optional().default(false)
    .describe('Include provider-assembled quoted history in preview.bodyHtml. This affects only the preview, never the stored or sent body.'),
});

const CreateForwardDraftOutput = z.object({
  success: z.boolean(),
  draftId: z.string().optional(),
  preview: DraftPreviewSchema.optional(),
  previewError: PreviewErrorSchema.optional(),
  error: z.object({
    code: z.string(),
    message: z.string(),
    recoverable: z.boolean(),
    availableMailboxes: z.array(z.string()).optional(),
    defaultMailbox: z.string().optional(),
  }).optional(),
});

export const createForwardDraftAction: EmailAction<
  z.infer<typeof CreateForwardDraftInput>,
  z.infer<typeof CreateForwardDraftOutput>
> = {
  name: 'create_forward_draft',
  description: 'Create a forward draft of an existing message. Preserves the source thread. Never sends — use send_draft to send.',
  input: CreateForwardDraftInput,
  output: CreateForwardDraftOutput,
  annotations: { readOnlyHint: false, destructiveHint: false },
  run: async (ctx, input) => {
    const mailboxError = checkMailboxRequired(input.mailbox, ctx.allMailboxes);
    if (mailboxError) {
      return { success: false, error: mailboxError };
    }

    if (!isPlausibleMessageId(input.message_id)) {
      return {
        success: false,
        error: {
          code: 'INVALID_MESSAGE_ID',
          message: 'message_id does not appear to be a valid provider message ID',
          recoverable: false,
        },
      };
    }

    if (!ctx.provider.createForwardDraft) {
      return {
        success: false,
        error: {
          code: 'NOT_SUPPORTED',
          message: 'Forward drafts are not supported by this email provider',
          recoverable: false,
        },
      };
    }

    const parsed = parseRecipients({ to: input.to, cc: input.cc });
    if ('error' in parsed) {
      return { success: false, error: parsed.error };
    }

    const attResult = await resolveAttachments(input.attachments, ctx.safeDir);
    if (attResult.error) {
      return { success: false, error: attResult.error };
    }
    const attachments = attResult.files!.length > 0 ? attResult.files : undefined;

    const comment = input.comment ?? '';
    const rendered = renderEmailBody(comment, { format: input.format, forceBlack: input.force_black });

    try {
      const result = await ctx.provider.createForwardDraft(input.message_id, {
        to: parsed.to,
        cc: parsed.cc.length > 0 ? parsed.cc : undefined,
        comment: rendered.body,
        bodyHtml: rendered.bodyHtml,
        attachments,
      });
      const previewResult = result.success && result.draftId
        ? await buildDraftPreview(ctx.provider, result.draftId, {
          authoredOnly: !input.include_quoted,
        })
        : {};
      return {
        success: result.success,
        draftId: result.draftId,
        ...previewResult,
        error: result.error ? {
          code: result.error.code,
          message: result.error.message,
          recoverable: result.error.recoverable,
        } : undefined,
      };
    } catch (err) {
      return handleProviderError(err, 'DRAFT_FAILED');
    }
  },
};
