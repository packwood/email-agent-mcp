/**
 * Shared HTTP 429 throttling policy for the Microsoft transports.
 *
 * Lives in its own module because both transports need it and `maton.ts` already
 * imports from `email-graph-provider.ts` — defining it in either would make the
 * import cycle.
 */

/**
 * Reads are retried on HTTP 429; writes are not. A throttled write may or may not
 * have been applied, and re-sending mail or re-patching a message is worse than
 * surfacing the error. This mirrors the Gmail transport's read-only retry policy.
 */
export const MAX_READ_ATTEMPTS = 4;

/**
 * How long to wait before retrying a throttled request. Prefers the server's own
 * `Retry-After` (seconds, or an HTTP date), capped at 30s so a mistaken or hostile
 * header cannot stall a briefing indefinitely. Falls back to jittered exponential
 * backoff, so concurrent callers do not resynchronise onto the same retry instant.
 */
/**
 * Read an error body without ever throwing. A response body can only be consumed
 * once, and the retry path cancels bodies it is discarding, so a failed read here
 * must not replace a meaningful HTTP status with an unrelated "Body is unusable"
 * TypeError. The status is the diagnostic that matters; the body is a bonus.
 */
export async function safeErrorText(response: Response): Promise<string> {
  try {
    return await response.text();
  } catch {
    return `<unreadable ${response.status} response body>`;
  }
}

export function readRetryDelayMs(response: Response, attempt: number): number {
  const retryAfter = response.headers.get('retry-after');
  if (retryAfter) {
    const seconds = Number(retryAfter);
    if (Number.isFinite(seconds) && seconds >= 0) return Math.min(30_000, seconds * 1000);
    const date = Date.parse(retryAfter);
    if (Number.isFinite(date)) return Math.min(30_000, Math.max(0, date - Date.now()));
  }
  const cap = Math.min(8_000, 500 * 2 ** attempt);
  return Math.floor(Math.random() * cap);
}
