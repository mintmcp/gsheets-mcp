/**
 * Error types and structured tool response helpers.
 */

export type GoogleApi = 'drive' | 'sheets';

/**
 * Typed API error carrying HTTP status and optional Retry-After hint.
 * Allows `wrapHandler` to surface 429 / 4xx-specific guidance to the LLM.
 */
export class ApiError extends Error {
  status: number;
  retryAfterSeconds?: number;
  api: GoogleApi;
  /** Google's `error.status` enum (e.g. FAILED_PRECONDITION), when present. */
  reason?: string;
  constructor(
    message: string,
    status: number,
    api: GoogleApi,
    retryAfterSeconds?: number,
    reason?: string,
  ) {
    super(message);
    this.name = 'ApiError';
    this.status = status;
    this.api = api;
    this.retryAfterSeconds = retryAfterSeconds;
    this.reason = reason;
  }
}

export function parseRetryAfter(header: string | null): number | undefined {
  if (!header) return undefined;
  const asInt = parseInt(header, 10);
  if (!isNaN(asInt) && asInt >= 0) return asInt;
  const asDate = Date.parse(header);
  if (!isNaN(asDate)) {
    return Math.max(0, Math.round((asDate - Date.now()) / 1000));
  }
  return undefined;
}

export function toolResponse<T>(structuredContent: T, notice?: string) {
  const json = JSON.stringify(structuredContent);
  const text = notice ? `${notice}\n${json}` : json;
  return {
    content: [{ type: 'text' as const, text }],
    structuredContent,
  };
}

export function toolError(message: string, extra?: Record<string, unknown>) {
  const payload = { error: message, ...(extra || {}) };
  return {
    content: [{ type: 'text' as const, text: JSON.stringify(payload) }],
    isError: true,
  };
}

export type ToolErrorResult = ReturnType<typeof toolError>;

/**
 * The wrapped handler can always return an error envelope, so its type is the
 * handler's own result OR that envelope. Claiming it returns `H` unchanged
 * made a handler that only throws infer `Promise<never>`, which is a lie the
 * catch block disproves.
 */
export function wrapHandler<A extends any[], R>(
  handler: (...args: A) => Promise<R>,
): (...args: A) => Promise<R | ToolErrorResult> {
  return async (...args: A) => {
    try {
      return await handler(...args);
    } catch (err: any) {
      if (err instanceof ApiError) {
        const extra: Record<string, unknown> = { status: err.status, api: err.api };
        if (err.status === 429) {
          extra.code = 'rate_limited';
          if (err.retryAfterSeconds !== undefined) {
            extra.retryAfterSeconds = err.retryAfterSeconds;
          }
          extra.hint = 'Rate limit exceeded. Back off and retry after the indicated delay.';
        } else if (err.status === 400) {
          extra.code = 'invalid_argument';
          extra.hint = 'Check that ranges, sheet names, IDs, and value shapes are valid.';
        } else if (err.status === 401) {
          extra.code = 'unauthenticated';
        } else if (err.status === 403) {
          extra.code = 'permission_denied';
        } else if (err.status === 404) {
          extra.code = 'not_found';
        } else if (err.status >= 500) {
          extra.code = 'server_error';
          extra.hint = 'Transient upstream error. Safe to retry after a brief delay.';
        }
        return toolError(err.message, extra);
      }
      const msg = err?.message ? String(err.message) : String(err);
      return toolError(msg);
    }
  };
}
