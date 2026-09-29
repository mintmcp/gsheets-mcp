/**
 * A per-tool-call time budget for upstream Google requests.
 *
 * MCP clients abandon a call after 60 seconds by default (the SDK's
 * DEFAULT_REQUEST_TIMEOUT_MSEC), but nothing here stopped at that point: the
 * server kept waiting on Google and the caller got an opaque "Request timed
 * out". Running every tool call inside a shorter deadline turns a stalled
 * upstream into an error the caller can act on while it is still listening.
 */

import { AsyncLocalStorage } from 'node:async_hooks';

/** Leaves ~15s of the client's 60s for this server, MintMCP and the network. */
export const TOOL_DEADLINE_MS = 45_000;

const deadline = new AsyncLocalStorage<AbortSignal>();

/**
 * Run `fn` with a deadline that aborts upstream requests after `ms`. A nested
 * call can shorten an enclosing deadline but never extend it.
 *
 * Built from a plain timer rather than AbortSignal.timeout/any: `any` holds
 * its sources weakly, so a timeout signal nothing else references can be
 * collected before it fires, and the deadline silently never arrives.
 */
export async function runWithDeadline<T>(
  fn: () => Promise<T>,
  ms: number = TOOL_DEADLINE_MS,
): Promise<T> {
  const outer = deadline.getStore();
  const controller = new AbortController();
  const timer = setTimeout(
    () => controller.abort(new DOMException('Tool call deadline reached', 'TimeoutError')),
    ms,
  );
  const followOuter = () => controller.abort(outer?.reason);
  if (outer?.aborted) followOuter();
  else outer?.addEventListener('abort', followOuter, { once: true });

  try {
    return await deadline.run(controller.signal, fn);
  } finally {
    clearTimeout(timer);
    outer?.removeEventListener('abort', followOuter);
  }
}

/** The signal upstream requests should honor, if a deadline is running. */
export function currentDeadline(): AbortSignal | undefined {
  return deadline.getStore();
}
