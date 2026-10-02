/**
 * The Sheets `:batchUpdate` call, in one place.
 *
 * Google applies a batch all-or-nothing — "if any request is not valid, no
 * requests will be applied" — so grouping requests here is also how a tool
 * gets atomicity across more than one mutation.
 *
 * The chart tools are the first callers that need `replies`: an added chart's
 * generated id arrives there and nowhere else.
 */

import { makeSheetsRequest } from './google.js';

export interface BatchUpdateOptions {
  /** Masks the response, not the request — the default reply echoes a lot. */
  fields?: string;
}

export async function applyBatchUpdate<R = unknown>(
  spreadsheetId: string,
  requests: object[],
  accessToken: string,
  options: BatchUpdateOptions = {},
): Promise<R[]> {
  if (requests.length === 0) {
    throw new Error('batchUpdate requires at least one request');
  }

  const query = options.fields
    ? `?fields=${encodeURIComponent(options.fields)}`
    : '';

  const result = await makeSheetsRequest(
    `/${encodeURIComponent(spreadsheetId)}:batchUpdate${query}`,
    accessToken,
    { method: 'POST', body: JSON.stringify({ requests }) },
  ) as { replies?: (R | null)[] };

  // Google returns an empty object for requests that carry no reply, so the
  // array is positional and its holes are meaningful. Callers index by the
  // position of the request they sent.
  return (result.replies ?? []) as R[];
}
