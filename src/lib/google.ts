/**
 * Low-level Google Drive / Sheets HTTP helpers.
 */

import { ApiError, parseRetryAfter, type GoogleApi } from './errors.js';

export const GOOGLE_DRIVE_API = 'https://www.googleapis.com/drive/v3';
const GOOGLE_SHEETS_API = 'https://sheets.googleapis.com/v4/spreadsheets';

/** Ceiling on any single upstream body. Bounded reads stay far below this. */
export const MAX_RESPONSE_BYTES = 25 * 1024 * 1024;

/** Error bodies are small; read enough to explain the failure, never more. */
export const MAX_ERROR_BYTES = 64 * 1024;

export interface StreamResult {
  bytes: Uint8Array;
  /** True when the body exceeded maxBytes, so `bytes` is short of the whole. */
  overflowed: boolean;
}

/**
 * Read a response body into memory, stopping once `maxBytes` is exceeded.
 * Returns null when the response carries no body.
 *
 * Whether overflow is an error or an acceptable truncation differs per caller
 * — an oversized error body should still be shown, an oversized data body
 * must not be — so this reports the fact and lets each caller decide.
 */
export async function collectStream(
  response: Response,
  maxBytes: number,
): Promise<StreamResult | null> {
  const reader = response.body?.getReader();
  if (!reader) return null;

  const chunks: Uint8Array[] = [];
  let total = 0;
  let overflowed = false;
  while (true) {
    const { done, value } = await reader.read();
    if (done) break;
    // Keep only what fits before storing it. Appending first and measuring
    // after made maxBytes an approximation: one oversized chunk was buffered
    // whole, so the ceiling could be exceeded by the size of that chunk.
    const room = maxBytes - total;
    if (value.length > room) {
      if (room > 0) {
        chunks.push(value.subarray(0, room));
        total = maxBytes;
      }
      overflowed = true;
      await reader.cancel().catch(() => {});
      break;
    }
    chunks.push(value);
    total += value.length;
  }

  const bytes = new Uint8Array(total);
  let offset = 0;
  for (const chunk of chunks) {
    bytes.set(chunk, offset);
    offset += chunk.length;
  }
  return { bytes, overflowed };
}

/** Reads at most `maxBytes` of a body as text, discarding the rest. */
export async function readTextCapped(
  response: Response,
  maxBytes: number = MAX_ERROR_BYTES,
): Promise<string> {
  const result = await collectStream(response, maxBytes);
  return result ? new TextDecoder().decode(result.bytes) : '';
}

/**
 * Parses a JSON body, refusing one over `maxBytes`. `response.json()` buffers
 * the whole body first, which is exactly the failure this guards against.
 */
export async function readJsonWithLimit(
  response: Response,
  maxBytes: number = MAX_RESPONSE_BYTES,
): Promise<any> {
  const result = await collectStream(response, maxBytes);
  if (!result) {
    throw new Error('Google API response could not be parsed (no body)');
  }
  if (result.overflowed) {
    throw new Error(
      `Google API response is too large (over ${Math.round(maxBytes / 1024 / 1024)}MB). `
      + 'Narrow the request — for get_sheet_data, pass a smaller `range`.',
    );
  }
  try {
    return JSON.parse(new TextDecoder().decode(result.bytes));
  } catch {
    throw new Error('Google API response could not be parsed as JSON');
  }
}

async function makeGoogleRequest(
  url: string,
  accessToken: string,
  api: GoogleApi,
  options: RequestInit,
): Promise<any> {
  const response = await fetch(url, {
    ...options,
    headers: {
      Authorization: `Bearer ${accessToken}`,
      Accept: 'application/json',
      ...options.headers,
    },
  });

  if (!response.ok) {
    const errorText = await readTextCapped(response);
    const apiLabel = api === 'drive' ? 'Google Drive' : 'Google Sheets';
    let errorMessage = `${apiLabel} API error (${response.status})`;
    let reason: string | undefined;

    try {
      const errorJson = JSON.parse(errorText);
      if (errorJson.error?.message) {
        errorMessage = errorJson.error.message;
      }
      reason = errorJson.error?.status;
    } catch {
      if (errorText) errorMessage = errorText;
    }

    if (response.status === 404) {
      errorMessage = api === 'drive' ? 'File not found' : 'Spreadsheet not found';
    } else if (response.status === 403) {
      errorMessage = `Permission denied. Make sure you have granted ${apiLabel} access.`;
    } else if (response.status === 401) {
      errorMessage = 'Authentication failed. Please re-authenticate.';
    }

    const retryAfterSeconds = parseRetryAfter(response.headers.get('Retry-After'));
    throw new ApiError(errorMessage, response.status, api, retryAfterSeconds, reason);
  }

  return readJsonWithLimit(response);
}

export async function makeDriveRequest(
  endpoint: string,
  accessToken: string,
  options: RequestInit = {},
): Promise<any> {
  const url = endpoint.startsWith('http') ? endpoint : `${GOOGLE_DRIVE_API}${endpoint}`;
  return makeGoogleRequest(url, accessToken, 'drive', options);
}

export async function makeSheetsRequest(
  endpoint: string,
  accessToken: string,
  options: RequestInit = {},
): Promise<any> {
  const url = endpoint.startsWith('http') ? endpoint : `${GOOGLE_SHEETS_API}${endpoint}`;
  return makeGoogleRequest(url, accessToken, 'sheets', {
    ...options,
    headers: {
      'Content-Type': 'application/json',
      ...options.headers,
    },
  });
}

/**
 * Get the sheetId for a given sheet name from spreadsheet metadata.
 */
export async function getSheetId(
  spreadsheetId: string,
  sheetName: string,
  accessToken: string,
): Promise<number> {
  const metadata = (await makeSheetsRequest(
    `/${encodeURIComponent(spreadsheetId)}?fields=sheets.properties(sheetId,title)`,
    accessToken,
    { method: 'GET' },
  )) as { sheets: Array<{ properties: { sheetId: number; title: string } }> };

  const sheet = metadata.sheets.find((s) => s.properties.title === sheetName);
  if (!sheet) {
    throw new Error(`Sheet tab "${sheetName}" not found`);
  }
  return sheet.properties.sheetId;
}
