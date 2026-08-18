/**
 * Low-level Google Drive / Sheets HTTP helpers.
 */

import { ApiError, parseRetryAfter, type GoogleApi } from './errors.js';

export const GOOGLE_DRIVE_API = 'https://www.googleapis.com/drive/v3';
export const GOOGLE_SHEETS_API = 'https://sheets.googleapis.com/v4/spreadsheets';

/** Ceiling on any single upstream body. Bounded reads stay far below this. */
export const MAX_RESPONSE_BYTES = 25 * 1024 * 1024;

/** Error bodies are small; read enough to explain the failure, never more. */
export const MAX_ERROR_BYTES = 64 * 1024;

/**
 * Reads at most `maxBytes` of a body as text, discarding the rest. Used on
 * the error path, where we want a message rather than an exception.
 */
export async function readTextCapped(
  response: Response,
  maxBytes: number = MAX_ERROR_BYTES,
): Promise<string> {
  const reader = response.body?.getReader();
  if (!reader) return '';

  const chunks: Uint8Array[] = [];
  let total = 0;
  while (total < maxBytes) {
    const { done, value } = await reader.read();
    if (done) break;
    chunks.push(value);
    total += value.length;
  }
  await reader.cancel().catch(() => {});

  const buffer = new Uint8Array(Math.min(total, maxBytes));
  let offset = 0;
  for (const chunk of chunks) {
    if (offset >= buffer.length) break;
    const slice = chunk.subarray(0, buffer.length - offset);
    buffer.set(slice, offset);
    offset += slice.length;
  }
  return new TextDecoder().decode(buffer);
}

/**
 * Reads a response body while counting bytes, aborting once the budget is
 * blown. `response.json()` buffers the whole body first, which is exactly
 * the failure this guards against.
 */
export async function readJsonWithLimit(
  response: Response,
  maxBytes: number = MAX_RESPONSE_BYTES,
): Promise<any> {
  const reader = response.body?.getReader();
  if (!reader) {
    throw new Error('Google API response could not be parsed (no body)');
  }

  const chunks: Uint8Array[] = [];
  let total = 0;
  while (true) {
    const { done, value } = await reader.read();
    if (done) break;
    total += value.length;
    if (total > maxBytes) {
      await reader.cancel().catch(() => {});
      throw new Error(
        `Google API response is too large (over ${Math.round(maxBytes / 1024 / 1024)}MB). `
        + 'Narrow the request — for get_sheet_data, pass a smaller `range`.',
      );
    }
    chunks.push(value);
  }

  const buffer = new Uint8Array(total);
  let offset = 0;
  for (const chunk of chunks) {
    buffer.set(chunk, offset);
    offset += chunk.length;
  }

  const text = new TextDecoder().decode(buffer);
  try {
    return JSON.parse(text);
  } catch {
    throw new Error('Google API response could not be parsed as JSON');
  }
}

export async function makeGoogleRequest(
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
    `/${encodeURIComponent(spreadsheetId)}?fields=sheets.properties`,
    accessToken,
    { method: 'GET' },
  )) as { sheets: Array<{ properties: { sheetId: number; title: string } }> };

  const sheet = metadata.sheets.find((s) => s.properties.title === sheetName);
  if (!sheet) {
    throw new Error(`Sheet tab "${sheetName}" not found`);
  }
  return sheet.properties.sheetId;
}
