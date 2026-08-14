/**
 * Low-level Google Drive / Sheets HTTP helpers.
 */

import { ApiError, parseRetryAfter, type GoogleApi } from './errors.js';

export const GOOGLE_DRIVE_API = 'https://www.googleapis.com/drive/v3';
export const GOOGLE_SHEETS_API = 'https://sheets.googleapis.com/v4/spreadsheets';

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
    const errorText = await response.text();
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

  return response.json();
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
