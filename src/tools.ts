/**
 * Google Sheets MCP Tools
 */

import { z } from 'zod';
import { withGoogleAuth as requirePermissionSecure } from "./auth.js";

const GOOGLE_DRIVE_API = 'https://www.googleapis.com/drive/v3';
const GOOGLE_SHEETS_API = 'https://sheets.googleapis.com/v4/spreadsheets';

/**
 * Typed API error carrying HTTP status and optional Retry-After hint.
 * Allows `wrapHandler` to surface 429 / 4xx-specific guidance to the LLM.
 */
class ApiError extends Error {
  status: number;
  retryAfterSeconds?: number;
  api: 'drive' | 'sheets';
  constructor(message: string, status: number, api: 'drive' | 'sheets', retryAfterSeconds?: number) {
    super(message);
    this.name = 'ApiError';
    this.status = status;
    this.api = api;
    this.retryAfterSeconds = retryAfterSeconds;
  }
}

function parseRetryAfter(header: string | null): number | undefined {
  if (!header) return undefined;
  const asInt = parseInt(header, 10);
  if (!isNaN(asInt) && asInt >= 0) return asInt;
  const asDate = Date.parse(header);
  if (!isNaN(asDate)) {
    return Math.max(0, Math.round((asDate - Date.now()) / 1000));
  }
  return undefined;
}

async function makeGoogleRequest(
  url: string,
  accessToken: string,
  api: 'drive' | 'sheets',
  options: RequestInit
): Promise<any> {
  const response = await fetch(url, {
    ...options,
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Accept': 'application/json',
      ...options.headers,
    },
  });

  if (!response.ok) {
    const errorText = await response.text();
    const apiLabel = api === 'drive' ? 'Google Drive' : 'Google Sheets';
    let errorMessage = `${apiLabel} API error (${response.status})`;

    try {
      const errorJson = JSON.parse(errorText);
      if (errorJson.error?.message) {
        errorMessage = errorJson.error.message;
      }
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
    throw new ApiError(errorMessage, response.status, api, retryAfterSeconds);
  }

  return response.json();
}

/**
 * Helper to make authenticated requests to Google Drive API
 */
async function makeDriveRequest(
  endpoint: string,
  accessToken: string,
  options: RequestInit = {}
): Promise<any> {
  const url = endpoint.startsWith('http') ? endpoint : `${GOOGLE_DRIVE_API}${endpoint}`;
  return makeGoogleRequest(url, accessToken, 'drive', options);
}

/**
 * Helper to make authenticated requests to Google Sheets API
 */
async function makeSheetsRequest(
  endpoint: string,
  accessToken: string,
  options: RequestInit = {}
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
 * Shared response helpers.
 */
function toolResponse<T>(structuredContent: T) {
  return {
    content: [{ type: 'text' as const, text: JSON.stringify(structuredContent, null, 2) }],
    structuredContent,
  };
}

function toolError(message: string, extra?: Record<string, unknown>) {
  const payload = { error: message, ...(extra || {}) };
  return {
    content: [{ type: 'text' as const, text: JSON.stringify(payload) }],
    isError: true,
  };
}

/**
 * Wrap a handler so thrown errors become structured `isError: true` JSON
 * responses rather than plain-text exception strings.
 */
function wrapHandler<H extends (...args: any[]) => Promise<any>>(handler: H): H {
  return (async (...args: any[]) => {
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
  }) as H;
}

/**
 * Quote a sheet name for use in A1 notation.
 * Wraps in single quotes and escapes any existing single quotes.
 */
function quoteSheetName(name: string): string {
  return `'${name.replace(/'/g, "''")}'`;
}

/**
 * Convert a column letter (e.g. "A", "B", "AA", "AZ") to a 0-based index.
 */
function columnLetterToIndex(letter: string): number {
  let index = 0;
  const upper = letter.toUpperCase();
  for (let i = 0; i < upper.length; i++) {
    index = index * 26 + (upper.charCodeAt(i) - 64);
  }
  return index - 1; // 0-based
}

/**
 * Validate that a user-supplied range string is a bare A1 range (no sheet
 * prefix). Throws a clear error if the caller accidentally included a
 * sheet name like "Sheet1!A1:C3" — those tools take `sheet_name` as a
 * separate argument.
 */
function assertBareA1Range(range: string, paramName = 'range'): void {
  if (typeof range !== 'string' || range.length === 0) {
    throw new Error(`${paramName} must be a non-empty A1 string (e.g. "A1:C3")`);
  }
  if (range.includes('!')) {
    throw new Error(
      `${paramName} must be a bare A1 range like "A1:C3" — do not include a sheet prefix. Pass the sheet name via the sheet_name argument instead.`
    );
  }
}

/**
 * Parse an A1-style range (e.g. "A1:C3", "B2", "A1") into grid indices.
 * Returns 0-based indices suitable for GridRange.
 */
function parseA1Range(range: string): {
  startRowIndex: number;
  endRowIndex: number;
  startColumnIndex: number;
  endColumnIndex: number;
} {
  const match = range.match(/^([A-Za-z]+)(\d+)(?::([A-Za-z]+)(\d+))?$/);
  if (!match) {
    throw new Error(`Invalid A1 range: ${range}`);
  }

  const startCol = columnLetterToIndex(match[1]);
  const startRow = parseInt(match[2], 10) - 1;
  const endCol = match[3] ? columnLetterToIndex(match[3]) : startCol;
  const endRow = match[4] ? parseInt(match[4], 10) - 1 : startRow;

  return {
    startRowIndex: startRow,
    endRowIndex: endRow + 1, // exclusive
    startColumnIndex: startCol,
    endColumnIndex: endCol + 1, // exclusive
  };
}

/**
 * Parse a color input into Google Sheets' RGB float form ({red,green,blue}, 0..1).
 * Accepts:
 *   - A hex string: "#FF0000", "FF0000", "#F00", "F00" (alpha not supported)
 *   - An object with red/green/blue floats in 0..1 (passthrough)
 * Returns undefined if the input is undefined or empty.
 */
function parseColor(input: unknown): { red?: number; green?: number; blue?: number } | undefined {
  if (input === undefined || input === null) return undefined;
  if (typeof input === 'string') {
    const hex = input.trim().replace(/^#/, '');
    let r: number, g: number, b: number;
    if (/^[0-9a-fA-F]{3}$/.test(hex)) {
      r = parseInt(hex[0] + hex[0], 16);
      g = parseInt(hex[1] + hex[1], 16);
      b = parseInt(hex[2] + hex[2], 16);
    } else if (/^[0-9a-fA-F]{6}$/.test(hex)) {
      r = parseInt(hex.slice(0, 2), 16);
      g = parseInt(hex.slice(2, 4), 16);
      b = parseInt(hex.slice(4, 6), 16);
    } else {
      throw new Error(`Invalid hex color: "${input}". Use "#RRGGBB", "#RGB", or an {red,green,blue} object with floats 0..1.`);
    }
    return { red: r / 255, green: g / 255, blue: b / 255 };
  }
  if (typeof input === 'object') {
    return input as { red?: number; green?: number; blue?: number };
  }
  throw new Error('Color must be a hex string (e.g. "#FF0000") or an {red,green,blue} object with floats 0..1.');
}

/**
 * Get the sheetId for a given sheet name from spreadsheet metadata.
 */
async function getSheetId(
  spreadsheetId: string,
  sheetName: string,
  accessToken: string
): Promise<number> {
  const metadata = await makeSheetsRequest(
    `/${encodeURIComponent(spreadsheetId)}?fields=sheets.properties`,
    accessToken,
    { method: 'GET' }
  ) as { sheets: Array<{ properties: { sheetId: number; title: string } }> };

  const sheet = metadata.sheets.find(
    (s) => s.properties.title === sheetName
  );
  if (!sheet) {
    throw new Error(`Sheet tab "${sheetName}" not found`);
  }
  return sheet.properties.sheetId;
}

/**
 * Google Sheets Tools
 */
export class GoogleSheetsTools {
  static getTools() {
    return {
      search_spreadsheets: {
        description: 'Search for Google Sheets spreadsheets by name. Returns matching spreadsheets with their IDs.',
        readOnlyHint: true,
        outputSchema: {
          spreadsheets: z.array(z.object({
            id: z.string(),
            name: z.string(),
            createdTime: z.string().optional(),
            modifiedTime: z.string().optional(),
            webViewLink: z.string().optional(),
            owner: z.string().optional(),
          })),
          nextPageToken: z.string().nullable(),
        },
        schema: {
          name: z.string().describe('Search by spreadsheet name (partial match)'),
          page_token: z.string().optional().describe('Token for fetching the next page of results'),
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/drive.readonly", wrapHandler(async ({ name, page_token }: any, context: any) => {
          const { accessToken } = context;

          let q = `mimeType = 'application/vnd.google-apps.spreadsheet'`;
          if (name) {
            // Drive's q syntax: escape backslashes first, then single quotes.
            // Reject newlines outright since they break q syntax.
            if (/[\r\n]/.test(name)) {
              throw new Error('Search name must not contain newline characters');
            }
            const safeName = name.replace(/\\/g, '\\\\').replace(/'/g, "\\'");
            q += ` and name contains '${safeName}'`;
          }
          q += ` and trashed = false`;

          const params = new URLSearchParams({
            pageSize: '20',
            fields: 'nextPageToken,files(id,name,createdTime,modifiedTime,webViewLink,owners)',
            supportsAllDrives: 'true',
            includeItemsFromAllDrives: 'true',
            q,
            ...(page_token && { pageToken: page_token }),
          });

          const result = await makeDriveRequest(`/files?${params}`, accessToken);

          const spreadsheets = (result.files || []).map((file: any) => ({
            id: file.id,
            name: file.name,
            createdTime: file.createdTime,
            modifiedTime: file.modifiedTime,
            webViewLink: file.webViewLink,
            owner: file.owners?.[0]?.emailAddress,
          }));

          return toolResponse({
            spreadsheets,
            nextPageToken: result.nextPageToken || null,
          });
        })),
      },

      get_metadata: {
        description: 'Get spreadsheet metadata including title and list of sheet tab names. Use this to discover available tabs before reading data.',
        readOnlyHint: true,
        outputSchema: {
          id: z.string(),
          title: z.string(),
          sheets: z.array(z.object({
            title: z.string(),
            index: z.number(),
          })),
          webViewLink: z.string(),
        },
        schema: {
          spreadsheet_id: z.string().describe('Google Sheets spreadsheet ID'),
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/spreadsheets", wrapHandler(async ({ spreadsheet_id }: any, context: any) => {
          const { accessToken } = context;

          const metadata = await makeSheetsRequest(
            `/${encodeURIComponent(spreadsheet_id)}?fields=properties.title,sheets.properties,spreadsheetUrl`,
            accessToken,
            { method: 'GET' }
          ) as {
            properties: { title: string };
            sheets: Array<{ properties: { title: string; index: number } }>;
            spreadsheetUrl: string;
          };

          return toolResponse({
            id: spreadsheet_id,
            title: metadata.properties.title,
            sheets: metadata.sheets.map((s) => ({
              title: s.properties.title,
              index: s.properties.index,
            })),
            webViewLink: metadata.spreadsheetUrl,
          });
        })),
      },

      get_sheet_data: {
        description: 'Read all data from a sheet tab. Returns each cell as an object with value, and optionally formula and hyperlinks (with character ranges for mixed-content cells). If sheet_name is omitted, reads the first tab.',
        readOnlyHint: true,
        outputSchema: {
          id: z.string(),
          sheetName: z.string(),
          data: z.array(z.array(z.object({
            value: z.string(),
            type: z.enum(['string', 'number', 'boolean', 'formula', 'empty']),
            hyperlinks: z.array(z.object({
              url: z.string(),
              start: z.number(),
              end: z.number(),
            })).optional(),
          }))),
          rowCount: z.number(),
          columnCount: z.number(),
        },
        schema: {
          spreadsheet_id: z.string().describe('Google Sheets spreadsheet ID'),
          sheet_name: z.string().optional().describe('Name of the sheet tab to read. If omitted, reads the first tab.'),
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/spreadsheets", wrapHandler(async ({ spreadsheet_id, sheet_name }: any, context: any) => {
          const { accessToken } = context;

          // If no sheet name provided, get the first tab
          let targetSheet = sheet_name;
          if (!targetSheet) {
            const metadata = await makeSheetsRequest(
              `/${encodeURIComponent(spreadsheet_id)}?fields=sheets.properties.title`,
              accessToken,
              { method: 'GET' }
            ) as { sheets: Array<{ properties: { title: string } }> };

            if (!metadata.sheets || metadata.sheets.length === 0) {
              throw new Error('Spreadsheet has no sheets');
            }
            targetSheet = metadata.sheets[0].properties.title;
          }

          // Use Grid Data API to get values, formulas, and hyperlinks in one call
          const fields = 'sheets.data.rowData.values(userEnteredValue,formattedValue,hyperlink,textFormatRuns)';
          const result = await makeSheetsRequest(
            `/${encodeURIComponent(spreadsheet_id)}?ranges=${encodeURIComponent(quoteSheetName(targetSheet))}&includeGridData=true&fields=${encodeURIComponent(fields)}`,
            accessToken,
            { method: 'GET' }
          ) as {
            sheets: Array<{
              data: Array<{
                rowData?: Array<{
                  values?: Array<{
                    userEnteredValue?: {
                      stringValue?: string;
                      numberValue?: number;
                      boolValue?: boolean;
                      formulaValue?: string;
                    };
                    formattedValue?: string;
                    hyperlink?: string;
                    textFormatRuns?: Array<{
                      startIndex?: number;
                      format?: { link?: { uri?: string } };
                    }>;
                  }>;
                }>;
              }>;
            }>;
          };

          const rowData = result.sheets?.[0]?.data?.[0]?.rowData || [];

          const data = rowData.map((row) => {
            return (row.values || []).map((cell) => {
              const uev = cell.userEnteredValue;

              // Determine type and value
              let type: 'string' | 'number' | 'boolean' | 'formula' | 'empty';
              let value: string;

              if (!uev) {
                type = 'empty';
                value = '';
              } else if (uev.formulaValue !== undefined) {
                type = 'formula';
                value = uev.formulaValue;
              } else if (uev.numberValue !== undefined) {
                type = 'number';
                value = cell.formattedValue || String(uev.numberValue);
              } else if (uev.boolValue !== undefined) {
                type = 'boolean';
                value = cell.formattedValue || String(uev.boolValue);
              } else {
                type = 'string';
                value = cell.formattedValue || uev.stringValue || '';
              }

              const cellObj: { value: string; type: string; hyperlinks?: Array<{ url: string; start: number; end: number }> } = { value, type };

              // Extract hyperlinks from textFormatRuns (mixed content)
              const runs = cell.textFormatRuns;
              if (runs && runs.length > 0) {
                const displayText = cell.formattedValue || value;
                const hyperlinks: Array<{ url: string; start: number; end: number }> = [];
                for (let i = 0; i < runs.length; i++) {
                  const run = runs[i];
                  if (run.format?.link?.uri) {
                    const start = run.startIndex || 0;
                    const end = i + 1 < runs.length ? (runs[i + 1].startIndex || displayText.length) : displayText.length;
                    hyperlinks.push({ url: run.format.link.uri, start, end });
                  }
                }
                if (hyperlinks.length > 0) {
                  cellObj.hyperlinks = hyperlinks;
                }
              } else if (cell.hyperlink) {
                // Whole-cell hyperlink (no textFormatRuns)
                const displayText = cell.formattedValue || value;
                cellObj.hyperlinks = [{ url: cell.hyperlink, start: 0, end: displayText.length }];
              }

              return cellObj;
            });
          });

          const rowCount = data.length;
          const columnCount = rowCount > 0 ? Math.max(...data.map((r) => r.length)) : 0;

          return toolResponse({
            id: spreadsheet_id,
            sheetName: targetSheet,
            data,
            rowCount,
            columnCount,
          });
        })),
      },

      create_spreadsheet: {
        description: 'Create a new Google Sheets spreadsheet with an optional first tab name. Optionally place it in a specific folder (including shared drive folders).',
        outputSchema: {
          id: z.string(),
          title: z.string(),
          webViewLink: z.string(),
          message: z.string(),
        },
        schema: {
          title: z.string().describe('Title for the new spreadsheet'),
          sheet_name: z.string().optional().describe('Name for the first sheet tab (defaults to "Sheet1")'),
          parent_folder_id: z.string().optional().describe('ID of the folder to create the spreadsheet in (supports shared drive folders)'),
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/spreadsheets", wrapHandler(async ({ title, sheet_name, parent_folder_id }: any, context: any) => {
          const { accessToken } = context;

          // Create via Drive API to support parent folder placement
          if (parent_folder_id) {
            const fileMetadata: any = {
              name: title,
              mimeType: 'application/vnd.google-apps.spreadsheet',
              parents: [parent_folder_id],
            };

            const file = await makeDriveRequest(
              `/files?supportsAllDrives=true`,
              accessToken,
              {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(fileMetadata),
              }
            ) as { id: string; name: string };

            // Rename the default sheet tab if requested
            if (sheet_name) {
              const spreadsheet = await makeSheetsRequest(`/${file.id}`, accessToken, { method: 'GET' }) as any;
              const defaultSheetId = spreadsheet.sheets?.[0]?.properties?.sheetId;
              if (defaultSheetId !== undefined) {
                await makeSheetsRequest(`/${file.id}:batchUpdate`, accessToken, {
                  method: 'POST',
                  body: JSON.stringify({
                    requests: [{
                      updateSheetProperties: {
                        properties: { sheetId: defaultSheetId, title: sheet_name },
                        fields: 'title',
                      },
                    }],
                  }),
                });
              }
            }

            return toolResponse({
              id: file.id,
              title: file.name,
              webViewLink: `https://docs.google.com/spreadsheets/d/${file.id}/edit`,
              message: 'Spreadsheet created successfully',
            });
          }

          // Default: create via Sheets API (My Drive)
          const result = await makeSheetsRequest('', accessToken, {
            method: 'POST',
            body: JSON.stringify({
              properties: { title },
              sheets: [{
                properties: { title: sheet_name || 'Sheet1' },
              }],
            }),
          }) as { spreadsheetId: string; properties: { title: string }; spreadsheetUrl: string };

          return toolResponse({
            id: result.spreadsheetId,
            title: result.properties.title,
            webViewLink: result.spreadsheetUrl,
            message: 'Spreadsheet created successfully',
          });
        })),
      },

      add_sheet: {
        description: 'Add a new sheet tab to an existing spreadsheet.',
        outputSchema: {
          id: z.string(),
          sheetTitle: z.string(),
          message: z.string(),
        },
        schema: {
          spreadsheet_id: z.string().describe('Google Sheets spreadsheet ID'),
          title: z.string().describe('Name for the new sheet tab'),
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/spreadsheets", wrapHandler(async ({ spreadsheet_id, title }: any, context: any) => {
          const { accessToken } = context;

          await makeSheetsRequest(
            `/${encodeURIComponent(spreadsheet_id)}:batchUpdate`,
            accessToken,
            {
              method: 'POST',
              body: JSON.stringify({
                requests: [{ addSheet: { properties: { title } } }],
              }),
            }
          );

          return toolResponse({
            id: spreadsheet_id,
            sheetTitle: title,
            message: `Sheet tab "${title}" added successfully`,
          });
        })),
      },

      insert_rows: {
        description: 'Append rows AFTER the last non-empty row of a sheet tab (using Sheets values:append with INSERT_ROWS). Values are interpreted as user input (USER_ENTERED), so formulas (e.g. "=SUM(A1:A2)") work automatically — but note: a leading "=" always becomes a formula, and string-typed values like "01" or "1.0" may be coerced (e.g. "01" → 1). Use update_range to overwrite an exact range of existing cells; use this tool when you want to add new rows at the end without specifying a target range.',
        outputSchema: {
          id: z.string(),
          updatedRows: z.number(),
          message: z.string(),
        },
        schema: {
          spreadsheet_id: z.string().describe('Google Sheets spreadsheet ID'),
          sheet_name: z.string().describe('Name of the sheet tab to append to'),
          data: z.array(z.array(z.string())).describe('Rows to append. Each row is an array of cell values. Formulas like "=SUM(A1:A2)" are supported.'),
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/spreadsheets", wrapHandler(async ({ spreadsheet_id, sheet_name, data }: any, context: any) => {
          const { accessToken } = context;

          if (!Array.isArray(data) || data.length === 0) {
            throw new Error('data must contain at least one row');
          }

          const params = new URLSearchParams({
            valueInputOption: 'USER_ENTERED',
            insertDataOption: 'INSERT_ROWS',
          });

          const result = await makeSheetsRequest(
            `/${encodeURIComponent(spreadsheet_id)}/values/${encodeURIComponent(quoteSheetName(sheet_name))}:append?${params}`,
            accessToken,
            {
              method: 'POST',
              body: JSON.stringify({ values: data }),
            }
          ) as { updates: { updatedRows: number } };

          const updatedRows = result.updates?.updatedRows ?? 0;
          return toolResponse({
            id: spreadsheet_id,
            updatedRows,
            message: `${updatedRows} row(s) appended`,
          });
        })),
      },

      update_cell: {
        description: 'Update a SINGLE cell by A1 notation, with optional inline hyperlinks. Content is an array of text segments, each optionally hyperlinked. For plain values and formulas, use a single segment. Values are interpreted as user input (USER_ENTERED): a leading "=" becomes a formula, and string-typed values like "01" may be coerced. Use update_range for ranges; use insert_rows to append. Examples: [{"text":"hello"}], [{"text":"=SUM(A1:A2)"}], [{"text":"Visit "},{"text":"Google","url":"https://google.com"},{"text":" today"}].',
        outputSchema: {
          id: z.string(),
          message: z.string(),
        },
        schema: {
          spreadsheet_id: z.string().describe('Google Sheets spreadsheet ID'),
          sheet_name: z.string().describe('Name of the sheet tab'),
          cell: z.string().describe('Cell in A1 notation (e.g. "B3", "AA1")'),
          content: z.array(z.object({
            text: z.string().describe('Text content for this segment'),
            url: z.string().optional().describe('Hyperlink URL for this segment (omit for plain text)'),
          })).describe('Cell content as text segments, each optionally hyperlinked'),
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/spreadsheets", wrapHandler(async ({ spreadsheet_id, sheet_name, cell, content }: any, context: any) => {
          const { accessToken } = context;

          if (!content || content.length === 0) {
            throw new Error('Content must have at least one segment');
          }
          if (typeof cell !== 'string' || cell.includes(':')) {
            throw new Error('cell must be a single cell in A1 notation (e.g. "B3"), not a range. Use update_range for ranges.');
          }
          if (!/^[A-Za-z]+\d+$/.test(cell)) {
            throw new Error(`Invalid A1 cell: ${cell}`);
          }

          const hasUrls = content.some((c: any) => c.url);

          if (!hasUrls) {
            // No hyperlinks: use Values API with USER_ENTERED for auto type detection
            const fullText = content.map((c: any) => c.text).join('');
            const range = `${quoteSheetName(sheet_name)}!${cell}`;
            const params = new URLSearchParams({
              valueInputOption: 'USER_ENTERED',
            });

            await makeSheetsRequest(
              `/${encodeURIComponent(spreadsheet_id)}/values/${encodeURIComponent(range)}?${params}`,
              accessToken,
              {
                method: 'PUT',
                body: JSON.stringify({ values: [[fullText]] }),
              }
            );
          } else {
            // Has hyperlinks: use batchUpdate with textFormatRuns
            const sheetId = await getSheetId(spreadsheet_id, sheet_name, accessToken);
            const gridRange = parseA1Range(cell);

            const fullText = content.map((c: any) => c.text).join('');
            const textFormatRuns: Array<{ startIndex: number; format: any }> = [];
            let offset = 0;

            for (const segment of content) {
              const format: any = {};
              if (segment.url) {
                format.link = { uri: segment.url };
              }
              textFormatRuns.push({ startIndex: offset, format });
              offset += segment.text.length;
            }

            await makeSheetsRequest(
              `/${encodeURIComponent(spreadsheet_id)}:batchUpdate`,
              accessToken,
              {
                method: 'POST',
                body: JSON.stringify({
                  requests: [{
                    updateCells: {
                      range: { sheetId, ...gridRange },
                      rows: [{
                        values: [{
                          userEnteredValue: { stringValue: fullText },
                          textFormatRuns,
                        }],
                      }],
                      fields: 'userEnteredValue,textFormatRuns',
                    },
                  }],
                }),
              }
            );
          }

          return toolResponse({
            id: spreadsheet_id,
            message: `Cell ${cell} updated`,
          });
        })),
      },

      update_range: {
        description: 'Overwrite a range of cells with a 2D array (values:PUT). Pass `range` as a bare A1 string (e.g. "A1:C3") — do NOT include a sheet prefix; use the `sheet_name` argument for that. Ragged rows are padded with empty strings. Values are interpreted as user input (USER_ENTERED): a leading "=" becomes a formula, and string-typed values like "01" may be coerced. Use update_cell for a single cell (especially when you need inline hyperlinks); use insert_rows to add new rows at the end.',
        outputSchema: {
          id: z.string(),
          updatedCells: z.number(),
          message: z.string(),
        },
        schema: {
          spreadsheet_id: z.string().describe('Google Sheets spreadsheet ID'),
          sheet_name: z.string().describe('Name of the sheet tab'),
          range: z.string().describe('Range in A1 notation (e.g. "A1:C3")'),
          data: z.array(z.array(z.string())).describe('2D array of values. Formulas like "=SUM(A1:A2)" are supported.'),
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/spreadsheets", wrapHandler(async ({ spreadsheet_id, sheet_name, range, data }: any, context: any) => {
          const { accessToken } = context;

          assertBareA1Range(range);

          if (!Array.isArray(data) || data.length === 0) {
            throw new Error('data must contain at least one row');
          }

          // Pad ragged rows
          const maxCols = Math.max(...data.map((r: string[]) => r.length));
          if (maxCols === 0) {
            throw new Error('data rows must contain at least one cell');
          }
          const paddedData = data.map((row: string[]) => {
            const padded = [...row];
            while (padded.length < maxCols) {
              padded.push('');
            }
            return padded;
          });

          const a1Range = `${quoteSheetName(sheet_name)}!${range}`;
          const params = new URLSearchParams({
            valueInputOption: 'USER_ENTERED',
          });

          const result = await makeSheetsRequest(
            `/${encodeURIComponent(spreadsheet_id)}/values/${encodeURIComponent(a1Range)}?${params}`,
            accessToken,
            {
              method: 'PUT',
              body: JSON.stringify({ values: paddedData }),
            }
          ) as { updatedCells: number };

          return toolResponse({
            id: spreadsheet_id,
            updatedCells: result.updatedCells || 0,
            message: `Range ${range} updated (${result.updatedCells || 0} cells)`,
          });
        })),
      },

      clear_values: {
        description: 'Clear cell values from one or more ranges in a sheet tab. Pass each range as a bare A1 string (e.g. "A1:B5") — do NOT include a sheet prefix; use the `sheet_name` argument for that. Only values are cleared; formatting is preserved. Use clear_formatting to reset visual styling instead.',
        destructiveHint: true,
        outputSchema: {
          id: z.string(),
          clearedRanges: z.array(z.string()),
          message: z.string(),
        },
        schema: {
          spreadsheet_id: z.string().describe('Google Sheets spreadsheet ID'),
          sheet_name: z.string().describe('Name of the sheet tab'),
          ranges: z.array(z.string()).min(1).describe('Array of ranges in A1 notation to clear (e.g. ["A1:B5", "D1:D10"])'),
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/spreadsheets", wrapHandler(async ({ spreadsheet_id, sheet_name, ranges }: any, context: any) => {
          const { accessToken } = context;

          if (!Array.isArray(ranges) || ranges.length === 0) {
            throw new Error('ranges must contain at least one A1 range');
          }
          for (const r of ranges) {
            assertBareA1Range(r, 'ranges[]');
          }

          const qualifiedRanges = ranges.map((r: string) => `${quoteSheetName(sheet_name)}!${r}`);

          const result = await makeSheetsRequest(
            `/${encodeURIComponent(spreadsheet_id)}/values:batchClear`,
            accessToken,
            {
              method: 'POST',
              body: JSON.stringify({ ranges: qualifiedRanges }),
            }
          ) as { clearedRanges: string[] };

          return toolResponse({
            id: spreadsheet_id,
            clearedRanges: result.clearedRanges || qualifiedRanges,
            message: `Cleared ${ranges.length} range(s)`,
          });
        })),
      },

      format_cells: {
        description: 'Apply formatting to cells in a range. Pass `range` as a bare A1 string (e.g. "A1:C3") — do NOT include a sheet prefix; use the `sheet_name` argument for that. Supports background color, text formatting (bold, italic, font size, font family, foreground color), alignment, wrap strategy, and number format. Colors accept either hex strings (e.g. "#FF0000", "#F00") or {red,green,blue} float objects (0..1). Use clear_formatting to reset styling.',
        outputSchema: {
          id: z.string(),
          message: z.string(),
        },
        schema: {
          spreadsheet_id: z.string().describe('Google Sheets spreadsheet ID'),
          sheet_name: z.string().describe('Name of the sheet tab'),
          range: z.string().describe('Range in A1 notation (e.g. "A1:C3")'),
          format: z.object({
            backgroundColor: z.union([
              z.string(),
              z.object({
                red: z.coerce.number().min(0).max(1).optional(),
                green: z.coerce.number().min(0).max(1).optional(),
                blue: z.coerce.number().min(0).max(1).optional(),
              }),
            ]).optional().describe('Background color. Accepts a hex string (e.g. "#FF0000", "#F00") or an {red,green,blue} object with floats 0..1.'),
            textFormat: z.object({
              bold: z.boolean().optional(),
              italic: z.boolean().optional(),
              fontSize: z.coerce.number().int().optional(),
              fontFamily: z.string().optional(),
              foregroundColor: z.union([
                z.string(),
                z.object({
                  red: z.coerce.number().min(0).max(1).optional(),
                  green: z.coerce.number().min(0).max(1).optional(),
                  blue: z.coerce.number().min(0).max(1).optional(),
                }),
              ]).optional().describe('Foreground color. Accepts a hex string (e.g. "#000000") or an {red,green,blue} object with floats 0..1.'),
            }).optional().describe('Text format options'),
            horizontalAlignment: z.enum(['LEFT', 'CENTER', 'RIGHT']).optional().describe('Horizontal alignment'),
            wrapStrategy: z.enum(['OVERFLOW_CELL', 'CLIP', 'WRAP']).optional().describe('Text wrap strategy'),
            numberFormat: z.object({
              type: z.enum(['TEXT', 'NUMBER', 'PERCENT', 'CURRENCY', 'DATE', 'TIME', 'DATE_TIME', 'SCIENTIFIC']),
              pattern: z.string().optional().describe('Format pattern (e.g. "#,##0.00", "yyyy-mm-dd")'),
            }).optional().describe('Number format'),
          }).describe('Formatting options to apply'),
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/spreadsheets", wrapHandler(async ({ spreadsheet_id, sheet_name, range, format }: any, context: any) => {
          const { accessToken } = context;

          assertBareA1Range(range);
          const sheetId = await getSheetId(spreadsheet_id, sheet_name, accessToken);
          const gridRange = parseA1Range(range);

          // Build the cell format and fields list
          const cellFormat: any = {};
          const fields: string[] = [];

          if (format.backgroundColor !== undefined) {
            const parsed = parseColor(format.backgroundColor);
            if (parsed) {
              cellFormat.backgroundColor = parsed;
              fields.push('userEnteredFormat.backgroundColor');
            }
          }
          if (format.textFormat) {
            const tf = { ...format.textFormat };
            if (tf.foregroundColor !== undefined) {
              const parsed = parseColor(tf.foregroundColor);
              if (parsed) tf.foregroundColor = parsed;
              else delete tf.foregroundColor;
            }
            cellFormat.textFormat = tf;
            fields.push('userEnteredFormat.textFormat');
          }
          if (format.horizontalAlignment) {
            cellFormat.horizontalAlignment = format.horizontalAlignment;
            fields.push('userEnteredFormat.horizontalAlignment');
          }
          if (format.wrapStrategy) {
            cellFormat.wrapStrategy = format.wrapStrategy;
            fields.push('userEnteredFormat.wrapStrategy');
          }
          if (format.numberFormat) {
            cellFormat.numberFormat = format.numberFormat;
            fields.push('userEnteredFormat.numberFormat');
          }

          if (fields.length === 0) {
            throw new Error('format must include at least one of: backgroundColor, textFormat, horizontalAlignment, wrapStrategy, numberFormat');
          }

          await makeSheetsRequest(
            `/${encodeURIComponent(spreadsheet_id)}:batchUpdate`,
            accessToken,
            {
              method: 'POST',
              body: JSON.stringify({
                requests: [{
                  repeatCell: {
                    range: {
                      sheetId,
                      ...gridRange,
                    },
                    cell: {
                      userEnteredFormat: cellFormat,
                    },
                    fields: fields.join(','),
                  },
                }],
              }),
            }
          );

          return toolResponse({
            id: spreadsheet_id,
            message: `Formatting applied to ${range}`,
          });
        })),
      },

      clear_formatting: {
        description: 'Clear all formatting from a range, resetting cells to default appearance. Pass `range` as a bare A1 string (e.g. "A1:C3") — do NOT include a sheet prefix; use the `sheet_name` argument for that. Cell values are preserved. Use clear_values to clear cell contents instead.',
        destructiveHint: true,
        outputSchema: {
          id: z.string(),
          message: z.string(),
        },
        schema: {
          spreadsheet_id: z.string().describe('Google Sheets spreadsheet ID'),
          sheet_name: z.string().describe('Name of the sheet tab'),
          range: z.string().describe('Range in A1 notation (e.g. "A1:C3")'),
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/spreadsheets", wrapHandler(async ({ spreadsheet_id, sheet_name, range }: any, context: any) => {
          const { accessToken } = context;

          assertBareA1Range(range);
          const sheetId = await getSheetId(spreadsheet_id, sheet_name, accessToken);
          const gridRange = parseA1Range(range);

          await makeSheetsRequest(
            `/${encodeURIComponent(spreadsheet_id)}:batchUpdate`,
            accessToken,
            {
              method: 'POST',
              body: JSON.stringify({
                requests: [{
                  repeatCell: {
                    range: {
                      sheetId,
                      ...gridRange,
                    },
                    cell: {
                      userEnteredFormat: {},
                    },
                    fields: 'userEnteredFormat',
                  },
                }],
              }),
            }
          );

          return toolResponse({
            id: spreadsheet_id,
            message: `Formatting cleared from ${range}`,
          });
        })),
      },

      copy_spreadsheet: {
        description: 'Create a copy of an entire spreadsheet via Google Drive. Optionally provide a new name.',
        outputSchema: {
          id: z.string(),
          name: z.string(),
          webViewLink: z.string(),
          message: z.string(),
        },
        schema: {
          spreadsheet_id: z.string().describe('Google Sheets spreadsheet ID to copy'),
          name: z.string().optional().describe('Name for the copy (defaults to "Copy of <original>")'),
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/drive.file", wrapHandler(async ({ spreadsheet_id, name }: any, context: any) => {
          const { accessToken } = context;

          const body: any = {};
          if (name) {
            body.name = name;
          }

          const result = await makeDriveRequest(
            `/files/${encodeURIComponent(spreadsheet_id)}/copy?supportsAllDrives=true`,
            accessToken,
            {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify(body),
            }
          ) as { id: string; name: string; webViewLink?: string };

          return toolResponse({
            id: result.id,
            name: result.name,
            webViewLink: result.webViewLink || `https://docs.google.com/spreadsheets/d/${result.id}`,
            message: 'Spreadsheet copied successfully',
          });
        })),
      },
    };
  }
}
