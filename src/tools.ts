/**
 * Google Sheets MCP Tools
 */

import { z } from 'zod';
import { withGoogleAuth as requirePermissionSecure } from "./auth.js";

const GOOGLE_DRIVE_API = 'https://www.googleapis.com/drive/v3';
const GOOGLE_SHEETS_API = 'https://sheets.googleapis.com/v4/spreadsheets';

/**
 * Helper to make authenticated requests to Google Drive API
 */
async function makeDriveRequest(
  endpoint: string,
  accessToken: string,
  options: RequestInit = {}
): Promise<any> {
  const url = endpoint.startsWith('http') ? endpoint : `${GOOGLE_DRIVE_API}${endpoint}`;

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
    let errorMessage = `Google Drive API error (${response.status})`;

    try {
      const errorJson = JSON.parse(errorText);
      if (errorJson.error?.message) {
        errorMessage = errorJson.error.message;
      }
    } catch {
      errorMessage = errorText || errorMessage;
    }

    if (response.status === 404) {
      throw new Error('File not found');
    } else if (response.status === 403) {
      throw new Error('Permission denied. Make sure you have granted access.');
    } else if (response.status === 401) {
      throw new Error('Authentication failed. Please re-authenticate.');
    }

    throw new Error(errorMessage);
  }

  return response.json();
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

  const response = await fetch(url, {
    ...options,
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
      'Accept': 'application/json',
      ...options.headers,
    },
  });

  if (!response.ok) {
    const errorText = await response.text();
    let errorMessage = `Google Sheets API error (${response.status})`;

    try {
      const errorJson = JSON.parse(errorText);
      if (errorJson.error?.message) {
        errorMessage = errorJson.error.message;
      }
    } catch {
      errorMessage = errorText || errorMessage;
    }

    if (response.status === 404) {
      throw new Error('Spreadsheet not found');
    } else if (response.status === 403) {
      throw new Error('Permission denied. Make sure you have granted Sheets access.');
    } else if (response.status === 401) {
      throw new Error('Authentication failed. Please re-authenticate.');
    }

    throw new Error(errorMessage);
  }

  return response.json();
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
        handler: requirePermissionSecure("https://www.googleapis.com/auth/drive.readonly", async ({ name, page_token }: any, context: any) => {
          const { accessToken } = context;

          let q = `mimeType = 'application/vnd.google-apps.spreadsheet'`;
          if (name) {
            q += ` and name contains '${name.replace(/'/g, "\\'")}'`;
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

          const output = {
            spreadsheets,
            nextPageToken: result.nextPageToken || null,
          };
          return {
            content: [{ type: 'text', text: JSON.stringify(output, null, 2) }],
            structuredContent: output,
          };
        }),
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
        handler: requirePermissionSecure("https://www.googleapis.com/auth/spreadsheets", async ({ spreadsheet_id }: any, context: any) => {
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

          const output = {
            id: spreadsheet_id,
            title: metadata.properties.title,
            sheets: metadata.sheets.map((s) => ({
              title: s.properties.title,
              index: s.properties.index,
            })),
            webViewLink: metadata.spreadsheetUrl,
          };
          return {
            content: [{ type: 'text', text: JSON.stringify(output, null, 2) }],
            structuredContent: output,
          };
        }),
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
        handler: requirePermissionSecure("https://www.googleapis.com/auth/spreadsheets", async ({ spreadsheet_id, sheet_name }: any, context: any) => {
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

          const output = {
            id: spreadsheet_id,
            sheetName: targetSheet,
            data,
            rowCount,
            columnCount,
          };
          return {
            content: [{ type: 'text', text: JSON.stringify(output, null, 2) }],
            structuredContent: output,
          };
        }),
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
        handler: requirePermissionSecure("https://www.googleapis.com/auth/spreadsheets", async ({ title, sheet_name, parent_folder_id }: any, context: any) => {
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

            const output = {
              id: file.id,
              title: file.name,
              webViewLink: `https://docs.google.com/spreadsheets/d/${file.id}/edit`,
              message: 'Spreadsheet created successfully',
            };
            return {
              content: [{ type: 'text', text: JSON.stringify(output, null, 2) }],
              structuredContent: output,
            };
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

          const output = {
            id: result.spreadsheetId,
            title: result.properties.title,
            webViewLink: result.spreadsheetUrl,
            message: 'Spreadsheet created successfully',
          };
          return {
            content: [{ type: 'text', text: JSON.stringify(output, null, 2) }],
            structuredContent: output,
          };
        }),
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
        handler: requirePermissionSecure("https://www.googleapis.com/auth/spreadsheets", async ({ spreadsheet_id, title }: any, context: any) => {
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

          const output = {
            id: spreadsheet_id,
            sheetTitle: title,
            message: `Sheet tab "${title}" added successfully`,
          };
          return {
            content: [{ type: 'text', text: JSON.stringify(output, null, 2) }],
            structuredContent: output,
          };
        }),
      },

      insert_rows: {
        description: 'Append rows at the bottom of a sheet tab. Values are interpreted as user input, so formulas (e.g. "=SUM(A1:A2)") work automatically.',
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
        handler: requirePermissionSecure("https://www.googleapis.com/auth/spreadsheets", async ({ spreadsheet_id, sheet_name, data }: any, context: any) => {
          const { accessToken } = context;

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

          const output = {
            id: spreadsheet_id,
            updatedRows: result.updates?.updatedRows || data.length,
            message: `${result.updates?.updatedRows || data.length} row(s) appended`,
          };
          return {
            content: [{ type: 'text', text: JSON.stringify(output, null, 2) }],
            structuredContent: output,
          };
        }),
      },

      update_cell: {
        description: 'Update a single cell by A1 notation. Content is an array of text segments, each optionally with a hyperlink URL. For plain values and formulas, use a single segment. Examples: [{"text":"hello"}], [{"text":"=SUM(A1:A2)"}], [{"text":"Visit "},{"text":"Google","url":"https://google.com"},{"text":" today"}]',
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
        handler: requirePermissionSecure("https://www.googleapis.com/auth/spreadsheets", async ({ spreadsheet_id, sheet_name, cell, content }: any, context: any) => {
          const { accessToken } = context;

          if (!content || content.length === 0) {
            throw new Error('Content must have at least one segment');
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

          const output = {
            id: spreadsheet_id,
            message: `Cell ${cell} updated`,
          };
          return {
            content: [{ type: 'text', text: JSON.stringify(output, null, 2) }],
            structuredContent: output,
          };
        }),
      },

      update_range: {
        description: 'Update a range of cells with a 2D array. Ragged rows are padded with empty strings. Values are interpreted as user input, so formulas work automatically.',
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
        handler: requirePermissionSecure("https://www.googleapis.com/auth/spreadsheets", async ({ spreadsheet_id, sheet_name, range, data }: any, context: any) => {
          const { accessToken } = context;

          // Pad ragged rows
          const maxCols = Math.max(...data.map((r: string[]) => r.length));
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

          const output = {
            id: spreadsheet_id,
            updatedCells: result.updatedCells || 0,
            message: `Range ${range} updated (${result.updatedCells || 0} cells)`,
          };
          return {
            content: [{ type: 'text', text: JSON.stringify(output, null, 2) }],
            structuredContent: output,
          };
        }),
      },

      clear_values: {
        description: 'Clear values from one or more ranges in a sheet tab. Only clears cell values; formatting is preserved.',
        destructiveHint: true,
        outputSchema: {
          id: z.string(),
          clearedRanges: z.array(z.string()),
          message: z.string(),
        },
        schema: {
          spreadsheet_id: z.string().describe('Google Sheets spreadsheet ID'),
          sheet_name: z.string().describe('Name of the sheet tab'),
          ranges: z.array(z.string()).describe('Array of ranges in A1 notation to clear (e.g. ["A1:B5", "D1:D10"])'),
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/spreadsheets", async ({ spreadsheet_id, sheet_name, ranges }: any, context: any) => {
          const { accessToken } = context;

          const qualifiedRanges = ranges.map((r: string) => `${quoteSheetName(sheet_name)}!${r}`);

          const result = await makeSheetsRequest(
            `/${encodeURIComponent(spreadsheet_id)}/values:batchClear`,
            accessToken,
            {
              method: 'POST',
              body: JSON.stringify({ ranges: qualifiedRanges }),
            }
          ) as { clearedRanges: string[] };

          const output = {
            id: spreadsheet_id,
            clearedRanges: result.clearedRanges || qualifiedRanges,
            message: `Cleared ${ranges.length} range(s)`,
          };
          return {
            content: [{ type: 'text', text: JSON.stringify(output, null, 2) }],
            structuredContent: output,
          };
        }),
      },

      format_cells: {
        description: 'Apply formatting to cells in a range. Supports background color, text formatting (bold, italic, font size, font family, foreground color), alignment, wrap strategy, and number format.',
        outputSchema: {
          id: z.string(),
          message: z.string(),
        },
        schema: {
          spreadsheet_id: z.string().describe('Google Sheets spreadsheet ID'),
          sheet_name: z.string().describe('Name of the sheet tab'),
          range: z.string().describe('Range in A1 notation (e.g. "A1:C3")'),
          format: z.object({
            backgroundColor: z.object({
              red: z.coerce.number().min(0).max(1).optional(),
              green: z.coerce.number().min(0).max(1).optional(),
              blue: z.coerce.number().min(0).max(1).optional(),
            }).optional().describe('Background color with RGB values 0-1'),
            textFormat: z.object({
              bold: z.boolean().optional(),
              italic: z.boolean().optional(),
              fontSize: z.coerce.number().int().optional(),
              fontFamily: z.string().optional(),
              foregroundColor: z.object({
                red: z.coerce.number().min(0).max(1).optional(),
                green: z.coerce.number().min(0).max(1).optional(),
                blue: z.coerce.number().min(0).max(1).optional(),
              }).optional(),
            }).optional().describe('Text format options'),
            horizontalAlignment: z.enum(['LEFT', 'CENTER', 'RIGHT']).optional().describe('Horizontal alignment'),
            wrapStrategy: z.enum(['OVERFLOW_CELL', 'CLIP', 'WRAP']).optional().describe('Text wrap strategy'),
            numberFormat: z.object({
              type: z.enum(['TEXT', 'NUMBER', 'PERCENT', 'CURRENCY', 'DATE', 'TIME', 'DATE_TIME', 'SCIENTIFIC']),
              pattern: z.string().optional().describe('Format pattern (e.g. "#,##0.00", "yyyy-mm-dd")'),
            }).optional().describe('Number format'),
          }).describe('Formatting options to apply'),
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/spreadsheets", async ({ spreadsheet_id, sheet_name, range, format }: any, context: any) => {
          const { accessToken } = context;

          const sheetId = await getSheetId(spreadsheet_id, sheet_name, accessToken);
          const gridRange = parseA1Range(range);

          // Build the cell format and fields list
          const cellFormat: any = {};
          const fields: string[] = [];

          if (format.backgroundColor) {
            cellFormat.backgroundColor = format.backgroundColor;
            fields.push('userEnteredFormat.backgroundColor');
          }
          if (format.textFormat) {
            cellFormat.textFormat = format.textFormat;
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
            throw new Error('At least one format property must be provided');
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

          const output = {
            id: spreadsheet_id,
            message: `Formatting applied to ${range}`,
          };
          return {
            content: [{ type: 'text', text: JSON.stringify(output, null, 2) }],
            structuredContent: output,
          };
        }),
      },

      clear_formatting: {
        description: 'Clear all formatting from a range, resetting cells to default appearance. Cell values are preserved.',
        outputSchema: {
          id: z.string(),
          message: z.string(),
        },
        schema: {
          spreadsheet_id: z.string().describe('Google Sheets spreadsheet ID'),
          sheet_name: z.string().describe('Name of the sheet tab'),
          range: z.string().describe('Range in A1 notation (e.g. "A1:C3")'),
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/spreadsheets", async ({ spreadsheet_id, sheet_name, range }: any, context: any) => {
          const { accessToken } = context;

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

          const output = {
            id: spreadsheet_id,
            message: `Formatting cleared from ${range}`,
          };
          return {
            content: [{ type: 'text', text: JSON.stringify(output, null, 2) }],
            structuredContent: output,
          };
        }),
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
        handler: requirePermissionSecure("https://www.googleapis.com/auth/drive.readonly", async ({ spreadsheet_id, name }: any, context: any) => {
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

          const output = {
            id: result.id,
            name: result.name,
            webViewLink: result.webViewLink || `https://docs.google.com/spreadsheets/d/${result.id}`,
            message: 'Spreadsheet copied successfully',
          };
          return {
            content: [{ type: 'text', text: JSON.stringify(output, null, 2) }],
            structuredContent: output,
          };
        }),
      },
    };
  }
}
