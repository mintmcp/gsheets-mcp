/**
 * Read-only tools: discovery and data retrieval. Both native Google Sheets
 * and uploaded .xlsx workbooks are readable here; .xlsx falls back through
 * the office helpers when the Sheets API refuses the file.
 */

import { z } from 'zod';
import { withGoogleAuth as requirePermissionSecure } from '../auth.js';
import { wrapHandler, toolResponse } from '../lib/errors.js';
import { quoteSheetName } from '../lib/a1.js';
import { maxRowLength } from '../lib/grid.js';
import { buildDriveSearchQuery } from '../lib/search.js';
import { makeDriveRequest, makeSheetsRequest } from '../lib/google.js';
import {
  driveFileKind,
  isOfficeFileError,
  loadXlsxWorkbook,
  xlsxMetadataOutput,
  xlsxSheetOutput,
  toolResultWithNotice,
  READ_ONLY_NOTICE,
} from '../lib/office.js';

export const readTools = {
      search_spreadsheets: {
        description: 'Search for spreadsheets by name. Covers both native Google Sheets and uploaded Excel (.xlsx) files; the `kind` field says which. Returns matching spreadsheets with their IDs.',
        readOnlyHint: true,
        outputSchema: {
          spreadsheets: z.array(z.object({
            id: z.string(),
            name: z.string(),
            kind: z.enum(['native', 'xlsx']).describe("'xlsx' files are readable but not editable"),
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

          const q = buildDriveSearchQuery(name);

          const params = new URLSearchParams({
            pageSize: '20',
            fields: 'nextPageToken,files(id,name,mimeType,createdTime,modifiedTime,webViewLink,owners)',
            supportsAllDrives: 'true',
            includeItemsFromAllDrives: 'true',
            q,
            ...(page_token && { pageToken: page_token }),
          });

          const result = await makeDriveRequest(`/files?${params}`, accessToken);

          const spreadsheets = (result.files || []).map((file: any) => ({
            id: file.id,
            name: file.name,
            kind: driveFileKind(file.mimeType || ''),
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
        description: 'Get spreadsheet metadata including title and list of sheet tab names. Use this to discover available tabs before reading data. Works on uploaded Excel (.xlsx) files as well as native Google Sheets.',
        readOnlyHint: true,
        outputSchema: {
          id: z.string(),
          title: z.string(),
          sheets: z.array(z.object({
            title: z.string(),
            index: z.number(),
          })),
          webViewLink: z.string(),
          kind: z.enum(['native', 'xlsx']).describe("'xlsx' uploads are readable but NOT editable"),
          truncated: z.boolean().optional().describe('Present only when the tab list is incomplete'),
          message: z.string().optional().describe('Explains why the tab list is incomplete'),
        },
        schema: {
          spreadsheet_id: z.string().describe('Google Sheets spreadsheet ID'),
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/spreadsheets", wrapHandler(async ({ spreadsheet_id }: any, context: any) => {
          const { accessToken } = context;

          try {
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
              kind: 'native' as const,
            });
          } catch (err) {
            if (!isOfficeFileError(err)) throw err;
            const { meta, workbook } = await loadXlsxWorkbook(
              spreadsheet_id, accessToken, err, { namesOnly: true }
            );
            const output = xlsxMetadataOutput(
              spreadsheet_id, meta.name, meta.webViewLink, workbook
            );
            return toolResultWithNotice(output, READ_ONLY_NOTICE);
          }
        })),
      },

      get_sheet_data: {
        description: 'Read all data from a sheet tab. Works on uploaded Excel (.xlsx) files as well as native Google Sheets — .xlsx files are read-only, and very large ones come back with truncated: true. Returns each cell as an object with value, and optionally formula and hyperlinks (with character ranges for mixed-content cells). If sheet_name is omitted, reads the first tab.',
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
          kind: z.enum(['native', 'xlsx']).describe("'xlsx' uploads are readable but NOT editable"),
          truncated: z.boolean().optional().describe('Present only when the cell cap was hit'),
          message: z.string().optional().describe('Explains why the data is partial'),
        },
        schema: {
          spreadsheet_id: z.string().describe('Google Sheets spreadsheet ID'),
          sheet_name: z.string().optional().describe('Name of the sheet tab to read. If omitted, reads the first tab.'),
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/spreadsheets", wrapHandler(async ({ spreadsheet_id, sheet_name }: any, context: any) => {
          const { accessToken } = context;

          try {
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
          const columnCount = maxRowLength(data);

          return toolResponse({
            id: spreadsheet_id,
            sheetName: targetSheet,
            data,
            rowCount,
            columnCount,
            kind: 'native' as const,
          });
          } catch (err) {
            if (!isOfficeFileError(err)) throw err;
            const { workbook } = await loadXlsxWorkbook(
              spreadsheet_id, accessToken, err, { sheet: sheet_name ?? 0 }
            );
            const output = xlsxSheetOutput(spreadsheet_id, workbook, sheet_name);
            return toolResultWithNotice(output, READ_ONLY_NOTICE);
          }
        })),
      },
};
