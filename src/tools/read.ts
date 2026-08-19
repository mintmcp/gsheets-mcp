/**
 * Read-only tools: discovery and data retrieval. Both native Google Sheets
 * and uploaded .xlsx workbooks are readable here; .xlsx falls back through
 * the office helpers when the Sheets API refuses the file.
 */

import { z } from 'zod';
import { withGoogleAuth as requirePermissionSecure } from '../auth.js';
import { wrapHandler, toolResponse } from '../lib/errors.js';
import { quoteSheetName, assertBareA1Range } from '../lib/a1.js';
import { readNativeWindow } from '../lib/sheetRead.js';
import { columnIndexToLetter, MAX_RESPONSE_COLUMNS } from '../lib/window.js';
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

/** Tabs listed by get_metadata. Matches the .xlsx path's own tab bound. */
const MAX_LISTED_TABS = 200;

/** Tabs we will fetch a header row for, bounding the opt-in second call. */
const MAX_HEADER_TABS = 50;

/**
 * Reads row 1 of each tab in a single values.batchGet. Uses the values API
 * rather than the grid API because headers only need display text, and a flat
 * string array is a fraction of the payload of a cell graph.
 */
async function fetchHeaderRows(
  spreadsheetId: string,
  tabs: Array<{ title: string; columnCount: number }>,
  accessToken: string,
): Promise<Array<string[] | undefined>> {
  const params = new URLSearchParams({
    majorDimension: 'ROWS',
    fields: 'valueRanges(values)',
  });
  for (const tab of tabs) {
    const width = Math.min(Math.max(tab.columnCount, 1), MAX_RESPONSE_COLUMNS);
    const lastColumn = columnIndexToLetter(width - 1);
    params.append('ranges', `${quoteSheetName(tab.title)}!A1:${lastColumn}1`);
  }

  const result = await makeSheetsRequest(
    `/${encodeURIComponent(spreadsheetId)}/values:batchGet?${params}`,
    accessToken,
    { method: 'GET' },
  ) as { valueRanges?: Array<{ values?: string[][] }> };

  // valueRanges come back in the order the ranges were requested.
  return tabs.map((_, i) => result.valueRanges?.[i]?.values?.[0]);
}

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
        description: 'Get spreadsheet structure: title, tab names, and each tab\'s row and column count. Call this BEFORE get_sheet_data — knowing a tab\'s size lets you request a targeted `range` instead of a blind read that comes back truncated. Set include_headers to also get the first row of each tab, which tells you what the columns actually contain. Works on uploaded Excel (.xlsx) files as well as native Google Sheets.',
        readOnlyHint: true,
        outputSchema: {
          id: z.string(),
          title: z.string(),
          sheets: z.array(z.object({
            title: z.string(),
            index: z.number(),
            rowCount: z.number().optional().describe('Rows allocated in the tab (an upper bound on used rows)'),
            columnCount: z.number().optional().describe('Columns allocated in the tab'),
            headers: z.array(z.string()).optional().describe('First row, when include_headers is set'),
          })),
          webViewLink: z.string(),
          kind: z.enum(['native', 'xlsx']).describe("'xlsx' uploads are readable but NOT editable"),
          truncated: z.boolean().optional().describe('Present only when the tab list is incomplete'),
          message: z.string().optional().describe('Explains why the tab list is incomplete'),
        },
        schema: {
          spreadsheet_id: z.string().describe('Google Sheets spreadsheet ID'),
          include_headers: z.boolean().optional().describe('Also return the first row of each tab (one extra API call, first 50 tabs)'),
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/spreadsheets", wrapHandler(async ({ spreadsheet_id, include_headers }: any, context: any) => {
          const { accessToken } = context;

          try {
            const metadata = await makeSheetsRequest(
              `/${encodeURIComponent(spreadsheet_id)}?fields=properties.title,sheets.properties,spreadsheetUrl`,
              accessToken,
              { method: 'GET' }
            ) as {
              properties: { title: string };
              sheets: Array<{
                properties: {
                  title: string;
                  index: number;
                  gridProperties?: { rowCount?: number; columnCount?: number };
                };
              }>;
              spreadsheetUrl: string;
            };

            const allTabs = metadata.sheets || [];
            const listed = allTabs.slice(0, MAX_LISTED_TABS);
            const tabsOmitted = allTabs.length - listed.length;

            const sheets = listed.map((s) => ({
              title: s.properties.title,
              index: s.properties.index,
              rowCount: s.properties.gridProperties?.rowCount ?? 0,
              columnCount: s.properties.gridProperties?.columnCount ?? 0,
            }));

            // Headers are opt-in: they cost a second call, and the dimensions
            // above already come free with the metadata fetch.
            let headersByIndex: Array<string[] | undefined> = [];
            if (include_headers && sheets.length > 0) {
              headersByIndex = await fetchHeaderRows(
                spreadsheet_id,
                sheets.slice(0, MAX_HEADER_TABS),
                accessToken,
              );
            }

            const notes: string[] = [];
            if (tabsOmitted > 0) {
              notes.push(`${tabsOmitted} further tab(s) are not listed; this spreadsheet has more than the ${MAX_LISTED_TABS}-tab limit.`);
            }
            if (include_headers && sheets.length > MAX_HEADER_TABS) {
              notes.push(`Headers were read for the first ${MAX_HEADER_TABS} tab(s) only.`);
            }

            return toolResponse({
              id: spreadsheet_id,
              title: metadata.properties.title,
              sheets: sheets.map((s, i) => (
                headersByIndex[i] ? { ...s, headers: headersByIndex[i] } : s
              )),
              webViewLink: metadata.spreadsheetUrl,
              kind: 'native' as const,
              ...(notes.length > 0 ? { truncated: true, message: notes.join(' ') } : {}),
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
        description: 'Read data from a sheet tab. Works on uploaded Excel (.xlsx) files as well as native Google Sheets — .xlsx files are read-only. Responses are capped at 5,000 cells, so a wider sheet returns fewer rows per call (5,000 ÷ column count). For native sheets, a larger tab comes back with truncated: true plus nextRange, and you should call this tool again passing that value as `range` to continue until nextRange is absent; pass an explicit bounded A1 `range` (e.g. "A1:C500") to read a specific window instead, and note that columns beyond the 256th are not returned. For .xlsx files the cap still applies but `range` and `nextRange` do not: an oversized workbook comes back with truncated: true and no way to page. Returns each cell as an object with value, and optionally formula and hyperlinks (with character ranges for mixed-content cells). If sheet_name is omitted, reads the first tab.',
        readOnlyHint: true,
        outputSchema: {
          id: z.string(),
          sheetName: z.string(),
          data: z.array(z.array(z.object({
            value: z.string(),
            type: z.enum(['string', 'number', 'boolean', 'formula', 'empty']).optional()
              .describe('Omitted for plain text cells; absent means string'),
            hyperlinks: z.array(z.object({
              url: z.string(),
              start: z.number(),
              end: z.number(),
            })).optional(),
          }))),
          rowCount: z.number(),
          columnCount: z.number(),
          kind: z.enum(['native', 'xlsx']).describe("'xlsx' uploads are readable but NOT editable"),
          // Optional because the .xlsx path reads a whole workbook rather than
          // an A1 window, so it has no range to report and cannot be paged.
          returnedRange: z.string().optional().describe('The A1 range actually returned (native sheets only)'),
          truncated: z.boolean().optional().describe('Present and true only when the response was capped'),
          nextRange: z.string().optional().describe('Pass as `range` to read the next window (native sheets only)'),
          message: z.string().optional().describe('Explains why the data is partial'),
        },
        schema: {
          spreadsheet_id: z.string().describe('Google Sheets spreadsheet ID'),
          sheet_name: z.string().optional().describe('Name of the sheet tab to read. If omitted, reads the first tab.'),
          range: z.string().optional().describe('Bounded A1 range to read, e.g. "A1:C500". Whole-column ("A:C") and whole-row ("1:3") forms are not supported. Omit to read a capped window from the top of the tab.'),
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/spreadsheets", wrapHandler(async ({ spreadsheet_id, sheet_name, range }: any, context: any) => {
          const { accessToken } = context;
          const scope = range !== undefined ? assertBareA1Range(range) : undefined;

          try {
            return toolResponse(
              await readNativeWindow(spreadsheet_id, sheet_name, scope, accessToken),
            );
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
