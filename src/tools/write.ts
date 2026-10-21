/**
 * Tools that create spreadsheets or mutate cell values. Cell mutation is
 * native-only (`nativeOnly`); Drive-level operations also accept .xlsx.
 */

import { z } from 'zod';
import { withGoogleAuth as requirePermissionSecure } from '../auth.js';
import { wrapHandler, toolResponse } from '../lib/errors.js';
import { quoteSheetName, assertSingleCell, parseA1Range, normalizeRange, rangeOnTab, requireTab } from '../lib/a1.js';
import { padRaggedRows } from '../lib/grid.js';
import { optionalSheetNameSchema } from './common.js';
import { makeDriveRequest, makeSheetsRequest, getSheetId } from '../lib/google.js';
import {
  nativeOnly,
  fetchDriveFileMeta,
  XLSX_MIME,
  XLS_MIME,
  NATIVE_SHEET_MIME,
} from '../lib/office.js';

/**
 * A cell as the values API takes it: Google types the cell from the JSON
 * type (bool, string, double), and skips a null.
 */
const cellValueSchema = z.union([z.string(), z.number(), z.boolean(), z.null()]);

const dataSchema = z.array(z.array(cellValueSchema))
  .describe('2D array of rows. Strings, numbers and booleans are written as that type; a string starting with "=" is a formula; null skips the cell, which keeps its value in update_range and stays blank in rows insert_rows appends.');

const RANGE_FORMS = 'Accepts bare A1 ("A1:C3"), a sheet-qualified range ("Sheet1!A1:C3", which may stand in for sheet_name and must agree with it when both are given) or absolute R1C1 ("R1C1:R3C3").';

/** Qualifies a bare A1 with the tab, the one form the values API never misreads. */
const qualify = (sheetName: string, a1: string) => `${quoteSheetName(sheetName)}!${a1}`;


export const writeTools = {
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
        handler: requirePermissionSecure("https://www.googleapis.com/auth/spreadsheets", wrapHandler(nativeOnly(async ({ spreadsheet_id, title }: any, context: any) => {
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
        }))),
      },

      insert_rows: {
        description: 'Append rows after the last row of a table of data (Sheets values:append with INSERT_ROWS). Google searches `range` for a table and appends under it; without `range` it searches the whole tab, which on a tab holding one table means after its last row. When a tab holds several tables stacked or side by side, pass `range` to pick one: "F1" appends under the table that starts at F1. Google inserts whole sheet rows for the new data, so anything beside the target table in those rows (another table in A:D when appending to F:G) shifts down with it; to add under one of two side-by-side tables without moving the other, write the cells below it with update_range instead. ' + RANGE_FORMS + ' This tool cannot insert rows at an arbitrary index or add columns; use insert_dimensions for that and update_range for mid-sheet writes. Strings, numbers and booleans are written as typed values and parsed as user input (USER_ENTERED): a leading "=" becomes a formula, and a string like "01" may be coerced to 1.',
        outputSchema: {
          id: z.string(),
          updatedRows: z.number(),
          updatedRange: z.string().optional().describe('Where the rows landed, sheet-qualified A1'),
          message: z.string(),
        },
        schema: {
          spreadsheet_id: z.string().describe('Google Sheets spreadsheet ID'),
          sheet_name: z.string().optional().describe('Name of the sheet tab to append to. May be omitted when `range` carries the tab ("Sheet1!A1").'),
          range: z.string().optional().describe('Range to search for the logical table to append to, e.g. "A1", "Sheet1!F1:H1" or whole columns "Sheet1!A:D". Omit to search the whole tab.'),
          data: dataSchema,
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/spreadsheets", wrapHandler(nativeOnly(async ({ spreadsheet_id, sheet_name, range, data }: any, context: any) => {
          const { accessToken } = context;

          if (!Array.isArray(data) || data.length === 0) {
            throw new Error('data must contain at least one row');
          }
          let anchor: string;
          if (range === undefined) {
            anchor = quoteSheetName(requireTab(sheet_name));
          } else {
            const target = rangeOnTab(sheet_name, range, { allowOpen: true });
            anchor = qualify(target.sheetName, target.a1);
          }

          const params = new URLSearchParams({
            valueInputOption: 'USER_ENTERED',
            insertDataOption: 'INSERT_ROWS',
          });

          const result = await makeSheetsRequest(
            `/${encodeURIComponent(spreadsheet_id)}/values/${encodeURIComponent(anchor)}:append?${params}`,
            accessToken,
            {
              method: 'POST',
              body: JSON.stringify({ values: data }),
            }
          ) as { updates?: { updatedRows?: number; updatedRange?: string } };

          const updatedRows = result.updates?.updatedRows ?? 0;
          const updatedRange = result.updates?.updatedRange;
          return toolResponse({
            id: spreadsheet_id,
            updatedRows,
            ...(updatedRange !== undefined && { updatedRange }),
            message: `${updatedRows} row(s) appended${updatedRange ? ` at ${updatedRange}` : ''}`,
          });
        }))),
      },

      update_cell: {
        description: 'Write ONE cell as text segments with inline hyperlinks, e.g. [{"text":"Visit "},{"text":"Google","url":"https://google.com"},{"text":" today"}]. This is the tool for links and rich text; for plain values, numbers, booleans and formulas use update_range with a one-cell range ("B3"). `cell` is a single cell: "B3", "Sheet1!B3" (may stand in for sheet_name and must agree with it when both are given) or absolute R1C1 "R3C2". Segments without any url are joined and written as user input (USER_ENTERED), so a leading "=" becomes a formula; with a url the whole cell is stored as text. Use insert_rows to append.',
        outputSchema: {
          id: z.string(),
          message: z.string(),
        },
        schema: {
          spreadsheet_id: z.string().describe('Google Sheets spreadsheet ID'),
          sheet_name: optionalSheetNameSchema.describe('Name of the sheet tab. May be omitted when `cell` carries the tab ("Sheet1!B3").'),
          cell: z.string().describe('Single cell: "B3", "Sheet1!B3" or "R3C2"'),
          content: z.array(z.object({
            text: z.string().describe('Text content for this segment'),
            url: z.string().optional().describe('Hyperlink URL for this segment (omit for plain text)'),
          })).describe('Cell content as text segments, each optionally hyperlinked'),
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/spreadsheets", wrapHandler(nativeOnly(async ({ spreadsheet_id, sheet_name, cell, content }: any, context: any) => {
          const { accessToken } = context;

          if (!content || content.length === 0) {
            throw new Error('Content must have at least one segment');
          }
          const target = rangeOnTab(sheet_name, cell, { paramName: 'cell' });
          const cleanCell = assertSingleCell(target.a1);

          const hasUrls = content.some((c: any) => c.url);

          if (!hasUrls) {
            // No hyperlinks: use Values API with USER_ENTERED for auto type detection
            const fullText = content.map((c: any) => c.text).join('');
            const range = qualify(target.sheetName, cleanCell);
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
            const sheetId = await getSheetId(spreadsheet_id, target.sheetName, accessToken);
            const gridRange = parseA1Range(cleanCell);

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
            message: `Cell ${cleanCell} updated`,
          });
        }))),
      },

      update_range: {
        description: 'Overwrite a range of cells with a 2D array (values:PUT). Pass `range` as a bounded range: a single cell like "A1" or a rectangle like "A1:C3". ' + RANGE_FORMS + ' Whole-column ("A:C") and whole-row ("1:3") forms are not supported. Sizing rules: for a multi-cell `range` the `data` matrix must fit within the range — the Sheets API rejects oversized matrices with a 400 INVALID_ARGUMENT, and if `data` is smaller than the range only the supplied cells are written (the rest keep their prior values). For a single-cell `range` (e.g. "A1"), the cell acts as a top-left anchor and the matrix expands down and right from it. Ragged rows are padded with empty strings. Strings, numbers and booleans are written as typed values and parsed as user input (USER_ENTERED): a string starting with "=" becomes a formula (so this is also the tool for writing formulas), and a string like "01" may be coerced to 1. To preserve a formula read via get_sheet_data, write its `formula` back, not its `value` — the value is only the last computed result. Use update_cell only when a cell needs inline hyperlinks; use insert_rows to add new rows at the end.',
        outputSchema: {
          id: z.string(),
          updatedCells: z.number(),
          message: z.string(),
        },
        schema: {
          spreadsheet_id: z.string().describe('Google Sheets spreadsheet ID'),
          sheet_name: optionalSheetNameSchema,
          range: z.string().describe('Bounded range: "A1", "A1:C3", "Sheet1!A1:C3" or "R1C1:R3C3". Whole-column/whole-row forms are not supported.'),
          data: dataSchema,
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/spreadsheets", wrapHandler(nativeOnly(async ({ spreadsheet_id, sheet_name, range, data }: any, context: any) => {
          const { accessToken } = context;

          const target = rangeOnTab(sheet_name, range);
          const paddedData = padRaggedRows(data);

          const a1Range = qualify(target.sheetName, target.a1);
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
            message: `Range ${target.a1} updated (${result.updatedCells || 0} cells)`,
          });
        }))),
      },

      clear_values: {
        description: 'Clear cell values from one or more ranges, on one tab or across several (values:batchClear). Each range may be bare ("A1:B5", qualified by `sheet_name`) or carry its own tab ("Other!C1:C10"), and may be a whole column ("A:C") or whole row ("1:3"); absolute R1C1 is accepted too. Only values are cleared; formatting is preserved. Use clear_formatting to reset visual styling instead.',
        destructiveHint: true,
        outputSchema: {
          id: z.string(),
          clearedRanges: z.array(z.string()),
          message: z.string(),
        },
        schema: {
          spreadsheet_id: z.string().describe('Google Sheets spreadsheet ID'),
          sheet_name: z.string().optional().describe('Tab for ranges that carry no "Sheet!" prefix. Required unless every range is qualified.'),
          ranges: z.array(z.string()).min(1).describe('Ranges to clear, e.g. ["A1:B5", "D:D", "Other!A1:C3"]'),
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/spreadsheets", wrapHandler(nativeOnly(async ({ spreadsheet_id, sheet_name, ranges }: any, context: any) => {
          const { accessToken } = context;

          if (!Array.isArray(ranges) || ranges.length === 0) {
            throw new Error('ranges must contain at least one A1 range');
          }
          const qualifiedRanges = ranges.map((r: unknown, i: number) => {
            const paramName = `ranges[${i}]`;
            // A prefix wins over sheet_name here: the tool clears across tabs,
            // so a qualified range names its own tab and a bare one takes the default
            const parsed = normalizeRange(r, { paramName, allowOpen: true });
            return qualify(requireTab(parsed.sheetName ?? sheet_name, paramName), parsed.a1);
          });

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
            message: `Cleared ${qualifiedRanges.length} range(s)`,
          });
        }))),
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

          const meta = await fetchDriveFileMeta(spreadsheet_id, accessToken);
          if (meta.mimeType === XLSX_MIME || meta.mimeType === XLS_MIME) {
            throw new Error(
              `'${meta.name}' is an Excel upload, not a native Google Sheet: copying it would produce ` +
              `another read-only Excel file. Call convert_to_google_sheet with file_id '${spreadsheet_id}' ` +
              `to get an editable native copy instead.`
            );
          }

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

      convert_to_google_sheet: {
        description:
          'Convert an uploaded Excel (.xlsx) file into a NEW, editable native Google Sheet. ' +
          'The original .xlsx is left untouched. Use this when the user wants to edit a file ' +
          'that search_spreadsheets or get_sheet_data reported as kind "xlsx". Drive performs the ' +
          'conversion, so number formats, formulas, hyperlinks and every tab are preserved — far ' +
          'better than re-typing the data into a new sheet.',
        outputSchema: {
          id: z.string().describe('ID of the new native Google Sheet'),
          name: z.string(),
          webViewLink: z.string(),
          sourceId: z.string().describe('The .xlsx this was converted from, unchanged'),
          message: z.string(),
        },
        schema: {
          file_id: z.string().describe('Drive file ID of the .xlsx to convert'),
          name: z.string().optional().describe('Name for the new Sheet (defaults to the original name)'),
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/drive.file", wrapHandler(async ({ file_id, name }: any, context: any) => {
          const { accessToken } = context;

          const meta = await fetchDriveFileMeta(file_id, accessToken);

          if (meta.mimeType === NATIVE_SHEET_MIME) {
            throw new Error(
              `'${meta.name}' is already a native Google Sheet and is editable as-is. Nothing to convert.`
            );
          }
          if (meta.mimeType !== XLSX_MIME && meta.mimeType !== XLS_MIME) {
            throw new Error(
              `'${meta.name}' is not an Excel file (${meta.mimeType}), so it cannot be converted to a Google Sheet.`
            );
          }

          const result = await makeDriveRequest(
            `/files/${encodeURIComponent(file_id)}/copy?supportsAllDrives=true`,
            accessToken,
            {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({
                mimeType: NATIVE_SHEET_MIME,
                ...(name ? { name } : {}),
              }),
            }
          ) as { id: string; name: string; webViewLink?: string };

          return toolResponse({
            id: result.id,
            name: result.name,
            webViewLink: result.webViewLink || `https://docs.google.com/spreadsheets/d/${result.id}`,
            sourceId: file_id,
            message:
              `Converted to a new native Google Sheet, which is fully editable. ` +
              `The original .xlsx '${meta.name}' is unchanged.`,
          });
        })),
      },
};
