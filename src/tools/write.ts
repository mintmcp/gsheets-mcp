/**
 * Tools that create spreadsheets or mutate cell values. Cell mutation is
 * native-only (`nativeOnly`); Drive-level operations also accept .xlsx.
 */

import { z } from 'zod';
import { withGoogleAuth as requirePermissionSecure } from '../auth.js';
import { wrapHandler, toolResponse } from '../lib/errors.js';
import { quoteSheetName, assertBareA1Range, assertSingleCell, parseA1Range } from '../lib/a1.js';
import { padRaggedRows } from '../lib/grid.js';
import { makeDriveRequest, makeSheetsRequest, getSheetId } from '../lib/google.js';
import {
  nativeOnly,
  fetchDriveFileMeta,
  XLSX_MIME,
  XLS_MIME,
  NATIVE_SHEET_MIME,
} from '../lib/office.js';

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
        description: 'Append rows AFTER the last non-empty row of a sheet tab (using Sheets values:append with INSERT_ROWS). This tool is append-only — it cannot insert rows at an arbitrary row index, and it cannot insert columns. For mid-sheet writes use update_range with the target A1 range. Values are interpreted as user input (USER_ENTERED), so formulas (e.g. "=SUM(A1:A2)") work automatically — but note: a leading "=" always becomes a formula, and string-typed values like "01" or "1.0" may be coerced (e.g. "01" → 1). Use update_range to overwrite an exact range of existing cells; use this tool when you want to add new rows at the end without specifying a target range.',
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
        handler: requirePermissionSecure("https://www.googleapis.com/auth/spreadsheets", wrapHandler(nativeOnly(async ({ spreadsheet_id, sheet_name, data }: any, context: any) => {
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
        }))),
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
        handler: requirePermissionSecure("https://www.googleapis.com/auth/spreadsheets", wrapHandler(nativeOnly(async ({ spreadsheet_id, sheet_name, cell, content }: any, context: any) => {
          const { accessToken } = context;

          if (!content || content.length === 0) {
            throw new Error('Content must have at least one segment');
          }
          const cleanCell = assertSingleCell(cell);

          const hasUrls = content.some((c: any) => c.url);

          if (!hasUrls) {
            // No hyperlinks: use Values API with USER_ENTERED for auto type detection
            const fullText = content.map((c: any) => c.text).join('');
            const range = `${quoteSheetName(sheet_name)}!${cleanCell}`;
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
        description: 'Overwrite a range of cells with a 2D array (values:PUT). Pass `range` as a bounded bare A1 range — either a single cell like "A1" or a rectangular range like "A1:C3". Whole-column ("A:C") and whole-row ("1:3") forms are not supported. Do NOT include a sheet prefix; use the `sheet_name` argument for that. Sizing rules: for a multi-cell `range` the `data` matrix must fit within the range — the Sheets API rejects oversized matrices with a 400 INVALID_ARGUMENT, and if `data` is smaller than the range only the supplied cells are written (the rest keep their prior values). For a single-cell `range` (e.g. "A1"), the cell acts as a top-left anchor and the matrix expands down and right from it. Ragged rows are padded with empty strings. Values are interpreted as user input (USER_ENTERED): a leading "=" becomes a formula, and string-typed values like "01" may be coerced. Use update_cell for a single cell (especially when you need inline hyperlinks); use insert_rows to add new rows at the end.',
        outputSchema: {
          id: z.string(),
          updatedCells: z.number(),
          message: z.string(),
        },
        schema: {
          spreadsheet_id: z.string().describe('Google Sheets spreadsheet ID'),
          sheet_name: z.string().describe('Name of the sheet tab'),
          range: z.string().describe('Bounded A1 range: "A1" or "A1:C3". Whole-column/whole-row forms are not supported.'),
          data: z.array(z.array(z.string())).describe('2D array of values. Formulas like "=SUM(A1:A2)" are supported.'),
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/spreadsheets", wrapHandler(nativeOnly(async ({ spreadsheet_id, sheet_name, range, data }: any, context: any) => {
          const { accessToken } = context;

          const cleanRange = assertBareA1Range(range);
          const paddedData = padRaggedRows(data);

          const a1Range = `${quoteSheetName(sheet_name)}!${cleanRange}`;
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
            message: `Range ${cleanRange} updated (${result.updatedCells || 0} cells)`,
          });
        }))),
      },

      clear_values: {
        description: 'Clear cell values from one or more ranges in a sheet tab. Each range must be a bounded bare A1 range — either a single cell like "A1" or a rectangular range like "A1:B5". Whole-column ("A:C") and whole-row ("1:3") forms are not supported. Do NOT include a sheet prefix; use the `sheet_name` argument for that. Only values are cleared; formatting is preserved. Use clear_formatting to reset visual styling instead.',
        destructiveHint: true,
        outputSchema: {
          id: z.string(),
          clearedRanges: z.array(z.string()),
          message: z.string(),
        },
        schema: {
          spreadsheet_id: z.string().describe('Google Sheets spreadsheet ID'),
          sheet_name: z.string().describe('Name of the sheet tab'),
          ranges: z.array(z.string()).min(1).describe('Array of bounded A1 ranges to clear (e.g. ["A1:B5", "D1:D10"]). Whole-column/whole-row forms are not supported.'),
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/spreadsheets", wrapHandler(nativeOnly(async ({ spreadsheet_id, sheet_name, ranges }: any, context: any) => {
          const { accessToken } = context;

          if (!Array.isArray(ranges) || ranges.length === 0) {
            throw new Error('ranges must contain at least one A1 range');
          }
          const cleanRanges = ranges.map((r: unknown) => assertBareA1Range(r, 'ranges[]'));

          const qualifiedRanges = cleanRanges.map((r: string) => `${quoteSheetName(sheet_name)}!${r}`);

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
            message: `Cleared ${cleanRanges.length} range(s)`,
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
