/**
 * Tools that change a spreadsheet's shape rather than its cells: tabs, rows,
 * columns, frozen panes, sort order and find/replace. All of them are a
 * single `batchUpdate` request, so all are native-only.
 */

import { z } from 'zod';
import { withGoogleAuth as requirePermissionSecure } from '../auth.js';
import { wrapHandler, toolResponse } from '../lib/errors.js';
import {
  assertBareA1Range,
  parseA1Range,
  columnLetterToIndex,
  columnIndexToLetter,
  a1Range,
} from '../lib/a1.js';
import { getSheetId, getSheetPosition } from '../lib/google.js';
import { applyBatchUpdate } from '../lib/batch.js';
import { nativeOnly } from '../lib/office.js';

const DIMENSIONS = ['ROWS', 'COLUMNS'] as const;
type Dimension = (typeof DIMENSIONS)[number];

const dimensionSchema = z.enum(DIMENSIONS).describe('ROWS or COLUMNS');

const startIndexSchema = z.union([z.number().int(), z.string()])
  .describe('For ROWS, the 1-based row number (e.g. 5). For COLUMNS, the column letter (e.g. "C" or "AA").');

const countSchema = z.coerce.number().int().min(1).optional()
  .describe('How many rows or columns. Defaults to 1.');

interface DimensionSpan {
  /** 0-based, inclusive. */
  startIndex: number;
  /** 0-based, exclusive. */
  endIndex: number;
  /** "row 5", "rows 5-6", "column C" or "columns C-D", as the caller named it. */
  label: string;
}

/**
 * Rows are addressed by number and columns by letter, as in A1 notation, so
 * the two forms are checked against the dimension rather than coerced. A
 * digit-only string still counts as a row number, as with the coerced numeric
 * inputs elsewhere, since some clients send numbers as strings.
 */
function parseDimensionSpan(dimension: Dimension, start: number | string, count: number): DimensionSpan {
  if (!Number.isInteger(count) || count < 1) {
    throw new Error(`count must be a positive integer, got ${JSON.stringify(count)}`);
  }
  if (dimension === 'ROWS') {
    const row = typeof start === 'string' && /^\d+$/.test(start) ? Number(start) : start;
    if (typeof row !== 'number' || !Number.isInteger(row) || row < 1) {
      throw new Error(`start_index must be a 1-based row number for ROWS, got ${JSON.stringify(start)}`);
    }
    const startIndex = row - 1;
    const endIndex = startIndex + count;
    const label = count === 1 ? `row ${row}` : `rows ${row}-${endIndex}`;
    return { startIndex, endIndex, label };
  }
  if (typeof start !== 'string' || !/^[A-Za-z]+$/.test(start)) {
    throw new Error(`start_index must be a column letter like "C" for COLUMNS, got ${JSON.stringify(start)}`);
  }
  const startIndex = columnLetterToIndex(start);
  const endIndex = startIndex + count;
  const letters = start.toUpperCase();
  const label = count === 1
    ? `column ${letters}`
    : `columns ${letters}-${columnIndexToLetter(endIndex - 1)}`;
  return { startIndex, endIndex, label };
}

const spreadsheetId = z.string().describe('Google Sheets spreadsheet ID');
const sheetName = z.string().describe('Name of the sheet tab');

const idAndMessage = {
  id: z.string(),
  message: z.string(),
};

export const structureTools = {
      rename_sheet: {
        description: 'Rename a sheet tab. Other tabs, formulas that reference the tab, and its contents are unaffected.',
        outputSchema: {
          id: z.string(),
          sheetTitle: z.string().describe('The new tab name'),
          message: z.string(),
        },
        schema: {
          spreadsheet_id: spreadsheetId,
          sheet_name: z.string().describe('Current name of the tab'),
          new_name: z.string().min(1).describe('New name for the tab'),
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/spreadsheets", wrapHandler(nativeOnly(async ({ spreadsheet_id, sheet_name, new_name }: any, context: any) => {
          const { accessToken } = context;
          const sheetId = await getSheetId(spreadsheet_id, sheet_name, accessToken);
          await applyBatchUpdate(spreadsheet_id, [{
            updateSheetProperties: { properties: { sheetId, title: new_name }, fields: 'title' },
          }], accessToken);
          return toolResponse({
            id: spreadsheet_id,
            sheetTitle: new_name,
            message: `Sheet tab "${sheet_name}" renamed to "${new_name}"`,
          });
        }))),
      },

      delete_sheet: {
        description: 'Delete a sheet tab and everything on it. This cannot be undone through the API, and a spreadsheet must keep at least one tab, so Google refuses to delete the last one.',
        destructiveHint: true,
        outputSchema: idAndMessage,
        schema: {
          spreadsheet_id: spreadsheetId,
          sheet_name: z.string().describe('Name of the tab to delete'),
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/spreadsheets", wrapHandler(nativeOnly(async ({ spreadsheet_id, sheet_name }: any, context: any) => {
          const { accessToken } = context;
          const sheetId = await getSheetId(spreadsheet_id, sheet_name, accessToken);
          await applyBatchUpdate(spreadsheet_id, [{ deleteSheet: { sheetId } }], accessToken);
          return toolResponse({
            id: spreadsheet_id,
            message: `Sheet tab "${sheet_name}" deleted`,
          });
        }))),
      },

      duplicate_sheet: {
        description: 'Copy a sheet tab within the same spreadsheet, including its values, formulas, formatting and charts. The copy is placed right after the original. Use copy_spreadsheet to copy a whole file instead.',
        outputSchema: {
          id: z.string(),
          sheetTitle: z.string().describe('Name of the new tab, as assigned by Google when new_name is omitted'),
          message: z.string(),
        },
        schema: {
          spreadsheet_id: spreadsheetId,
          sheet_name: z.string().describe('Name of the tab to copy'),
          new_name: z.string().min(1).optional().describe('Name for the copy (defaults to "Copy of <original>")'),
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/spreadsheets", wrapHandler(nativeOnly(async ({ spreadsheet_id, sheet_name, new_name }: any, context: any) => {
          const { accessToken } = context;
          const source = await getSheetPosition(spreadsheet_id, sheet_name, accessToken);
          const [reply] = await applyBatchUpdate<{ duplicateSheet?: { properties?: { title?: string } } }>(
            spreadsheet_id,
            [{
              duplicateSheet: {
                sourceSheetId: source.sheetId,
                insertSheetIndex: source.index + 1,
                ...(new_name && { newSheetName: new_name }),
              },
            }],
            accessToken,
          );
          const sheetTitle = reply?.duplicateSheet?.properties?.title ?? new_name ?? `Copy of ${sheet_name}`;
          return toolResponse({
            id: spreadsheet_id,
            sheetTitle,
            message: `Sheet tab "${sheet_name}" duplicated as "${sheetTitle}"`,
          });
        }))),
      },

      insert_rows_or_columns: {
        description: 'Insert empty rows or columns at a position, shifting existing ones down or right. By default the new cells inherit formatting from the row above / column to the left of the insertion point, or from the row below / column to the right when inserting at row 1 / column A; set inherit_from_before to choose. For ROWS pass start_index as the 1-based row number the first new row will occupy; for COLUMNS pass the letter of the column the first new column will occupy. Use insert_rows to append data at the end of a tab instead.',
        outputSchema: idAndMessage,
        schema: {
          spreadsheet_id: spreadsheetId,
          sheet_name: sheetName,
          dimension: dimensionSchema,
          start_index: startIndexSchema,
          count: countSchema,
          inherit_from_before: z.boolean().optional()
            .describe('true copies formatting from the row above / column to the left, false from the row below / column to the right. true is not allowed at row 1 / column A.'),
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/spreadsheets", wrapHandler(nativeOnly(async ({ spreadsheet_id, sheet_name, dimension, start_index, count = 1, inherit_from_before }: any, context: any) => {
          const { accessToken } = context;
          const span = parseDimensionSpan(dimension, start_index, count);
          // Google refuses inheriting from before the first row or column
          if (inherit_from_before === true && span.startIndex === 0) {
            throw new Error(`inherit_from_before cannot be true at ${span.label}: there is nothing before it to inherit from`);
          }
          const sheetId = await getSheetId(spreadsheet_id, sheet_name, accessToken);
          await applyBatchUpdate(spreadsheet_id, [{
            insertDimension: {
              range: { sheetId, dimension, startIndex: span.startIndex, endIndex: span.endIndex },
              inheritFromBefore: inherit_from_before ?? span.startIndex > 0,
            },
          }], accessToken);
          return toolResponse({
            id: spreadsheet_id,
            message: `${span.label} inserted`,
          });
        }))),
      },

      delete_rows_or_columns: {
        description: 'Delete whole rows or columns, shifting the ones after them up or left. For ROWS pass start_index as the 1-based number of the first row to delete; for COLUMNS pass the letter of the first column to delete. Use clear_values to empty cells without removing them.',
        destructiveHint: true,
        outputSchema: idAndMessage,
        schema: {
          spreadsheet_id: spreadsheetId,
          sheet_name: sheetName,
          dimension: dimensionSchema,
          start_index: startIndexSchema,
          count: countSchema,
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/spreadsheets", wrapHandler(nativeOnly(async ({ spreadsheet_id, sheet_name, dimension, start_index, count = 1 }: any, context: any) => {
          const { accessToken } = context;
          const span = parseDimensionSpan(dimension, start_index, count);
          const sheetId = await getSheetId(spreadsheet_id, sheet_name, accessToken);
          await applyBatchUpdate(spreadsheet_id, [{
            deleteDimension: {
              range: { sheetId, dimension, startIndex: span.startIndex, endIndex: span.endIndex },
            },
          }], accessToken);
          return toolResponse({
            id: spreadsheet_id,
            message: `${span.label} deleted`,
          });
        }))),
      },

      freeze_panes: {
        description: 'Freeze the first N rows and/or columns of a tab so they stay visible while scrolling. Pass 0 to unfreeze. A count you omit is left as it is.',
        outputSchema: idAndMessage,
        schema: {
          spreadsheet_id: spreadsheetId,
          sheet_name: sheetName,
          rows: z.coerce.number().int().min(0).optional().describe('Number of rows to freeze at the top (0 unfreezes)'),
          columns: z.coerce.number().int().min(0).optional().describe('Number of columns to freeze at the left (0 unfreezes)'),
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/spreadsheets", wrapHandler(nativeOnly(async ({ spreadsheet_id, sheet_name, rows, columns }: any, context: any) => {
          const { accessToken } = context;
          const gridProperties: Record<string, number> = {};
          const fields: string[] = [];
          if (rows !== undefined) {
            gridProperties.frozenRowCount = rows;
            fields.push('gridProperties.frozenRowCount');
          }
          if (columns !== undefined) {
            gridProperties.frozenColumnCount = columns;
            fields.push('gridProperties.frozenColumnCount');
          }
          if (fields.length === 0) {
            throw new Error('Pass rows and/or columns: the number to freeze, or 0 to unfreeze');
          }
          const sheetId = await getSheetId(spreadsheet_id, sheet_name, accessToken);
          await applyBatchUpdate(spreadsheet_id, [{
            updateSheetProperties: { properties: { sheetId, gridProperties }, fields: fields.join(',') },
          }], accessToken);
          const parts = [
            ...(rows !== undefined ? [`${rows} row(s)`] : []),
            ...(columns !== undefined ? [`${columns} column(s)`] : []),
          ];
          return toolResponse({
            id: spreadsheet_id,
            message: `Frozen ${parts.join(' and ')} on "${sheet_name}"`,
          });
        }))),
      },

      sort_range: {
        description: 'Sort the rows of a range in place by one or more of its columns. Pass `range` as a bounded bare A1 range ("A1:D50"); whole-column forms are not supported. Columns in `sort_by` are named by letter and must fall inside the range; the first entry is the primary key. Set has_header to keep the first row of the range in place. Only the cells inside the range move, so include every column of the table or rows will be scrambled relative to the columns left out.',
        destructiveHint: true,
        outputSchema: idAndMessage,
        schema: {
          spreadsheet_id: spreadsheetId,
          sheet_name: sheetName,
          range: z.string().describe('Bounded A1 range to sort, e.g. "A1:D50"'),
          sort_by: z.array(z.object({
            column: z.string().describe('Column letter, e.g. "B"'),
            descending: z.boolean().optional().describe('Sort this key descending (default ascending)'),
          })).min(1).describe('Sort keys, primary first'),
          has_header: z.boolean().optional().describe('The first row of the range is a header and stays in place'),
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/spreadsheets", wrapHandler(nativeOnly(async ({ spreadsheet_id, sheet_name, range, sort_by, has_header }: any, context: any) => {
          const { accessToken } = context;
          const cleanRange = assertBareA1Range(range);
          const rect = parseA1Range(cleanRange);
          const grid = has_header ? { ...rect, startRowIndex: rect.startRowIndex + 1 } : rect;
          if (grid.startRowIndex >= grid.endRowIndex) {
            throw new Error(`range "${cleanRange}" has no data rows to sort below its header`);
          }

          const sortSpecs = sort_by.map(({ column, descending }: { column: string; descending?: boolean }) => {
            if (!/^[A-Za-z]+$/.test(column)) {
              throw new Error(`sort_by column must be a letter like "B", got ${JSON.stringify(column)}`);
            }
            const dimensionIndex = columnLetterToIndex(column);
            if (dimensionIndex < grid.startColumnIndex || dimensionIndex >= grid.endColumnIndex) {
              throw new Error(`sort_by column ${column.toUpperCase()} is outside the range ${cleanRange}`);
            }
            return { dimensionIndex, sortOrder: descending ? 'DESCENDING' : 'ASCENDING' };
          });

          const sheetId = await getSheetId(spreadsheet_id, sheet_name, accessToken);
          await applyBatchUpdate(spreadsheet_id, [{
            sortRange: { range: { sheetId, ...grid }, sortSpecs },
          }], accessToken);
          const sorted = a1Range(
            grid.startColumnIndex, grid.startRowIndex + 1, grid.endColumnIndex - 1, grid.endRowIndex,
          );
          return toolResponse({
            id: spreadsheet_id,
            message: `Sorted ${sorted} by ${sort_by.map((s: { column: string }) => s.column.toUpperCase()).join(', ')}`,
          });
        }))),
      },

      find_replace: {
        description: 'Find and replace text across the whole spreadsheet, one tab, or a range within a tab. Matching is case-insensitive and substring by default; set match_entire_cell for whole-cell matches and use_regex to treat `find` as a regular expression (Java regex syntax, with $1-style group references in `replacement`). Formulas are left alone unless include_formulas is set. Returns how many cells changed. There is no dry run, so read the range first if you need to know what will match.',
        destructiveHint: true,
        outputSchema: {
          id: z.string(),
          occurrencesChanged: z.number().describe('Total matches replaced'),
          valuesChanged: z.number().describe('Plain-value cells changed'),
          formulasChanged: z.number().describe('Formula cells changed'),
          rowsChanged: z.number(),
          sheetsChanged: z.number(),
          message: z.string(),
        },
        schema: {
          spreadsheet_id: spreadsheetId,
          find: z.string().min(1).describe('Text (or regex when use_regex) to search for'),
          replacement: z.string().describe('Replacement text; may be empty to delete matches'),
          sheet_name: z.string().optional().describe('Limit to this tab. Omit to search every tab.'),
          range: z.string().optional().describe('Bounded A1 range within sheet_name, e.g. "A1:C100". Requires sheet_name.'),
          match_case: z.boolean().optional().describe('Case-sensitive match (default false)'),
          match_entire_cell: z.boolean().optional().describe('Only cells whose whole content matches (default false)'),
          use_regex: z.boolean().optional().describe('Treat `find` as a regular expression (default false)'),
          include_formulas: z.boolean().optional().describe('Also search inside formula text (default false)'),
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/spreadsheets", wrapHandler(nativeOnly(async ({
          spreadsheet_id, find, replacement, sheet_name, range,
          match_case = false, match_entire_cell = false, use_regex = false, include_formulas = false,
        }: any, context: any) => {
          const { accessToken } = context;
          if (range !== undefined && sheet_name === undefined) {
            throw new Error('range requires sheet_name: pass the tab the range lives on');
          }
          const cleanRange = range !== undefined ? assertBareA1Range(range) : undefined;

          let scope: object = { allSheets: true };
          if (sheet_name !== undefined) {
            const sheetId = await getSheetId(spreadsheet_id, sheet_name, accessToken);
            scope = cleanRange !== undefined
              ? { range: { sheetId, ...parseA1Range(cleanRange) } }
              : { sheetId };
          }

          const [reply] = await applyBatchUpdate<{ findReplace?: Record<string, number> }>(
            spreadsheet_id,
            [{
              findReplace: {
                find,
                replacement,
                ...scope,
                matchCase: match_case,
                matchEntireCell: match_entire_cell,
                searchByRegex: use_regex,
                includeFormulas: include_formulas,
              },
            }],
            accessToken,
          );
          const counts = reply?.findReplace ?? {};
          const occurrencesChanged = counts.occurrencesChanged ?? 0;
          const where = cleanRange !== undefined
            ? `${cleanRange} on "${sheet_name}"`
            : sheet_name !== undefined ? `"${sheet_name}"` : 'every tab';
          return toolResponse({
            id: spreadsheet_id,
            occurrencesChanged,
            valuesChanged: counts.valuesChanged ?? 0,
            formulasChanged: counts.formulasChanged ?? 0,
            rowsChanged: counts.rowsChanged ?? 0,
            sheetsChanged: counts.sheetsChanged ?? 0,
            message: `Replaced ${occurrencesChanged} occurrence(s) of ${JSON.stringify(find)} in ${where}`,
          });
        }))),
      },
};
