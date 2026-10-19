/**
 * Tools that change cell appearance rather than content. Formatting is a
 * Sheets-API operation, so every tool here is native-only.
 */

import { z } from 'zod';
import { withGoogleAuth as requirePermissionSecure } from '../auth.js';
import { wrapHandler, toolResponse } from '../lib/errors.js';
import { assertBareA1Range, parseA1Range } from '../lib/a1.js';
import { parseColor } from '../lib/color.js';
import { resolveGridRange } from '../lib/google.js';
import { applyBatchUpdate } from '../lib/batch.js';
import { nativeOnly } from '../lib/office.js';

const BORDER_STYLES = ['NONE', 'SOLID', 'SOLID_MEDIUM', 'SOLID_THICK', 'DOTTED', 'DASHED', 'DOUBLE'] as const;

const BORDER_SIDES = ['top', 'bottom', 'left', 'right', 'innerHorizontal', 'innerVertical'] as const;
type BorderSide = (typeof BORDER_SIDES)[number];

/** Named groups of sides, resolved before the request is built. */
const SIDE_GROUPS: Record<string, readonly BorderSide[]> = {
  all: BORDER_SIDES,
  outer: ['top', 'bottom', 'left', 'right'],
  inner: ['innerHorizontal', 'innerVertical'],
};

const SIDE_NAMES = [...BORDER_SIDES, ...Object.keys(SIDE_GROUPS)];

function resolveSides(sides: unknown): Set<BorderSide> {
  const named = sides === undefined ? ['all'] : sides;
  if (!Array.isArray(named) || named.length === 0) {
    throw new Error(`sides must be a non-empty list drawn from ${SIDE_NAMES.join(', ')}`);
  }
  const resolved = new Set<BorderSide>();
  for (const side of named) {
    if ((BORDER_SIDES as readonly unknown[]).includes(side)) resolved.add(side as BorderSide);
    else if (typeof side === 'string' && side in SIDE_GROUPS) SIDE_GROUPS[side].forEach((s) => resolved.add(s));
    else throw new Error(`Unknown side ${JSON.stringify(side)}; use ${SIDE_NAMES.join(', ')}`);
  }
  return resolved;
}

const spreadsheetId = z.string().describe('Google Sheets spreadsheet ID');
const sheetName = z.string().describe('Name of the sheet tab');
const boundedRange = z.string().describe('Bounded A1 range: "A1" or "A1:C3". Whole-column/whole-row forms are not supported.');
const idAndMessage = { id: z.string(), message: z.string() };

export const formatTools = {
      format_cells: {
        description: 'Apply formatting uniformly to every cell in a range (the same `format` object is applied to all cells — there is no per-cell variation in a single call; call this tool multiple times with different ranges to vary formatting). Pass `range` as a bounded bare A1 range — either "A1" or "A1:C3". Whole-column ("A:C") and whole-row ("1:3") forms are not supported. Do NOT include a sheet prefix; use the `sheet_name` argument for that. Supports background color, text formatting (bold, italic, font size, font family, foreground color), horizontal/vertical alignment, wrap strategy, and number format. Colors accept either hex strings (e.g. "#FF0000", "#F00") or {red,green,blue} float objects (0..1). Use clear_formatting to reset styling.',
        outputSchema: {
          id: z.string(),
          message: z.string(),
        },
        schema: {
          spreadsheet_id: z.string().describe('Google Sheets spreadsheet ID'),
          sheet_name: z.string().describe('Name of the sheet tab'),
          range: z.string().describe('Bounded A1 range: "A1" or "A1:C3". Whole-column/whole-row forms are not supported.'),
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
            verticalAlignment: z.enum(['TOP', 'MIDDLE', 'BOTTOM']).optional().describe('Vertical alignment'),
            wrapStrategy: z.enum(['OVERFLOW_CELL', 'CLIP', 'WRAP']).optional().describe('Text wrap strategy'),
            numberFormat: z.object({
              type: z.enum(['TEXT', 'NUMBER', 'PERCENT', 'CURRENCY', 'DATE', 'TIME', 'DATE_TIME', 'SCIENTIFIC']),
              pattern: z.string().optional().describe('Format pattern (e.g. "#,##0.00", "yyyy-mm-dd")'),
            }).optional().describe('Number format'),
          }).describe('Formatting options to apply'),
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/spreadsheets", wrapHandler(nativeOnly(async ({ spreadsheet_id, sheet_name, range, format }: any, context: any) => {
          const { accessToken } = context;

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
          if (format.verticalAlignment) {
            cellFormat.verticalAlignment = format.verticalAlignment;
            fields.push('userEnteredFormat.verticalAlignment');
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
            throw new Error('format must include at least one of: backgroundColor, textFormat, horizontalAlignment, verticalAlignment, wrapStrategy, numberFormat');
          }

          const { a1, grid } = await resolveGridRange(spreadsheet_id, sheet_name, range, accessToken);
          await applyBatchUpdate(spreadsheet_id, [{
            repeatCell: { range: grid, cell: { userEnteredFormat: cellFormat }, fields: fields.join(',') },
          }], accessToken);

          return toolResponse({
            id: spreadsheet_id,
            message: `Formatting applied to ${a1}`,
          });
        }))),
      },

      clear_formatting: {
        description: 'Clear all formatting from a range, resetting cells to default appearance. Pass `range` as a bounded bare A1 range — either "A1" or "A1:C3". Whole-column ("A:C") and whole-row ("1:3") forms are not supported. Do NOT include a sheet prefix; use the `sheet_name` argument for that. Cell values are preserved. Use clear_values to clear cell contents instead.',
        destructiveHint: true,
        outputSchema: {
          id: z.string(),
          message: z.string(),
        },
        schema: {
          spreadsheet_id: z.string().describe('Google Sheets spreadsheet ID'),
          sheet_name: z.string().describe('Name of the sheet tab'),
          range: z.string().describe('Bounded A1 range: "A1" or "A1:C3". Whole-column/whole-row forms are not supported.'),
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/spreadsheets", wrapHandler(nativeOnly(async ({ spreadsheet_id, sheet_name, range }: any, context: any) => {
          const { accessToken } = context;

          const { a1, grid } = await resolveGridRange(spreadsheet_id, sheet_name, range, accessToken);
          await applyBatchUpdate(spreadsheet_id, [{
            repeatCell: { range: grid, cell: { userEnteredFormat: {} }, fields: 'userEnteredFormat' },
          }], accessToken);

          return toolResponse({
            id: spreadsheet_id,
            message: `Formatting cleared from ${a1}`,
          });
        }))),
      },

      set_borders: {
        description: 'Draw or remove cell borders on a range. One style and colour is applied to the chosen sides: top, bottom, left, right (the outline of the range), innerHorizontal and innerVertical (the lines between cells inside it), or the shorthands outer, inner and all (the default). Sides not named are left as they are. Style NONE removes borders from the named sides. Call once per distinct look, e.g. a SOLID_THICK outer frame and a SOLID inner grid in two calls.',
        outputSchema: idAndMessage,
        schema: {
          spreadsheet_id: spreadsheetId,
          sheet_name: sheetName,
          range: boundedRange,
          style: z.enum(BORDER_STYLES).describe('Line style; NONE removes the border'),
          color: z.string().optional().describe('Hex colour like "#000000" (default black)'),
          sides: z.array(z.enum(SIDE_NAMES as [string, ...string[]])).optional().describe('Which sides to change (default ["all"])'),
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/spreadsheets", wrapHandler(nativeOnly(async ({ spreadsheet_id, sheet_name, range, style, color, sides }: any, context: any) => {
          const { accessToken } = context;
          if (!(BORDER_STYLES as readonly unknown[]).includes(style)) {
            throw new Error(`style must be one of ${BORDER_STYLES.join(', ')}, got ${JSON.stringify(style)}`);
          }
          const resolved = resolveSides(sides);
          const parsed = color !== undefined ? parseColor(color) : undefined;
          const border = { style, ...(parsed && { color: parsed }) };
          const { a1, grid } = await resolveGridRange(spreadsheet_id, sheet_name, range, accessToken);
          await applyBatchUpdate(spreadsheet_id, [{
            updateBorders: { range: grid, ...Object.fromEntries([...resolved].map((side) => [side, border])) },
          }], accessToken);
          return toolResponse({
            id: spreadsheet_id,
            message: `${style === 'NONE' ? 'Borders removed from' : `${style} borders set on`} ${[...resolved].join(', ')} of ${a1}`,
          });
        }))),
      },

      merge_cells: {
        description: 'Merge a range into one cell (MERGE_ALL, the default), or into one cell per row (MERGE_ROWS) or per column (MERGE_COLUMNS). Only the top-left value survives a merge; the others are discarded. Use unmerge_cells to undo.',
        destructiveHint: true,
        outputSchema: idAndMessage,
        schema: {
          spreadsheet_id: spreadsheetId,
          sheet_name: sheetName,
          range: boundedRange,
          merge_type: z.enum(['MERGE_ALL', 'MERGE_ROWS', 'MERGE_COLUMNS']).optional().describe('Default MERGE_ALL'),
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/spreadsheets", wrapHandler(nativeOnly(async ({ spreadsheet_id, sheet_name, range, merge_type = 'MERGE_ALL' }: any, context: any) => {
          const { accessToken } = context;
          const rect = parseA1Range(assertBareA1Range(range));
          if (rect.endRowIndex - rect.startRowIndex === 1 && rect.endColumnIndex - rect.startColumnIndex === 1) {
            throw new Error(`range "${range}" is a single cell; pass a range of two or more cells to merge`);
          }
          const { a1, grid } = await resolveGridRange(spreadsheet_id, sheet_name, range, accessToken);
          await applyBatchUpdate(spreadsheet_id, [{ mergeCells: { range: grid, mergeType: merge_type } }], accessToken);
          return toolResponse({
            id: spreadsheet_id,
            message: `Merged ${a1} on "${sheet_name}" (${merge_type})`,
          });
        }))),
      },

      unmerge_cells: {
        description: 'Split every merged cell that lies within a range back into individual cells. The merged value stays in the top-left cell.',
        outputSchema: idAndMessage,
        schema: {
          spreadsheet_id: spreadsheetId,
          sheet_name: sheetName,
          range: boundedRange,
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/spreadsheets", wrapHandler(nativeOnly(async ({ spreadsheet_id, sheet_name, range }: any, context: any) => {
          const { accessToken } = context;
          const { a1, grid } = await resolveGridRange(spreadsheet_id, sheet_name, range, accessToken);
          await applyBatchUpdate(spreadsheet_id, [{ unmergeCells: { range: grid } }], accessToken);
          return toolResponse({
            id: spreadsheet_id,
            message: `Unmerged ${a1} on "${sheet_name}"`,
          });
        }))),
      },
};
