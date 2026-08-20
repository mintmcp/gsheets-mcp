/**
 * Tools that change cell appearance rather than content. Formatting is a
 * Sheets-API operation, so every tool here is native-only.
 */

import { z } from 'zod';
import { withGoogleAuth as requirePermissionSecure } from '../auth.js';
import { wrapHandler, toolResponse } from '../lib/errors.js';
import { assertBareA1Range, parseA1Range } from '../lib/a1.js';
import { parseColor } from '../lib/color.js';
import { makeSheetsRequest, getSheetId } from '../lib/google.js';
import { nativeOnly } from '../lib/office.js';

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

          const cleanRange = assertBareA1Range(range);
          const sheetId = await getSheetId(spreadsheet_id, sheet_name, accessToken);
          const gridRange = parseA1Range(cleanRange);

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
            message: `Formatting applied to ${cleanRange}`,
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

          const cleanRange = assertBareA1Range(range);
          const sheetId = await getSheetId(spreadsheet_id, sheet_name, accessToken);
          const gridRange = parseA1Range(cleanRange);

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
            message: `Formatting cleared from ${cleanRange}`,
          });
        }))),
      },
};
