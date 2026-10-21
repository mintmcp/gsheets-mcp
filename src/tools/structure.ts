/**
 * Tools that change a spreadsheet's shape rather than its cells: rows and
 * columns, tabs copied between files, and the raw batchUpdate passthrough
 * that reaches every request kind the typed tools do not.
 */

import { z } from 'zod';
import { toolResponse, toolError, ApiError } from '../lib/errors.js';
import { columnLetterToIndex, columnIndexToLetter } from '../lib/a1.js';
import { getSheetId, makeSheetsRequest } from '../lib/google.js';
import { applyBatchUpdate } from '../lib/batch.js';
import { isOfficeFileError, explainOfficeFileError } from '../lib/office.js';
import { nativeEdit, spreadsheetIdSchema, sheetNameSchema } from './common.js';

const DIMENSIONS = ['ROWS', 'COLUMNS'] as const;
type Dimension = (typeof DIMENSIONS)[number];

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
 * digit-only string still counts as a row number, since some clients send
 * numbers as strings.
 */
function parseDimensionSpan(dimension: Dimension, start: number | string, count: number): DimensionSpan {
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

/** Requests per batch_update call. Google's own limit is the request body, not a count. */
export const MAX_BATCH_REQUESTS = 500;

/** Serialized request bytes per batch_update call, under Google's 10MB-ish body ceiling. */
export const MAX_BATCH_BYTES = 2 * 1024 * 1024;

const rawRequestSchema = z.record(z.string(), z.unknown())
  .describe('One spreadsheets.batchUpdate Request object with exactly one key naming its kind, e.g. {"addSheet": {"properties": {"title": "New"}}}');

/**
 * Google phrases a bad request as "Invalid requests[3].addSheet: ..." (or
 * "Invalid requests[3]: ..."), so the index and kind come out of the message.
 */
export function failingRequest(message: string): { failingRequestIndex: number; failingRequestKind?: string } | undefined {
  const match = message.match(/requests\[(\d+)\](?:\.([A-Za-z]+))?/);
  if (!match) return undefined;
  return {
    failingRequestIndex: parseInt(match[1], 10),
    ...(match[2] !== undefined && { failingRequestKind: match[2] }),
  };
}

export const structureTools = {
      insert_dimensions: {
        description: 'Insert empty rows or columns at a position, shifting existing ones down or right (batchUpdate insertDimension). By default the new cells inherit formatting from the row above / column to the left of the insertion point, or from the row below / column to the right when inserting at row 1 / column A; set inherit_from_before to choose. For ROWS pass start_index as the 1-based row number the first new row will occupy; for COLUMNS pass the letter of the column the first new column will occupy. Use insert_rows to append data at the end of a tab instead.',
        outputSchema: { id: z.string(), message: z.string() },
        schema: {
          spreadsheet_id: spreadsheetIdSchema,
          sheet_name: sheetNameSchema,
          dimension: z.enum(DIMENSIONS).describe('ROWS or COLUMNS'),
          start_index: z.union([z.number().int(), z.string()])
            .describe('For ROWS, the 1-based row number (e.g. 5). For COLUMNS, the column letter (e.g. "C" or "AA").'),
          count: z.coerce.number().int().min(1).optional().describe('How many rows or columns. Defaults to 1.'),
          inherit_from_before: z.boolean().optional()
            .describe('true copies formatting from the row above / column to the left, false from the row below / column to the right. true is not allowed at row 1 / column A.'),
        },
        handler: nativeEdit(async ({ spreadsheet_id, sheet_name, dimension, start_index, count = 1, inherit_from_before }: any, accessToken) => {
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
            message: `${span.label} inserted in "${sheet_name}"`,
          });
        }),
      },

      copy_sheet_to_spreadsheet: {
        description: 'Copy one tab into another spreadsheet (spreadsheets.sheets.copyTo), with its values, formulas, formatting and charts. The copy is appended to the destination, normally as "Copy of <name>" (Google picks another title if that one is taken; `sheetTitle` returns the real one); rename it there with batch_update updateSheetProperties if needed. Use copy_spreadsheet to copy a whole file instead.',
        outputSchema: {
          id: z.string().describe('The destination spreadsheet'),
          sheetId: z.number().describe('Google\'s id for the new tab in the destination'),
          sheetTitle: z.string(),
          index: z.number().optional().describe('0-based position of the new tab in the destination'),
          message: z.string(),
        },
        schema: {
          spreadsheet_id: spreadsheetIdSchema.describe('Spreadsheet holding the tab to copy'),
          sheet_name: sheetNameSchema.describe('Name of the tab to copy'),
          destination_spreadsheet_id: z.string().describe('Spreadsheet to copy the tab into'),
        },
        handler: nativeEdit(async ({ spreadsheet_id, sheet_name, destination_spreadsheet_id }: any, accessToken) => {
          const sheetId = await getSheetId(spreadsheet_id, sheet_name, accessToken);
          let props: { sheetId?: number; title?: string; index?: number };
          try {
            props = await makeSheetsRequest(
              `/${encodeURIComponent(spreadsheet_id)}/sheets/${sheetId}:copyTo`,
              accessToken,
              { method: 'POST', body: JSON.stringify({ destinationSpreadsheetId: destination_spreadsheet_id }) },
            );
          } catch (err) {
            // The source already read as native above, so an Office refusal here is the destination's
            throw await explainOfficeFileError(err, destination_spreadsheet_id, accessToken);
          }
          if (props.sheetId === undefined || props.title === undefined) {
            throw new Error('Google copied the tab but returned no id or title for it');
          }
          return toolResponse({
            id: destination_spreadsheet_id,
            sheetId: props.sheetId,
            sheetTitle: props.title,
            ...(props.index !== undefined && { index: props.index }),
            message: `Tab "${sheet_name}" copied to spreadsheet ${destination_spreadsheet_id} as "${props.title}"`,
          });
        }),
      },

      batch_update: {
        description: 'Apply raw spreadsheets.batchUpdate requests, for anything the typed tools do not cover: conditional formatting, protected ranges, banding, pivot tables, dimension groups, named ranges, filters, sheet and spreadsheet properties, cut/copy/paste, text to columns, developer metadata and the rest. Each entry of `requests` is one Request object from https://developers.google.com/workspace/sheets/api/reference/rest/v4/spreadsheets/request with exactly one key naming its kind, using Google\'s own field names (sheetId, 0-based GridRange indices with exclusive ends, field masks); get_metadata gives the sheetIds; inspect_spreadsheet gives the current state. Requests are applied in order and all-or-nothing: if any is invalid none is applied, and the error names the failing index. Returns Google\'s replies array, positional with the requests ({} for a request that returns nothing). Destructive requests (deleteSheet, deleteRange, updateCells over existing data) cannot be undone.',
        destructiveHint: true,
        outputSchema: {
          id: z.string(),
          requestCount: z.number(),
          replies: z.array(z.record(z.string(), z.unknown())).describe('One reply per request, in order; {} when the request returns nothing'),
          message: z.string(),
        },
        schema: {
          spreadsheet_id: spreadsheetIdSchema,
          requests: z.array(rawRequestSchema).min(1).max(MAX_BATCH_REQUESTS)
            .describe(`Request objects to apply in order, at most ${MAX_BATCH_REQUESTS}`),
        },
        handler: nativeEdit(async ({ spreadsheet_id, requests }: any, accessToken) => {
          if (!Array.isArray(requests) || requests.length === 0) {
            throw new Error('requests must contain at least one Request object');
          }
          requests.forEach((request: unknown, i: number) => {
            const keys = request && typeof request === 'object' && !Array.isArray(request) ? Object.keys(request) : [];
            if (keys.length !== 1) {
              throw new Error(`requests[${i}] must be an object with exactly one key naming the request kind (e.g. "addSheet"), got ${keys.length === 0 ? 'none' : keys.join(', ')}`);
            }
          });
          const bytes = Buffer.byteLength(JSON.stringify({ requests }));
          if (bytes > MAX_BATCH_BYTES) {
            throw new Error(`requests serialize to ${bytes} bytes, over the ${MAX_BATCH_BYTES} byte cap; split the batch`);
          }

          let replies: Array<Record<string, unknown> | null>;
          try {
            replies = await applyBatchUpdate(spreadsheet_id, requests, accessToken);
          } catch (err) {
            // An uploaded .xlsx is a 400 too; nativeOnly turns it into the convert_to_google_sheet pointer
            if (!(err instanceof ApiError) || err.status !== 400 || isOfficeFileError(err)) throw err;
            return toolError(err.message, {
              status: 400,
              api: err.api,
              code: 'invalid_argument',
              hint: 'No request was applied. Fix the request named and resend the whole batch.',
              ...failingRequest(err.message),
            });
          }

          return toolResponse({
            id: spreadsheet_id,
            requestCount: requests.length,
            replies: replies.map((reply) => reply ?? {}),
            message: `${requests.length} request(s) applied`,
          });
        }),
      },
};
