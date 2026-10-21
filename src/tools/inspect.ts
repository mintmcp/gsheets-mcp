/**
 * The raw document, for callers who know the Sheets API: spreadsheets.get
 * with the caller's own field mask, as the official connector exposes it.
 * get_metadata stays the curated view; this is the escape hatch that reaches
 * conditional formats, protected ranges, pivot tables, cell formats and
 * anything else the typed tools do not surface.
 */

import { z } from 'zod';
import { toolResponse } from '../lib/errors.js';
import { makeSheetsRequest, ResponseTooLargeError } from '../lib/google.js';
import { normalizeRange, quoteSheetName } from '../lib/a1.js';
import { nativeTool, spreadsheetIdSchema } from './common.js';
import { SCOPES } from '../scopes.js';

/**
 * Ceiling on the serialized document. Grid data runs about 700 bytes per
 * cell with formats, so this is roughly 3,000 cells: enough for a range
 * inspection, far short of a whole-sheet dump.
 */
export const MAX_RAW_BYTES = 2 * 1024 * 1024;

/**
 * spreadsheets.get takes A1 ranges. A string with "!", ":" or "[" is a range:
 * it is translated (so R1C1 works) and re-quoted, and its parse errors
 * stand. Otherwise a quoted name is a tab, a string that parses as a cell
 * is that cell on the first tab, as Google reads it, a plain identifier
 * passes through for Google to resolve as a tab or named range, and anything
 * else is quoted as a tab name. A tab named like a cell ("Q1") has to be quoted.
 */
export function rawRange(input: string, paramName: string): string {
  const quoted = input.match(/^'(.*)'$/);
  if (quoted) return quoteSheetName(quoted[1].replace(/''/g, "'"));
  let parsed;
  try {
    parsed = normalizeRange(input, { allowOpen: true, paramName });
  } catch (err) {
    if (/[!:[]/.test(input)) throw err;
    // Unquoted, Google resolves a plain name as a tab or a named range
    return /^[A-Za-z_][A-Za-z0-9_]*$/.test(input) ? input : quoteSheetName(input);
  }
  return parsed.sheetName === undefined ? parsed.a1 : `${quoteSheetName(parsed.sheetName)}!${parsed.a1}`;
}

export const inspectTools = {
      inspect_spreadsheet: {
        description: 'Return the raw Spreadsheet resource (spreadsheets.get) filtered by a field mask you choose, as the Sheets API documents it: sheetIds, conditional formats, protected ranges, banded ranges, row/column groups, merges, cell formats and formulas (under sheets.data), pivot tables, named ranges, and so on. `fields` is required and uses the response\'s hierarchical paths, e.g. ["sheets.properties", "sheets.conditionalFormats", "sheets.protectedRanges", "namedRanges"]. Grid data is returned exactly when the mask reaches under sheets.data: "sheets.data.rowData.values.userEnteredFormat" with `ranges` ["Data!A1:D10"] reads the cell formats of one block (Google ignores includeGridData once a mask is set, so there is no separate switch). Keep masks narrow: the answer is capped at 2MB and grid data costs about 700 bytes per cell, so never ask for sheets.data without ranges. Comment threads are part of the resource: mask "comments" for the threads (text, author, status, replies, quoted cell) and "sheets.commentAnchors" for the cell each one sits on, matched by anchorId. Google leaves both out for a caller without comment access on the file. Use get_metadata for the curated tab list and get_sheet_data for values.',
        readOnlyHint: true,
        outputSchema: {
          id: z.string(),
          spreadsheet: z.record(z.string(), z.unknown()).describe('The Spreadsheet resource, masked by `fields`'),
          bytes: z.number().describe('Serialized size of `spreadsheet`, against the 2MB cap'),
        },
        schema: {
          spreadsheet_id: spreadsheetIdSchema,
          fields: z.union([z.array(z.string()), z.string()])
            .describe('Field mask paths, as an array or one comma-separated string, e.g. ["sheets.properties.sheetId", "sheets.properties.title"]. Required.'),
          ranges: z.array(z.string()).optional()
            .describe('Limit grid data to these ranges: "Data!A1:D10", R1C1 "Data!R1C1:R10C4", whole columns "Data!A:C", or a tab or named-range name (quote a tab name that looks like a cell or contains ":", "!" or "[": "\'Q1\'"). Only meaningful with a mask under sheets.data.'),
        },
        handler: nativeTool(SCOPES.SPREADSHEETS, async ({ spreadsheet_id, fields, ranges }: any, accessToken) => {
          const mask = (Array.isArray(fields) ? fields : String(fields).split(','))
            .map((f: unknown) => String(f).trim())
            .filter((f: string) => f.length > 0);
          if (mask.length === 0) {
            throw new Error('fields must name at least one path to return, e.g. ["sheets.properties"]; an unmasked spreadsheets.get is a grid dump');
          }

          // Without a comments mode Google drops comments and commentAnchors whatever the mask says
          const params = new URLSearchParams({ fields: mask.join(','), commentsViewMode: 'COMMENTS_VIEW_MODE_DEFAULT_FOR_CURRENT_ACCESS' });
          (ranges ?? []).forEach((range: unknown, i: number) => params.append('ranges', rawRange(String(range), `ranges[${i}]`)));

          let spreadsheet: Record<string, unknown>;
          try {
            spreadsheet = await makeSheetsRequest(
              `/${encodeURIComponent(spreadsheet_id)}?${params}`,
              accessToken,
              { method: 'GET', maxBytes: MAX_RAW_BYTES },
            );
          } catch (err) {
            if (!(err instanceof ResponseTooLargeError)) throw err;
            const narrow = ranges?.length ? 'or shrink `ranges`' : 'or add `ranges` to a sheets.data mask';
            throw new Error(
              `The masked document is over the ${Math.round(MAX_RAW_BYTES / 1024 / 1024)}MB cap. Narrow \`fields\` to the paths you need, ${narrow}; use get_sheet_data to read values.`,
            );
          }

          return toolResponse({
            id: spreadsheet_id,
            spreadsheet,
            bytes: Buffer.byteLength(JSON.stringify(spreadsheet)),
          });
        }),
      },
};
