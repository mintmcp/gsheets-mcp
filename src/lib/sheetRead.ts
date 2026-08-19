/**
 * One bounded read of a native Google Sheets tab.
 *
 * Kept out of the tool definition so the paging arithmetic — window, probe
 * row, budget truncation, nextRange — can be read and tested on its own.
 * The .xlsx fallback stays in the tool, since it is a different data source
 * rather than a different way of windowing this one.
 */

import { quoteSheetName, a1Range } from './a1.js';
import { decodeGrid, type DecodeResult, type RawRow } from './cells.js';
import { MAX_CELLS, MAX_OUTPUT_CHARS, type Cell } from './sheetBudget.js';
import { makeSheetsRequest } from './google.js';
import { windowFor, type SheetWindow } from './window.js';

const GRID_FIELDS =
  'sheets.data.rowData.values(userEnteredValue,formattedValue,hyperlink,textFormatRuns)';

export interface NativeSheetPayload {
  id: string;
  sheetName: string;
  data: Cell[][];
  rowCount: number;
  columnCount: number;
  kind: 'native';
  returnedRange?: string;
  truncated?: true;
  nextRange?: string;
  message?: string;
}

interface TabProperties {
  title: string;
  gridProperties?: { rowCount?: number; columnCount?: number };
}

/**
 * Resolve the tab to read and its allocated dimensions. Always called: the
 * window needs the grid, and it turns a bad sheet_name into a clean error
 * instead of a Google 400. On an .xlsx upload this is the call that fails
 * and routes the caller to the office fallback.
 */
async function resolveTab(
  spreadsheetId: string,
  sheetName: string | undefined,
  accessToken: string,
): Promise<{ title: string; grid: { rowCount: number; columnCount: number } }> {
  const metadata = await makeSheetsRequest(
    `/${encodeURIComponent(spreadsheetId)}?fields=sheets.properties(title,gridProperties)`,
    accessToken,
    { method: 'GET' },
  ) as { sheets?: Array<{ properties: TabProperties }> };

  const sheets = metadata.sheets || [];
  if (sheets.length === 0) throw new Error('Spreadsheet has no sheets');

  const target = sheetName
    ? sheets.find((s) => s.properties.title === sheetName)
    : sheets[0];
  if (!target) throw new Error(`Sheet tab "${sheetName}" not found`);

  return {
    title: target.properties.title,
    grid: {
      rowCount: target.properties.gridProperties?.rowCount ?? 1,
      columnCount: target.properties.gridProperties?.columnCount ?? 1,
    },
  };
}

async function fetchRows(
  spreadsheetId: string,
  tabTitle: string,
  a1: string,
  accessToken: string,
): Promise<RawRow[]> {
  const range = encodeURIComponent(`${quoteSheetName(tabTitle)}!${a1}`);
  const result = await makeSheetsRequest(
    `/${encodeURIComponent(spreadsheetId)}?ranges=${range}&fields=${encodeURIComponent(GRID_FIELDS)}`,
    accessToken,
    { method: 'GET' },
  ) as { sheets?: Array<{ data?: Array<{ rowData?: RawRow[] }> }> };

  return result.sheets?.[0]?.data?.[0]?.rowData || [];
}

/**
 * Everything the response says about itself, derived from the window and what
 * the decoder actually produced. Pure, so the paging contract can be checked
 * without standing up an HTTP stub.
 */
export function describeRead(
  window: SheetWindow,
  decoded: DecodeResult,
  moreRowsExist: boolean,
): Pick<NativeSheetPayload, 'returnedRange' | 'truncated' | 'nextRange' | 'message'> {
  const { startColumn, columns, columnsOmitted, startRow, scopeEndRow } = window;
  const lastColumn = startColumn + columns - 1;
  // Paging resumes from the first row NOT returned, which is where the
  // decoder stopped when a budget tripped mid-window.
  const lastRowReturned = startRow + decoded.rowCount - 1;
  const morePages = moreRowsExist || decoded.truncated;

  const notes: string[] = [];
  if (decoded.truncated) {
    notes.push(`Output capped at ${MAX_CELLS} cells / ${MAX_OUTPUT_CHARS} characters.`);
  }
  if (decoded.partialRow) {
    notes.push('The final row is incomplete: it exceeds the character budget on its own.');
  }
  if (morePages) {
    notes.push(`Rows ${startRow}-${lastRowReturned} returned; pass nextRange as \`range\` to continue.`);
  }
  if (columnsOmitted > 0) {
    notes.push(`${columnsOmitted} column(s) beyond the ${columns}-column limit were not returned.`);
  }

  return {
    // Reports what is actually in `data`, so it is absent when nothing came
    // back. A zero column count covers the empty case too, since a grid with
    // no rows has no columns either.
    ...(decoded.columnCount > 0 && {
      returnedRange: a1Range(
        startColumn, startRow, startColumn + decoded.columnCount - 1, lastRowReturned,
      ),
    }),
    // Every note describes something the caller did not get, so the flag and
    // the explanation cannot drift apart.
    ...(notes.length > 0 && { truncated: true as const, message: notes.join(' ') }),
    // nextRange is the REMAINING SCOPE, not the next window: it gets clamped
    // again on receipt. Handing back a single window instead would shrink the
    // scope on every page and strand the tail.
    ...(morePages && {
      nextRange: a1Range(startColumn, lastRowReturned + 1, lastColumn, scopeEndRow),
    }),
  };
}

export async function readNativeWindow(
  spreadsheetId: string,
  sheetName: string | undefined,
  scope: string | undefined,
  accessToken: string,
): Promise<NativeSheetPayload> {
  const { title, grid } = await resolveTab(spreadsheetId, sheetName, accessToken);

  // An explicit range narrows the scope but never widens what we will fetch:
  // it is clamped to the same caps as a default read.
  const window = windowFor(grid, scope);
  const windowRows = window.endRow - window.startRow + 1;

  // Ask for one row past the window. If it comes back, more data really
  // exists; if it does not, the used range ended inside the window and
  // nothing was clipped. That turns the allocated-grid guess into an exact
  // answer.
  const hasProbe = window.endRow < window.scopeEndRow;
  const lastColumn = window.startColumn + window.columns - 1;
  const a1 = hasProbe
    ? a1Range(window.startColumn, window.startRow, lastColumn, window.endRow + 1)
    : window.a1;

  const rowData = await fetchRows(spreadsheetId, title, a1, accessToken);
  const decoded = decodeGrid(rowData.slice(0, windowRows));

  return {
    id: spreadsheetId,
    sheetName: title,
    data: decoded.data,
    rowCount: decoded.rowCount,
    columnCount: decoded.columnCount,
    kind: 'native',
    ...describeRead(window, decoded, hasProbe && rowData.length > windowRows),
  };
}
