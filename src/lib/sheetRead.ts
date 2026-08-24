/**
 * One bounded read of a native Google Sheets tab: the A1 window to request,
 * and the read that uses it.
 *
 * The window is computed from the tab's ALLOCATED grid, which is usually
 * larger than the used range, so it is an upper bound — the read compares the
 * rows actually returned against it to decide whether anything was clipped.
 *
 * The .xlsx fallback stays in the tool, since it is a different data source
 * rather than a different way of windowing this one.
 */

import { quoteSheetName, a1Range, parseA1Range } from './a1.js';
import { decodeGrid, type RawRow } from './cells.js';
import { MAX_CELLS, MAX_OUTPUT_CHARS, truncationFields, type Cell } from './sheetBudget.js';
import { makeSheetsRequest } from './google.js';

export const MAX_RESPONSE_COLUMNS = 256;

export interface GridDimensions {
  rowCount: number;
  columnCount: number;
}

interface SheetWindow {
  /** 1-based inclusive. */
  startRow: number;
  /** 1-based inclusive. */
  endRow: number;
  /** 0-based index of the leftmost column in the window. */
  startColumn: number;
  columns: number;
  columnsOmitted: number;
  /**
   * Last row of the overall scope being paged through — the tab for a
   * default read, or the caller's rectangle for an explicit range. Paging
   * stops once `endRow` reaches it.
   */
  scopeEndRow: number;
  a1: string;
  /**
   * True when the requested rectangle lies outside the allocated grid, so
   * there is nothing to fetch. Sheets rejects an out-of-grid range with a
   * 400 rather than returning nothing, so the caller must not send it.
   */
  empty: boolean;
}

function buildWindow(
  startRow: number,
  scopeEndRow: number,
  startColumn: number,
  availableColumns: number,
  empty = false,
): SheetWindow {
  const columns = Math.min(availableColumns, MAX_RESPONSE_COLUMNS);
  const columnsOmitted = Math.max(0, availableColumns - columns);
  const rowsPerWindow = Math.max(1, Math.floor(MAX_CELLS / columns));
  // Never emit a reversed range: a startRow past the end of the scope (a
  // paging call made after the data ran out) collapses to a single row.
  const endRow = Math.max(startRow, Math.min(scopeEndRow, startRow + rowsPerWindow - 1));

  return {
    startRow,
    endRow,
    startColumn,
    columns,
    columnsOmitted,
    scopeEndRow,
    a1: a1Range(startColumn, startRow, startColumn + columns - 1, endRow),
    empty,
  };
}

/**
 * The window to fetch for one read: the whole tab when `scope` is absent, or
 * the caller's bounded A1 rectangle intersected with the tab.
 *
 * Intersecting matters for more than efficiency. Sizing an explicit range off
 * the rectangle alone made `columnsOmitted` count columns the tab never had,
 * so a fully satisfied request came back flagged as truncated.
 */
export function windowFor(
  grid: GridDimensions,
  scope: string | undefined,
): SheetWindow {
  const tabRows = Math.max(1, grid.rowCount);
  const tabColumns = Math.max(1, grid.columnCount);

  if (scope === undefined) {
    return buildWindow(1, tabRows, 0, tabColumns);
  }

  const rect = parseA1Range(scope);
  const startColumn = rect.startColumnIndex;
  const startRow = Math.max(1, rect.startRowIndex + 1);
  const columnsInGrid = Math.min(rect.endColumnIndex, tabColumns) - startColumn;
  // A rectangle beginning past the tab's last row or column has no
  // intersection with the grid. The A1 string still has to be well formed,
  // but nothing should be fetched for it.
  const empty = columnsInGrid <= 0 || startRow > tabRows;

  return buildWindow(
    startRow,
    Math.max(1, Math.min(rect.endRowIndex, tabRows)),
    startColumn,
    Math.max(1, columnsInGrid),
    empty,
  );
}


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

export async function readNativeWindow(
  spreadsheetId: string,
  sheetName: string | undefined,
  scope: string | undefined,
  accessToken: string,
): Promise<NativeSheetPayload> {
  const { title, grid } = await resolveTab(spreadsheetId, sheetName, accessToken);

  const window = windowFor(grid, scope);
  const windowRows = window.endRow - window.startRow + 1;

  // One row past the window: if it comes back the used range really extends
  // beyond, which turns the allocated-grid guess into an exact answer.
  const hasProbe = window.endRow < window.scopeEndRow;
  const lastColumn = window.startColumn + window.columns - 1;
  const a1 = hasProbe
    ? a1Range(window.startColumn, window.startRow, lastColumn, window.endRow + 1)
    : window.a1;

  // Sheets rejects an out-of-grid range with a 400, so an empty intersection
  // is answered without a request rather than by asking for nothing.
  let rowData = window.empty
    ? []
    : await fetchRows(spreadsheetId, title, a1, accessToken);

  // Sheets drops trailing blank rows, so a short answer cannot tell "the data
  // ended here" from "the next row is past a run of blanks longer than the probe"
  const probeInconclusive = hasProbe
    && rowData.length <= windowRows
    && window.endRow + 1 < window.scopeEndRow;
  if (!window.empty && probeInconclusive) {
    rowData = await fetchRows(
      spreadsheetId,
      title,
      a1Range(window.startColumn, window.startRow, lastColumn, window.scopeEndRow),
      accessToken,
    );
  }
  const decoded = decodeGrid(rowData.slice(0, windowRows));

  const lastRowReturned = window.startRow + decoded.rowCount - 1;
  const morePages = (hasProbe && rowData.length > windowRows) || decoded.truncated;

  const notes: string[] = [];
  if (decoded.truncated) {
    // Characters, always: the window holds at most floor(MAX_CELLS/columns)
    // rows, so the cell cap bounds the FETCH and can never end a page here.
    notes.push(`Output capped at ${MAX_OUTPUT_CHARS} characters.`);
  }
  if (decoded.partialRow) {
    notes.push('The final row is incomplete: it exceeds the character budget on its own.');
  }
  if (morePages) {
    notes.push(`Rows ${window.startRow}-${lastRowReturned} returned; pass nextRange as \`range\` to continue.`);
  }
  if (window.columnsOmitted > 0) {
    notes.push(`${window.columnsOmitted} column(s) past the ${window.columns}-column window were not returned; read them by passing a \`range\` that starts at a later column.`);
  }

  return {
    id: spreadsheetId,
    sheetName: title,
    data: decoded.data,
    rowCount: decoded.rowCount,
    columnCount: decoded.columnCount,
    kind: 'native',
    ...(decoded.columnCount > 0 && {
      returnedRange: a1Range(
        window.startColumn, window.startRow,
        window.startColumn + decoded.columnCount - 1, lastRowReturned,
      ),
    }),
    ...truncationFields(notes),
    // The REMAINING SCOPE, not the next window: it gets clamped again on
    // receipt. Guarded on rows actually remaining, or a partial single-row
    // page emits a reversed range the caller cannot pass back.
    ...(morePages && lastRowReturned < window.scopeEndRow && {
      nextRange: a1Range(
        window.startColumn, lastRowReturned + 1, lastColumn, window.scopeEndRow,
      ),
    }),
  };
}
