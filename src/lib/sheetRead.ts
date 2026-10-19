/**
 * One bounded read of a native Google Sheets tab: the A1 scope to page
 * through, and the read that walks it.
 *
 * The scope's columns come from the tab's ALLOCATED grid, but the page's
 * height is sized by the width the data turns out to have. Google trims
 * trailing blank cells from every row and trailing blank rows from every
 * response, so each chunk fetched reveals the used width, and the next chunk
 * can be taller: a 3-column tab under a 26-column grid pages every 1,666
 * rows, not every 192.
 *
 * The .xlsx fallback stays in the tool, since it is a different data source
 * rather than a different way of windowing this one.
 */

import { quoteSheetName, a1Range, parseA1Range } from './a1.js';
import { decodeGrid, type RawRow } from './cells.js';
import { MAX_CELLS, MAX_OUTPUT_CHARS, truncationFields, type Cell } from './sheetBudget.js';
import { makeSheetsRequest } from './google.js';

export const MAX_RESPONSE_COLUMNS = 256;

/**
 * Ceiling on the cells one request may ask for, counted at the allocated
 * width. A chunk is sized by the width seen so far, so this is what bounds
 * the request when the rows it reaches turn out to be full width after all.
 */
export const MAX_FETCH_CELLS = 50_000;

/**
 * Ceiling on the cells one look-ahead may cover. A look-ahead only asks
 * which rows are populated, at about three bytes per populated cell, so it
 * can span far more than a fetch of content. The ceiling keeps the worst
 * case (a dense block right after a long gap) around 12MB, under the 25MB
 * response guard, while a 1,000,000-row grid that simply ended is settled
 * in about ten small reads.
 */
export const MAX_LOOKAHEAD_CELLS = 4_000_000;

/**
 * Ceiling on the rows one look-ahead may cover, since a populated row costs
 * bytes of its own: a 4,000,000-row single-column look-ahead would be mostly
 * row envelopes and still pass the cell ceiling.
 */
export const MAX_LOOKAHEAD_ROWS = 400_000;

export interface GridDimensions {
  rowCount: number;
  columnCount: number;
}

export interface SheetWindow {
  /** 1-based inclusive. */
  startRow: number;
  /**
   * Last row of the overall scope being paged through — the tab for a
   * default read, or the caller's rectangle for an explicit range. Paging
   * stops once the data or this row is reached.
   */
  scopeEndRow: number;
  /** 0-based index of the leftmost column in the window. */
  startColumn: number;
  columns: number;
  columnsOmitted: number;
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
  return {
    startRow,
    // Never a reversed scope: a startRow past the end (a paging call made
    // after the data ran out) collapses to a single row
    scopeEndRow: Math.max(startRow, scopeEndRow),
    startColumn,
    columns,
    columnsOmitted: Math.max(0, availableColumns - columns),
    empty,
  };
}

/**
 * The scope to page through for one read: the whole tab when `scope` is
 * absent, or the caller's bounded A1 rectangle intersected with the tab.
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

/** Rows `startRow..endRow` (1-based, inclusive) of the window's columns. */
export type FetchRows = (startRow: number, endRow: number) => Promise<RawRow[]>;

export interface RowSource {
  /** The rows with their content. */
  fetchRows: FetchRows;
  /**
   * The same rows with presence only: a populated cell is `{}`. Google
   * still trims trailing blanks, so this answers "where does data resume"
   * at a fraction of the cost of the content.
   */
  fetchPresence: FetchRows;
}

export interface Page {
  rows: RawRow[];
  /**
   * The row the next page starts at, present only when a populated row is
   * known to exist there or beyond: the row after the last one taken when
   * the probe row came back, or where data resumes after a run of blank
   * rows longer than a chunk. Absent means the data ended.
   */
  nextRow?: number;
}

function widest(rows: RawRow[]): number {
  let width = 0;
  for (const row of rows) {
    const cells = row.values?.length ?? 0;
    if (cells > width) width = cells;
  }
  return width;
}

/** A floor on what the rows will serialize to: the budget charges more. */
function textLength(rows: RawRow[]): number {
  let length = 0;
  for (const row of rows) {
    for (const cell of row.values ?? []) length += cell.formattedValue?.length ?? 0;
  }
  return length;
}

/**
 * Google trims trailing blank rows, so a chunk that comes back short cannot
 * say whether the data ended or a run of blanks longer than the chunk
 * follows. This settles it with presence-only reads of growing spans, so a
 * tab that really ended costs a few small empty answers rather than a read
 * of everything after it. Returns the first populated row at or after
 * `from`, or undefined when there is none in the scope.
 */
export async function findNextPopulatedRow(
  window: SheetWindow,
  fetchPresence: FetchRows,
  from: number,
): Promise<number | undefined> {
  const first = Math.max(1, Math.floor(MAX_FETCH_CELLS / window.columns));
  const cap = Math.max(first, Math.min(MAX_LOOKAHEAD_ROWS, Math.floor(MAX_LOOKAHEAD_CELLS / window.columns)));
  let span = first;
  for (let start = from; start <= window.scopeEndRow; start += span, span = Math.min(span * 4, cap)) {
    const found = await fetchPresence(start, Math.min(window.scopeEndRow, start + span - 1));
    const populated = found.findIndex((row) => (row.values?.length ?? 0) > 0);
    if (populated >= 0) return start + populated;
  }
  return undefined;
}

/**
 * Fetches the window in chunks until the page holds MAX_CELLS at the width
 * the data has, the character budget is spent, or the scope ends. The first
 * chunk is sized at the allocated width; each later one at the width the
 * rows so far revealed. Both measures are fetch heuristics: `decodeGrid`
 * remains the one place the budgets are enforced exactly.
 *
 * Every chunk asks for one row past its end. That row coming back is proof
 * the data continues, which is what lets a page end without a guess.
 */
export async function collectPage(window: SheetWindow, source: RowSource): Promise<Page> {
  const rows: RawRow[] = [];
  if (window.empty) return { rows };

  const chunkCap = Math.max(1, Math.floor(MAX_FETCH_CELLS / window.columns));
  let measured = 0;
  let chars = 0;
  let next = window.startRow;

  // Rows the page may still take: MAX_CELLS at the width seen so far, with
  // the allocated width standing in until the first chunk is back
  const room = () => Math.floor(MAX_CELLS / (measured || window.columns)) - rows.length;

  while (next <= window.scopeEndRow) {
    if (room() <= 0 || chars >= MAX_OUTPUT_CHARS) return { rows, nextRow: next };

    const count = Math.min(room(), chunkCap, window.scopeEndRow - next + 1);
    const end = next + count - 1;
    const probe = end < window.scopeEndRow ? 1 : 0;
    const fetched = await source.fetchRows(next, end + probe);
    const taken = fetched.slice(0, count);
    rows.push(...taken);
    measured = Math.max(measured, widest(taken));
    chars += textLength(taken);

    if (fetched.length === count + probe) {
      next = end + 1;
      continue;
    }
    if (probe === 0) break;
    const nextRow = await findNextPopulatedRow(window, source.fetchPresence, next + fetched.length);
    return nextRow === undefined ? { rows } : { rows, nextRow };
  }
  return { rows };
}

const GRID_FIELDS =
  'sheets.data.rowData.values(userEnteredValue.formulaValue,effectiveValue,formattedValue,hyperlink,textFormatRuns)';

/**
 * Asks for a field no cell carries, so every populated cell comes back as
 * `{}` and the response says only which rows and columns hold something.
 */
const PRESENCE_FIELDS = 'sheets.data.rowData.values.effectiveValue.errorValue.type';

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
  fields: string,
  accessToken: string,
): Promise<RawRow[]> {
  const range = encodeURIComponent(`${quoteSheetName(tabTitle)}!${a1}`);
  const result = await makeSheetsRequest(
    `/${encodeURIComponent(spreadsheetId)}?ranges=${range}&fields=${encodeURIComponent(fields)}`,
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
  const lastColumn = window.startColumn + window.columns - 1;
  const reader = (fields: string): FetchRows => (from, to) => fetchRows(
    spreadsheetId, title, a1Range(window.startColumn, from, lastColumn, to), fields, accessToken,
  );
  const page = await collectPage(window, {
    fetchRows: reader(GRID_FIELDS),
    fetchPresence: reader(PRESENCE_FIELDS),
  });
  const decoded = decodeGrid(page.rows);

  const lastRowReturned = window.startRow + decoded.rowCount - 1;
  // A page the decoder cut short resumes at the row after the cut, which is
  // known to hold data since it was fetched; otherwise where the fetch found
  // the data to continue
  const nextRow = decoded.truncated ? lastRowReturned + 1 : page.nextRow;
  const morePages = nextRow !== undefined;

  const notes: string[] = [];
  if (decoded.truncated) {
    // Usually characters: the page holds floor(MAX_CELLS / width) rows. The
    // cell cap can bind only when later chunks turned out wider than the
    // rows that sized them
    notes.push(decoded.cappedBy === 'cells'
      ? `Output capped at ${MAX_CELLS} cells.`
      : `Output capped at ${MAX_OUTPUT_CHARS} characters.`);
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
    ...(nextRow !== undefined && nextRow <= window.scopeEndRow && {
      nextRange: a1Range(window.startColumn, nextRow, lastColumn, window.scopeEndRow),
    }),
  };
}
