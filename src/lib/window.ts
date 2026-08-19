/**
 * Computes the A1 window to request from the Sheets API so the response is
 * bounded before it is fetched, rather than truncated after it arrives.
 *
 * `gridProperties` describes the ALLOCATED grid, not the used range, so the
 * window is an upper bound: the caller compares the rows actually returned
 * against the window to decide whether anything was really clipped.
 */

import { MAX_CELLS } from './sheetBudget.js';
import { parseA1Range, a1Range } from './a1.js';

export const MAX_RESPONSE_COLUMNS = 256;

export interface GridDimensions {
  rowCount: number;
  columnCount: number;
}

export interface WindowOptions {
  maxCells?: number;
  maxColumns?: number;
}

export interface SheetWindow {
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
}

function buildWindow(
  startRow: number,
  scopeEndRow: number,
  startColumn: number,
  availableColumns: number,
  maxCells: number,
  maxColumns: number,
): SheetWindow {
  const columns = Math.min(availableColumns, maxColumns);
  const columnsOmitted = Math.max(0, availableColumns - columns);
  const rowsPerWindow = Math.max(1, Math.floor(maxCells / columns));
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
  opts: WindowOptions = {},
): SheetWindow {
  const tabRows = Math.max(1, grid.rowCount);
  const tabColumns = Math.max(1, grid.columnCount);
  const maxCells = opts.maxCells ?? MAX_CELLS;
  const maxColumns = opts.maxColumns ?? MAX_RESPONSE_COLUMNS;

  if (scope === undefined) {
    return buildWindow(1, tabRows, 0, tabColumns, maxCells, maxColumns);
  }

  const rect = parseA1Range(scope);
  const startColumn = rect.startColumnIndex;
  // Floors at 1 so a rectangle starting past the tab's last column still
  // produces a valid A1 string; the fetch then simply comes back empty.
  const availableColumns = Math.max(
    1,
    Math.min(rect.endColumnIndex, tabColumns) - startColumn,
  );

  return buildWindow(
    Math.max(1, rect.startRowIndex + 1),
    Math.max(1, Math.min(rect.endRowIndex, tabRows)),
    startColumn,
    availableColumns,
    maxCells,
    maxColumns,
  );
}
