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
