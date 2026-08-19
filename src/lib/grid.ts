/**
 * 2D grid helpers shared by value-writing tools.
 */

/**
 * Inbound writes are a separate concern from read page size: the request body
 * is already bounded at 10MB, and shrinking batches to the read cap would
 * force callers into needless round trips.
 */
export const MAX_WRITE_CELLS = 50_000;

/**
 * Widest row length. Uses a loop rather than `Math.max(...rows)` because
 * spread passes one argument per element and overflows the call stack
 * somewhere past ~100k rows — reachable with a tall sheet.
 */
export function maxRowLength(rows: ReadonlyArray<ReadonlyArray<unknown>>): number {
  let max = 0;
  for (const row of rows) {
    if (row.length > max) max = row.length;
  }
  return max;
}

/**
 * Reject an oversized write before `padRaggedRows` copies the matrix and it
 * is serialized again for the Sheets API body.
 */
export function assertCellCount(
  rows: ReadonlyArray<ReadonlyArray<unknown>>,
  max: number,
): void {
  let total = 0;
  for (const row of rows) total += row.length;
  if (total > max) {
    throw new Error(
      `data contains ${total} cells, over the limit of ${max}. Split the write into smaller batches.`,
    );
  }
}

/**
 * Pad ragged rows with empty strings so every row has the same length as
 * the widest row. Throws if `data` is empty or every row is empty.
 */
export function padRaggedRows(data: unknown): string[][] {
  if (!Array.isArray(data) || data.length === 0) {
    throw new Error('data must contain at least one row');
  }
  for (const row of data) {
    if (!Array.isArray(row)) {
      throw new Error('data must be a 2D array of strings');
    }
  }
  const rows = data as string[][];
  const maxCols = maxRowLength(rows);
  if (maxCols === 0) {
    throw new Error('data rows must contain at least one cell');
  }
  return rows.map((row) => {
    const padded = [...row];
    while (padded.length < maxCols) {
      padded.push('');
    }
    return padded;
  });
}
