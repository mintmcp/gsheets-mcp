/**
 * 2D grid helpers shared by value-writing tools.
 */

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
 * The memory cliff from rectangularizing a sparse matrix, which the 7MB body
 * limit does not see: 25,000 one-cell rows plus one row of 25,000 cells is
 * well under 1MB to send and pads to 625 million entries. At this ceiling the
 * padded array and its JSON come to roughly 25MB.
 */
export const MAX_PADDED_CELLS = 2_000_000;

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
  const paddedCells = rows.length * maxCols;
  if (paddedCells > MAX_PADDED_CELLS) {
    throw new Error(
      `data pads out to ${paddedCells} cells because its widest row has ${maxCols}. `
      + 'Send rows of even width, or split the write into smaller batches.',
    );
  }
  return rows.map((row) => {
    const padded = [...row];
    while (padded.length < maxCols) {
      padded.push('');
    }
    return padded;
  });
}
