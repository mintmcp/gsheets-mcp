/**
 * 2D grid helpers shared by value-writing tools.
 */

/**
 * Pad ragged rows with empty strings so every row has the same length as
 * the widest row. Throws if `data` is empty or every row is empty.
 * Returns the padded matrix.
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
  const maxCols = Math.max(...rows.map((r) => r.length));
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
