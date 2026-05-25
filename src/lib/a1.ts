/**
 * A1-notation parsing and validation helpers.
 */

/**
 * Quote a sheet name for use in A1 notation.
 * Wraps in single quotes and escapes any existing single quotes.
 */
export function quoteSheetName(name: string): string {
  return `'${name.replace(/'/g, "''")}'`;
}

/**
 * Convert a column letter (e.g. "A", "B", "AA", "AZ") to a 0-based index.
 */
export function columnLetterToIndex(letter: string): number {
  let index = 0;
  const upper = letter.toUpperCase();
  for (let i = 0; i < upper.length; i++) {
    index = index * 26 + (upper.charCodeAt(i) - 64);
  }
  return index - 1;
}

/**
 * Validate that a user-supplied range string is a bounded bare A1 range
 * (no sheet prefix), with both endpoints fully specified (column letters
 * AND row digits) and not reversed. Returns the trimmed value.
 *
 * Accepts: "A1" or "A1:C3". Rejects sheet-qualified strings, open-ended
 * whole-column / whole-row ranges, row 0, and reversed endpoints.
 */
export function assertBareA1Range(range: unknown, paramName = 'range'): string {
  if (typeof range !== 'string') {
    throw new Error(`${paramName} must be a string in A1 notation (e.g. "A1:C3")`);
  }
  const trimmed = range.trim();
  if (trimmed.length === 0) {
    throw new Error(`${paramName} must be a non-empty A1 string (e.g. "A1:C3")`);
  }
  if (trimmed.includes('!')) {
    throw new Error(
      `${paramName} must be a bare A1 range like "A1:C3" — do not include a sheet prefix. Pass the sheet name via the sheet_name argument instead.`,
    );
  }
  const match = trimmed.match(/^([A-Za-z]+)(\d+)(?::([A-Za-z]+)(\d+))?$/);
  if (!match) {
    throw new Error(
      `${paramName} "${range}" is not a bounded A1 range. Use "A1" for a single cell or "A1:C3" for a range; whole-column ("A:C") and whole-row ("1:3") forms are not supported.`,
    );
  }
  const startRow = parseInt(match[2], 10);
  if (startRow < 1) {
    throw new Error(`${paramName} "${range}" has invalid row 0; A1 rows are 1-indexed.`);
  }
  if (match[3] && match[4]) {
    const endRow = parseInt(match[4], 10);
    if (endRow < 1) {
      throw new Error(`${paramName} "${range}" has invalid row 0; A1 rows are 1-indexed.`);
    }
    const startCol = columnLetterToIndex(match[1]);
    const endCol = columnLetterToIndex(match[3]);
    if (endRow < startRow || endCol < startCol) {
      throw new Error(
        `${paramName} "${range}" has reversed endpoints; the second cell must be at or after the first (e.g. "A1:C3", not "C3:A1").`,
      );
    }
  }
  return trimmed;
}

/**
 * Validate that the input is a single A1 cell (no range, no sheet prefix).
 * Returns the input unchanged on success.
 */
export function assertSingleCell(cell: unknown, paramName = 'cell'): string {
  if (typeof cell !== 'string') {
    throw new Error(`${paramName} must be a single cell in A1 notation (e.g. "B3")`);
  }
  if (cell.includes(':')) {
    throw new Error(`${paramName} must be a single cell in A1 notation (e.g. "B3"), not a range. Use update_range for ranges.`);
  }
  if (!/^[A-Za-z]+\d+$/.test(cell)) {
    throw new Error(`Invalid A1 cell: ${cell}`);
  }
  return cell;
}

/**
 * Parse an A1-style range (e.g. "A1:C3", "B2", "A1") into grid indices.
 * Returns 0-based indices suitable for GridRange.
 */
export function parseA1Range(range: string): {
  startRowIndex: number;
  endRowIndex: number;
  startColumnIndex: number;
  endColumnIndex: number;
} {
  const match = range.match(/^([A-Za-z]+)(\d+)(?::([A-Za-z]+)(\d+))?$/);
  if (!match) {
    throw new Error(`Invalid A1 range: ${range}`);
  }

  const startCol = columnLetterToIndex(match[1]);
  const startRow = parseInt(match[2], 10) - 1;
  const endCol = match[3] ? columnLetterToIndex(match[3]) : startCol;
  const endRow = match[4] ? parseInt(match[4], 10) - 1 : startRow;

  return {
    startRowIndex: startRow,
    endRowIndex: endRow + 1,
    startColumnIndex: startCol,
    endColumnIndex: endCol + 1,
  };
}
