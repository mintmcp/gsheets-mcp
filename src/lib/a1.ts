/**
 * A1-notation parsing and validation helpers.
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
 * Convert a 0-based column index to A1 letters ("A", "Z", "AA").
 */
export function columnIndexToLetter(index: number): string {
  if (!Number.isInteger(index) || index < 0) {
    throw new Error(`column index must be a non-negative integer, got ${index}`);
  }
  let n = index + 1;
  let out = '';
  while (n > 0) {
    const remainder = (n - 1) % 26;
    out = String.fromCharCode(65 + remainder) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

/**
 * Every range this server emits goes through here, so the off-by-one lives in
 * one place.
 */
export function a1Range(
  startColumn: number,
  startRow: number,
  endColumn: number,
  endRow: number,
): string {
  return `${columnIndexToLetter(startColumn)}${startRow}`
    + `:${columnIndexToLetter(endColumn)}${endRow}`;
}

export interface GridRangeBounds {
  startRowIndex: number;
  endRowIndex: number;
  startColumnIndex: number;
  endColumnIndex: number;
}

/**
 * Render a GridRange back to bare A1. The inverse of `parseA1Range`, used to
 * report chart source ranges in the notation callers pass them in.
 *
 * GridRange end indices are exclusive, so an unbounded side (Google omits the
 * key for "to the end of the sheet") has no A1 spelling and is rejected rather
 * than guessed at.
 */
export function gridRangeToA1(range: Partial<GridRangeBounds>): string {
  const { startRowIndex, endRowIndex, startColumnIndex, endColumnIndex } = range;
  if (
    startRowIndex === undefined || endRowIndex === undefined
    || startColumnIndex === undefined || endColumnIndex === undefined
  ) {
    throw new Error('gridRangeToA1 requires a fully bounded range');
  }
  return a1Range(startColumnIndex, startRowIndex + 1, endColumnIndex - 1, endRowIndex);
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

export interface NormalizeRangeOptions {
  paramName?: string;
  /**
   * The tab the caller named separately. A prefix on the range must agree
   * with it, and either one becomes the result's sheetName.
   */
  sheetName?: string;
  /** Accept whole-column ("A:C"), whole-row ("1:3") and half-open ("B2:D") forms. */
  allowOpen?: boolean;
}

export interface NormalizedRange {
  /** The tab the range targets: its "Sheet!" prefix, else options.sheetName. */
  sheetName?: string;
  /** The range in bare A1, with any R1C1 translated. */
  a1: string;
}

/**
 * Split a range the way the official connector accepts it, "Sheet1!A1:B2"
 * or R1C1, into the tab it names and bare A1. Everything downstream works
 * in bare A1 on a known tab, so the translation happens once, here.
 *
 * Only absolute R1C1 (R1C1, R2C5:R9C7) is translated: a relative reference
 * like R[1]C[1] is relative to a cell this server does not have.
 */
export function normalizeRange(input: unknown, options: NormalizeRangeOptions = {}): NormalizedRange {
  const { paramName = 'range', sheetName, allowOpen = false } = options;
  if (typeof input !== 'string') {
    throw new Error(`${paramName} must be a string in A1 notation (e.g. "A1:C3" or "Sheet1!A1:C3")`);
  }
  let text = input.trim();
  if (text.length === 0) {
    throw new Error(`${paramName} must be a non-empty A1 string (e.g. "A1:C3")`);
  }

  let prefixed: string | undefined;
  const bang = text.lastIndexOf('!');
  if (bang !== -1) {
    const rawPrefix = text.slice(0, bang);
    text = text.slice(bang + 1);
    const quoted = rawPrefix.match(/^'(.*)'$/);
    prefixed = quoted ? quoted[1].replace(/''/g, "'") : rawPrefix;
    if (prefixed.length === 0) {
      throw new Error(`${paramName} "${input}" has an empty sheet name before "!"`);
    }
    if (text.length === 0) {
      throw new Error(`${paramName} "${input}" names a tab but no cells; give an A1 range after "!" or pass the tab via sheet_name alone`);
    }
    if (sheetName !== undefined && prefixed !== sheetName) {
      throw new Error(`${paramName} "${input}" names tab "${prefixed}" but sheet_name is "${sheetName}"; use one or make them agree`);
    }
  }

  if (/\[/.test(text) && /^R/i.test(text)) {
    throw new Error(`${paramName} "${input}" is a relative R1C1 reference; use absolute R1C1 ("R2C3") or A1 ("C2")`);
  }
  const a1 = r1c1ToA1(text) ?? text;
  const bounded = assertA1Form(a1, input, paramName, allowOpen);
  const tab = prefixed ?? sheetName;
  return { ...(tab !== undefined && { sheetName: tab }), a1: bounded };
}

export function requireTab(sheetName: string | undefined, paramName = 'range'): string {
  if (sheetName === undefined) {
    throw new Error(`Pass sheet_name, or a ${paramName} qualified with the tab like "Sheet1!A1"`);
  }
  return sheetName;
}

export function rangeOnTab(
  sheetName: string | undefined,
  range: unknown,
  options: Omit<NormalizeRangeOptions, 'sheetName'> = {},
): { sheetName: string; a1: string } {
  const parsed = normalizeRange(range, { ...options, sheetName });
  return { sheetName: requireTab(parsed.sheetName, options.paramName), a1: parsed.a1 };
}

/**
 * "R1C1:R4C4" -> "A1:D4". Only cell pairs are translated: "R1:R3" and
 * "C1:C3" are also the A1 rectangles in columns R and C, and A1 wins, so
 * whole rows and columns are spelled "1:3" and "A:C".
 */
function r1c1ToA1(text: string): string | undefined {
  const cellPart = (part: string): string | undefined => {
    const m = part.match(/^R(\d+)C(\d+)$/i);
    if (!m) return undefined;
    return `${columnIndexToLetter(parseInt(m[2], 10) - 1)}${m[1]}`;
  };
  const [first, second, extra] = text.split(':');
  if (extra !== undefined) return undefined;
  const a = cellPart(first);
  if (!a) return undefined;
  if (second === undefined) return a;
  const b = cellPart(second);
  return b ? `${a}:${b}` : undefined;
}

/**
 * The A1 grammar: a cell, a bounded rectangle, and, when `allowOpen`, whole
 * columns, whole rows and a rectangle open on its far side.
 */
function assertA1Form(a1: string, original: unknown, paramName: string, allowOpen: boolean): string {
  const bounded = a1.match(/^([A-Za-z]+)(\d+)(?::([A-Za-z]+)(\d+))?$/);
  if (bounded) {
    checkCorners(bounded[1], bounded[2], bounded[3], bounded[4], original, paramName);
    return a1;
  }
  const columns = a1.match(/^([A-Za-z]+):([A-Za-z]+)$/);
  const rows = a1.match(/^(\d+):(\d+)$/);
  const openRows = a1.match(/^([A-Za-z]+)(\d+):([A-Za-z]+)$/);
  const openColumns = a1.match(/^([A-Za-z]+)(\d+):(\d+)$/);
  const halfOpen = openRows ?? openColumns;
  if (!columns && !rows && !halfOpen) {
    throw new Error(
      `${paramName} "${original}" is not an A1 range. Use "A1" for a single cell or "A1:C3" for a range${allowOpen ? ', "A:C" for whole columns or "1:3" for whole rows' : ''}.`,
    );
  }
  if (!allowOpen) {
    const form = columns ? 'whole-column ("A:C")' : rows ? 'whole-row ("1:3")' : 'open-ended ("A1:C")';
    throw new Error(
      `${paramName} "${original}" is a ${form} range, which this tool does not support; give both corners, e.g. "A1:C3".`,
    );
  }
  if (columns && columnLetterToIndex(columns[2]) < columnLetterToIndex(columns[1])) {
    throw new Error(`${paramName} "${original}" has reversed endpoints.`);
  }
  if (rows && (parseInt(rows[1], 10) < 1 || parseInt(rows[2], 10) < parseInt(rows[1], 10))) {
    throw new Error(`${paramName} "${original}" has reversed endpoints or row 0; A1 rows are 1-indexed.`);
  }
  if (halfOpen && parseInt(halfOpen[2], 10) < 1) {
    throw new Error(`${paramName} "${original}" has invalid row 0; A1 rows are 1-indexed.`);
  }
  if (
    (openRows && columnLetterToIndex(openRows[3]) < columnLetterToIndex(openRows[1]))
    || (openColumns && parseInt(openColumns[3], 10) < parseInt(openColumns[2], 10))
  ) {
    throw new Error(`${paramName} "${original}" has reversed endpoints.`);
  }
  return a1;
}

function checkCorners(c1: string, r1: string, c2: string | undefined, r2: string | undefined, original: unknown, paramName: string): void {
  const startRow = parseInt(r1, 10);
  if (startRow < 1) {
    throw new Error(`${paramName} "${original}" has invalid row 0; A1 rows are 1-indexed.`);
  }
  if (c2 && r2) {
    const endRow = parseInt(r2, 10);
    if (endRow < 1) {
      throw new Error(`${paramName} "${original}" has invalid row 0; A1 rows are 1-indexed.`);
    }
    if (endRow < startRow || columnLetterToIndex(c2) < columnLetterToIndex(c1)) {
      throw new Error(
        `${paramName} "${original}" has reversed endpoints; the second cell must be at or after the first (e.g. "A1:C3", not "C3:A1").`,
      );
    }
  }
}
