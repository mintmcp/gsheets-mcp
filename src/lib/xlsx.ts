/**
 * Reads .xlsx into the shape the native Sheets path returns. SheetJS parses;
 * this module owns the SheetJS-specific iteration. The cell shape, limits and
 * budget accounting are shared with the native decoder via sheetBudget.ts —
 * the two must stay identical, since both feed one tool's output schema.
 */

import { read, utils, type CellObject, type WorkSheet } from 'xlsx';
import {
  admitCell,
  admitEmptyRow,
  makeCell,
  createBudget,
  gridDimensions,
  type Budget,
  type Cell,
  type CellType,
} from './sheetBudget.js';

/**
 * Response caps for this path only. The native reader is windowed and pages
 * through `nextRange`, so anything its budget refuses arrives on the next
 * call and the caps can be sized for one ingestible page. Nothing here can
 * page, so a refused cell is data the caller has no way to ask for again.
 * These are the values .xlsx shipped with; revisit them alongside paging.
 */
export const XLSX_MAX_CELLS = 50_000;
export const XLSX_MAX_OUTPUT_CHARS = 4_000_000;

/** Hard structural limits of the .xlsx format, distinct from response caps. */
const XLSX_MAX_ROWS = 1_048_576;
const XLSX_MAX_COLUMNS = 16_384;

/** Rows in the first, width-measuring parse of a tab. */
const FIRST_PASS_ROWS = 64;
export const MAX_SHEETS = 1_000;
export const MAX_SHEET_NAME_CHARS = 255;

export class XlsxInvalidError extends Error {
  constructor(message: string) { super(message); this.name = 'XlsxInvalidError'; }
}
export class XlsxEncryptedError extends Error {
  constructor(message: string) { super(message); this.name = 'XlsxEncryptedError'; }
}

export type XlsxCell = Cell;

export interface XlsxSheet {
  name: string;
  rawName: string;
  data: XlsxCell[][];
  rowCount: number;
  columnCount: number;
  truncated: boolean;
  unreadable?: boolean;
  notRequested?: boolean;
  nameShortened?: boolean;
}

export interface XlsxWorkbook {
  sheets: XlsxSheet[];
  truncated: boolean;
  sheetsOmitted: number;
  cells: number;
  chars: number;
}

export interface ParseOptions {
  maxCells?: number;
  maxChars?: number;
  namesOnly?: boolean;
  sheet?: string | number;
}


function resultType(t: CellObject['t']): CellType {
  switch (t) {
    case 'n': return 'number';
    case 'd': return 'number';
    case 'b': return 'boolean';
    case 'e': return 'error';
    default: return 'string';
  }
}

/** Rendered text first, so a date is never a raw serial or a local-time Date. */
function resultOf(cell: CellObject | undefined): { display: string; type: CellType } {
  // A formula whose cached result was never written has a type tag but no value
  if (!cell || cell.v === undefined || cell.v === null) return { display: '', type: 'empty' };
  if (cell.w !== undefined) return { display: cell.w, type: resultType(cell.t) };
  if (cell.v instanceof Date) return { display: cell.v.toISOString().slice(0, 10), type: 'number' };
  return { display: String(cell.v), type: resultType(cell.t) };
}

export function toCell(cell: CellObject | undefined, budget: Budget = createBudget()): XlsxCell {
  const { display, type } = resultOf(cell);
  const url = cell?.l?.Target;
  return makeCell({
    display,
    type,
    formula: cell?.f !== undefined ? `=${cell.f}` : undefined,
    links: url ? [{ url, start: 0, end: display.length }] : undefined,
  }, budget);
}

type ParsedSheet = Omit<XlsxSheet, 'name' | 'rawName'>;

function readSheet(ws: WorkSheet | undefined, budget: Budget): ParsedSheet {
  const sheet: ParsedSheet = { data: [], rowCount: 0, columnCount: 0, truncated: false };
  if (!ws) return { ...sheet, unreadable: true };

  if (!ws['!ref']) return sheet;

  const range = utils.decode_range(ws['!ref']);
  const lastRow = Math.min(range.e.r, XLSX_MAX_ROWS - 1);
  const lastCol = Math.min(range.e.c, XLSX_MAX_COLUMNS - 1);

  for (let r = 0; r < Math.min(range.s.r, XLSX_MAX_ROWS - 1); r++) {
    if (!admitEmptyRow(budget)) {
      sheet.truncated = true;
      return finish(sheet);
    }
    sheet.data.push([]);
  }

  for (let r = range.s.r; r <= lastRow; r++) {
    const row: XlsxCell[] = [];
    for (let c = 0; c <= lastCol; c++) {
      const cell = toCell(ws[utils.encode_cell({ c, r })] as CellObject | undefined, budget);
      if (!admitCell(budget, cell)) {
        sheet.truncated = true;
        // Drop the partial row, matching the native decoder: a row missing
        // its tail in the middle of `data` cannot be told apart from a short
        // row, and .xlsx has no paging to recover the remainder.
        if (sheet.data.length === 0 && row.length) sheet.data.push(row);
        return finish(sheet);
      }
      row.push(cell);
    }
    while (row.length && row[row.length - 1].type === 'empty') row.pop();
    sheet.data.push(row);
  }
  return finish(sheet);
}

function finish(sheet: ParsedSheet): ParsedSheet {
  Object.assign(sheet, gridDimensions(sheet.data));
  return sheet;
}

/**
 * Parse only as many rows of one tab as the cell budget can use.
 *
 * Without `sheetRows`, SheetJS builds an object for every cell of the tab
 * before readSheet stops at the budget: returning 50,000 cells of a
 * 3.26M-cell tab took ~755MB of heap and crashed the 1GB machine. The rows
 * the budget needs depend on the tab's width, which only a parse reveals, so
 * a short first pass measures it and a second reads that many rows. Width
 * only grows as rows are added, so the second pass never falls short.
 *
 * When the file declares no <dimension> and the first pass holds no cell, an
 * empty tab and one whose data starts further down look the same, so that
 * case (and a budget that would need every row) is read whole, as every tab
 * was before.
 */
function readBoundedSheet(
  bytes: Uint8Array,
  index: number,
  rawName: string,
  maxCells: number,
): WorkSheet | undefined {
  // undefined means no row limit: the whole-tab parse this replaced
  let rows: number | undefined = FIRST_PASS_ROWS;
  for (;;) {
    const wb = read(bytes, {
      type: 'array',
      cellDates: true,
      sheets: index,
      ...(rows !== undefined && { sheetRows: rows }),
    });
    const ws = wb.Sheets?.[rawName];
    if (!ws || rows === undefined) return ws;

    // With sheetRows, SheetJS clamps `!ref` to the rows it parsed (or leaves
    // a placeholder "A1" when none held a cell) and moves the <dimension>
    // the file declares, when there is one, to `!fullref`.
    const parsed = ws['!ref'] ? utils.decode_range(ws['!ref']) : undefined;
    const declared = ws['!fullref'] ? utils.decode_range(ws['!fullref']) : undefined;
    const extent = declared ?? (parsed && hasCells(ws) ? parsed : undefined);
    if (!extent) {
      rows = undefined;
      continue;
    }
    const cut = (parsed !== undefined && parsed.e.r >= rows - 1)
      || (declared !== undefined && declared.e.r > rows - 1);
    if (!cut) return ws;

    // readSheet charges one cell per leading blank row, then every column
    // from A to the last on each row after it. One spare row makes the
    // budget, not the parse, end the read.
    const width = extent.e.c + 1;
    const needed = extent.s.r + Math.ceil(Math.max(0, maxCells - extent.s.r) / width) + 1;
    if (needed <= rows) return ws;
    rows = needed < XLSX_MAX_ROWS ? needed : undefined;
  }
}

/** True when a parsed sheet holds at least one cell, not only `!` metadata. */
function hasCells(ws: WorkSheet): boolean {
  for (const key in ws) {
    if (key[0] !== '!') return true;
  }
  return false;
}

function startsWith(bytes: Uint8Array, sig: number[]): boolean {
  return sig.every((b, i) => bytes[i] === b);
}

function resolveSheet(rawNames: string[], sheet: ParseOptions['sheet']): number {
  if (sheet === undefined) return 0;
  if (typeof sheet === 'number') return sheet;
  return rawNames.findIndex(
    (raw) => raw === sheet || raw.slice(0, MAX_SHEET_NAME_CHARS) === sheet
  );
}

export function parseXlsx(bytes: Uint8Array, opts: ParseOptions = {}): XlsxWorkbook {
  if (startsWith(bytes, [0xd0, 0xcf, 0x11, 0xe0])) {
    throw new XlsxEncryptedError(
      'the file is an OLE compound document — either password-protected or a legacy .xls'
    );
  }
  if (!startsWith(bytes, [0x50, 0x4b, 0x03, 0x04])) {
    throw new XlsxInvalidError('the file is not a readable .xlsx (no zip header)');
  }

  try {
    return readWorkbook(bytes, opts);
  } catch (err) {
    if (err instanceof XlsxInvalidError || err instanceof XlsxEncryptedError) throw err;
    const message = err instanceof Error ? err.message : String(err);
    if (/password|encrypt/i.test(message)) {
      throw new XlsxEncryptedError('the workbook is password-protected');
    }
    throw new XlsxInvalidError(`the file could not be parsed (${message.slice(0, 120)})`);
  }
}

function readWorkbook(bytes: Uint8Array, opts: ParseOptions): XlsxWorkbook {
  const index = read(bytes, { type: 'array', bookSheets: true });
  const rawNames = index.SheetNames ?? [];
  if (rawNames.length === 0) throw new XlsxInvalidError('the workbook contains no sheets');

  const requested = resolveSheet(rawNames, opts.sheet);

  const sheetsOmitted = Math.max(rawNames.length - MAX_SHEETS, 0);
  const listed = rawNames.slice(0, MAX_SHEETS);

  const stub = (raw: string): XlsxSheet => ({
    name: raw.slice(0, MAX_SHEET_NAME_CHARS),
    rawName: raw,
    data: [], rowCount: 0, columnCount: 0, truncated: false,
    ...(raw.length > MAX_SHEET_NAME_CHARS ? { nameShortened: true } : {}),
  });

  if (opts.namesOnly) {
    return { sheets: listed.map(stub), truncated: false, sheetsOmitted, cells: 0, chars: 0 };
  }

  const budget = createBudget({
    maxCells: opts.maxCells ?? XLSX_MAX_CELLS,
    maxChars: opts.maxChars ?? XLSX_MAX_OUTPUT_CHARS,
  });
  const sheets = listed.map((raw, i) => {
    const base = stub(raw);
    if (i !== requested) return { ...base, notRequested: true };
    if (listed.indexOf(raw) !== i) return { ...base, unreadable: true };
    return { ...base, ...readSheet(readBoundedSheet(bytes, i, raw, budget.maxCells), budget) };
  });

  const parsed = sheets.filter((s) => !s.notRequested);
  if (parsed.length > 0 && parsed.every((s) => s.unreadable)) {
    throw new XlsxInvalidError('no worksheet in the workbook could be read');
  }

  return {
    sheets,
    truncated: sheets.some((s) => s.truncated),
    sheetsOmitted,
    cells: budget.cells,
    chars: budget.chars,
  };
}
