/**
 * Reads .xlsx into the shape the native Sheets path returns. SheetJS parses;
 * this module owns the SheetJS-specific iteration. The cell shape, limits and
 * budget accounting are shared with the native decoder via sheetBudget.ts —
 * the two must stay identical, since both feed one tool's output schema.
 */

import { read, utils, type CellObject, type Range, type WorkSheet } from 'xlsx';
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

/**
 * `rowsBeyond` says the parse stopped short of rows the tab is known to hold,
 * so the tab is reported truncated even if the budget was not spent.
 */
function readSheet(ws: WorkSheet | undefined, budget: Budget, rowsBeyond = false): ParsedSheet {
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
  if (rowsBeyond) sheet.truncated = true;
  return finish(sheet);
}

function finish(sheet: ParsedSheet): ParsedSheet {
  Object.assign(sheet, gridDimensions(sheet.data));
  return sheet;
}

interface BoundedSheet {
  ws: WorkSheet | undefined;
  /** The tab holds rows past those parsed; see readSheet. */
  rowsBeyond: boolean;
}

/**
 * Parse only as many rows of one tab as the cell budget can use.
 *
 * Without `sheetRows`, SheetJS builds an object for every cell of the tab
 * before readSheet stops at the budget: returning 50,000 cells of a
 * 3.26M-cell tab took ~755MB of heap and crashed the 1GB machine. The rows
 * the budget needs depend on the tab's width, so a short first pass measures
 * it and a second reads that many rows.
 *
 * The result must read exactly as a whole-tab parse would, which takes two
 * rules:
 *
 * - With `sheetRows`, SheetJS narrows `!ref` to the cells it saw and moves
 *   the file's declared <dimension>, if any, to `!fullref`. A whole-tab parse
 *   reads with the declared extent, so the returned sheet's `!ref` is reset
 *   to it (cut at the rows parsed). Otherwise a wide cell further down sizes
 *   the pass without being charged by readSheet, and a truncated tab reads as
 *   complete.
 * - Without a declared dimension, a pass shows only what it found: one that
 *   ends before its last row may have stopped at a gap with data below. Only
 *   a pass filled to its last row is sized from; any other is read whole, as
 *   every tab was before. SheetJS reports the dimension only when it runs
 *   past the pass, so a tab short enough for the first pass is read whole
 *   too, which costs little because it is short.
 */
function readBoundedSheet(
  bytes: Uint8Array,
  index: number,
  rawName: string,
  maxCells: number,
): BoundedSheet {
  // undefined means no row limit: the whole-tab parse this replaced
  let rows: number | undefined = FIRST_PASS_ROWS;
  let declared: Range | undefined;
  for (;;) {
    const wb = read(bytes, {
      type: 'array',
      cellDates: true,
      sheets: index,
      ...(rows !== undefined && { sheetRows: rows }),
    });
    const ws = wb.Sheets?.[rawName];
    if (!ws || rows === undefined) return { ws, rowsBeyond: false };

    // With no cell parsed, SheetJS leaves a placeholder "A1" in `!ref`
    const parsed = ws['!ref'] && hasCells(ws) ? utils.decode_range(ws['!ref']) : undefined;
    // `!fullref` appears only when the declared extent runs past the pass. A
    // later pass that lacks it has the whole declared extent in hand.
    if (ws['!fullref']) declared = utils.decode_range(ws['!fullref']);

    if (!declared && (!parsed || parsed.e.r < rows - 1)) {
      rows = undefined;
      continue;
    }
    const extent = declared ? union(parsed, declared) : parsed!;

    // readSheet charges one cell per leading blank row, then every column
    // from A to the last on each row after it. One spare row makes the
    // budget, not the parse, end the read.
    const width = extent.e.c + 1;
    const needed = extent.s.r + Math.ceil(Math.max(0, maxCells - extent.s.r) / width) + 1;
    // A declared extent that ends inside the pass means the whole tab is in
    // hand; a parsed one can only say so once the budget would run out.
    if ((declared && extent.e.r < rows) || needed <= rows) {
      ws['!ref'] = utils.encode_range({
        s: extent.s,
        e: { r: Math.min(extent.e.r, rows - 1), c: extent.e.c },
      });
      return { ws, rowsBeyond: declared !== undefined && extent.e.r > rows - 1 };
    }
    rows = needed < XLSX_MAX_ROWS ? needed : undefined;
  }
}

/** The smallest range covering both. */
function union(a: Range | undefined, b: Range): Range {
  if (!a) return b;
  return {
    s: { r: Math.min(a.s.r, b.s.r), c: Math.min(a.s.c, b.s.c) },
    e: { r: Math.max(a.e.r, b.e.r), c: Math.max(a.e.c, b.e.c) },
  };
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
    const { ws, rowsBeyond } = readBoundedSheet(bytes, i, raw, budget.maxCells);
    return { ...base, ...readSheet(ws, budget, rowsBeyond) };
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
