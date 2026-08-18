/**
 * Reads .xlsx into the shape the native Sheets path returns. SheetJS parses;
 * this module owns the SheetJS-specific iteration. The cell shape, limits and
 * budget accounting are shared with the native decoder via sheetBudget.ts —
 * the two must stay identical, since both feed one tool's output schema.
 */

import { read, utils, type CellObject, type WorkSheet } from 'xlsx';
import {
  chargeCell,
  chargeEmptyRow,
  clipValue,
  boundLinks,
  safeLinkUrl,
  createBudget,
  exhausted,
  gridDimensions,
  MAX_CELL_CHARS,
  type Budget,
  type Cell,
  type CellType,
} from './sheetBudget.js';

/** Hard structural limits of the .xlsx format, distinct from response caps. */
const XLSX_MAX_ROWS = 1_048_576;
const XLSX_MAX_COLUMNS = 16_384;
export const MAX_SHEETS = 1_000;
export const MAX_SHEET_NAME_CHARS = 255;

export class XlsxInvalidError extends Error {
  constructor(message: string) { super(message); this.name = 'XlsxInvalidError'; }
}
export class XlsxEncryptedError extends Error {
  constructor(message: string) { super(message); this.name = 'XlsxEncryptedError'; }
}

/** The .xlsx decoder emits the same cell shape as the native decoder. */
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


function cellText(cell: CellObject): string {
  if (cell.w !== undefined) return cell.w;
  if (cell.v instanceof Date) return cell.v.toISOString().slice(0, 10);
  return cell.v === undefined || cell.v === null ? '' : String(cell.v);
}

function cellType(cell: CellObject): CellType {
  if (cell.f !== undefined) return 'formula';
  switch (cell.t) {
    case 'n': return 'number';
    case 'd': return 'number';
    case 'b': return 'boolean';
    case 'z': return 'empty';
    default: return 'string';
  }
}

export function toCell(cell: CellObject | undefined, budget: Budget = createBudget()): XlsxCell {
  if (!cell) return { value: '', type: 'empty' };

  const value = cell.f !== undefined ? `=${cell.f}` : cellText(cell);
  const out: XlsxCell = {
    value: clipValue(value, budget),
    type: cellType(cell),
  };

  const url = safeLinkUrl(cell.l?.Target, budget);
  if (url) {
    const display = cell.f !== undefined ? cellText(cell) : out.value;
    out.hyperlinks = boundLinks(
      [{ url, start: 0, end: display.length }],
      Math.min(display.length, budget.maxCellChars),
    );
  }
  return out;
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
    if (exhausted(budget)) {
      sheet.truncated = true;
      return finish(sheet);
    }
    sheet.data.push([]);
    chargeEmptyRow(budget);
  }

  for (let r = range.s.r; r <= lastRow; r++) {
    const row: XlsxCell[] = [];
    for (let c = 0; c <= lastCol; c++) {
      if (exhausted(budget)) {
        sheet.truncated = true;
        if (row.length) sheet.data.push(row);
        return finish(sheet);
      }
      const cell = toCell(ws[utils.encode_cell({ c, r })] as CellObject | undefined, budget);
      row.push(cell);
      chargeCell(budget, cell);
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
  const wb = opts.namesOnly ? index : read(bytes, {
    type: 'array',
    cellDates: true,
    sheets: requested,
  });

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

  const budget = createBudget({ maxCells: opts.maxCells, maxChars: opts.maxChars });
  const sheets = listed.map((raw, i) => {
    const base = stub(raw);
    if (i !== requested) return { ...base, notRequested: true };
    if (listed.indexOf(raw) !== i) return { ...base, unreadable: true };
    return { ...base, ...readSheet(wb.Sheets?.[raw], budget) };
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
