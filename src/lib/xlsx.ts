/**
 * Reads .xlsx into the shape the native Sheets path returns. SheetJS parses;
 * this module owns the output shape and the budgets that keep one hostile
 * workbook from filling a 128MB isolate.
 */

import { read, utils, type CellObject, type WorkSheet } from 'xlsx';

export const MAX_CELLS = 50_000;
export const MAX_OUTPUT_CHARS = 4_000_000;
export const MAX_CELL_CHARS = 32_768;
const CELL_ENVELOPE_CHARS = 50;
const MAX_ROWS = 1_048_576;
const MAX_COLUMNS = 16_384;
export const MAX_SHEETS = 1_000;
export const MAX_SHEET_NAME_CHARS = 255;

export class XlsxInvalidError extends Error {
  constructor(message: string) { super(message); this.name = 'XlsxInvalidError'; }
}
export class XlsxEncryptedError extends Error {
  constructor(message: string) { super(message); this.name = 'XlsxEncryptedError'; }
}

export interface XlsxCell {
  value: string;
  type: 'string' | 'number' | 'boolean' | 'formula' | 'empty';
  hyperlinks?: Array<{ url: string; start: number; end: number }>;
}

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

const SAFE_LINK = /^(https?|mailto):/i;

function cellText(cell: CellObject): string {
  if (cell.w !== undefined) return cell.w;
  if (cell.v instanceof Date) return cell.v.toISOString().slice(0, 10);
  return cell.v === undefined || cell.v === null ? '' : String(cell.v);
}

function cellType(cell: CellObject): XlsxCell['type'] {
  if (cell.f !== undefined) return 'formula';
  switch (cell.t) {
    case 'n': return 'number';
    case 'd': return 'number';
    case 'b': return 'boolean';
    case 'z': return 'empty';
    default: return 'string';
  }
}

export function toCell(cell: CellObject | undefined): XlsxCell {
  if (!cell) return { value: '', type: 'empty' };

  const value = cell.f !== undefined ? `=${cell.f}` : cellText(cell);
  const out: XlsxCell = {
    value: value.slice(0, MAX_CELL_CHARS),
    type: cellType(cell),
  };

  const url = cell.l?.Target;
  if (url && SAFE_LINK.test(url)) {
    const display = cell.f !== undefined ? cellText(cell) : out.value;
    out.hyperlinks = [{
      url: url.slice(0, MAX_CELL_CHARS),
      start: 0,
      end: Math.min(display.length, MAX_CELL_CHARS),
    }];
  }
  return out;
}

interface Budget {
  cells: number;
  chars: number;
  maxCells: number;
  maxChars: number;
}

const exhausted = (b: Budget) => b.cells >= b.maxCells || b.chars >= b.maxChars;

type ParsedSheet = Omit<XlsxSheet, 'name' | 'rawName'>;

function readSheet(ws: WorkSheet | undefined, budget: Budget): ParsedSheet {
  const sheet: ParsedSheet = { data: [], rowCount: 0, columnCount: 0, truncated: false };
  if (!ws) return { ...sheet, unreadable: true };

  if (!ws['!ref']) return sheet;

  const range = utils.decode_range(ws['!ref']);
  const lastRow = Math.min(range.e.r, MAX_ROWS - 1);
  const lastCol = Math.min(range.e.c, MAX_COLUMNS - 1);

  for (let r = 0; r < Math.min(range.s.r, MAX_ROWS - 1); r++) {
    if (exhausted(budget)) {
      sheet.truncated = true;
      return finish(sheet);
    }
    sheet.data.push([]);
    budget.cells++;
    budget.chars += CELL_ENVELOPE_CHARS;
  }

  for (let r = range.s.r; r <= lastRow; r++) {
    const row: XlsxCell[] = [];
    for (let c = 0; c <= lastCol; c++) {
      if (exhausted(budget)) {
        sheet.truncated = true;
        if (row.length) sheet.data.push(row);
        return finish(sheet);
      }
      const cell = toCell(ws[utils.encode_cell({ c, r })] as CellObject | undefined);
      row.push(cell);
      budget.cells++;
      budget.chars += cell.value.length + (cell.hyperlinks?.[0]?.url.length ?? 0)
        + CELL_ENVELOPE_CHARS;
    }
    while (row.length && row[row.length - 1].type === 'empty') row.pop();
    sheet.data.push(row);
  }
  return finish(sheet);
}

function finish(sheet: ParsedSheet): ParsedSheet {
  sheet.rowCount = sheet.data.length;
  sheet.columnCount = sheet.data.reduce((max, row) => Math.max(max, row.length), 0);
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

  const budget: Budget = {
    cells: 0,
    chars: 0,
    maxCells: opts.maxCells ?? MAX_CELLS,
    maxChars: opts.maxChars ?? MAX_OUTPUT_CHARS,
  };
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
