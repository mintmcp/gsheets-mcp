/**
 * The .xlsx reader must parse only the rows its cell budget can return. A
 * whole-tab parse of a 3.26M-cell upload took ~755MB of heap to return 50,000
 * cells, which crashed the 1GB machine. SheetJS is wrapped so the tests can
 * see how many rows each parse was allowed to read, and so the parity tests
 * can force the whole-tab parse this replaced.
 */

import { describe, it, expect, vi, beforeEach } from 'vitest';
import { utils, write } from 'xlsx';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { parseXlsx, XLSX_MAX_CELLS, type ParseOptions } from '../lib/xlsx.js';
import { xlsxSheetOutput } from '../lib/office.js';

const parses = vi.hoisted(() => [] as Array<number | undefined>);
const sheetJs = vi.hoisted(() => ({ wholeTab: false }));

vi.mock('xlsx', async (importOriginal) => {
  const actual = await importOriginal<typeof import('xlsx')>();
  return {
    ...actual,
    read: (data: any, opts: any) => {
      if (opts?.bookSheets) return actual.read(data, opts);
      if (sheetJs.wholeTab) {
        const { sheetRows: _dropped, ...whole } = opts;
        return actual.read(data, whole);
      }
      parses.push(opts?.sheetRows);
      return actual.read(data, opts);
    },
  };
});

/** An .xlsx with one tab of `rows` x `cols` cells, the first at `origin`. */
function book(rows: number, cols: number, origin = 'A1'): Uint8Array {
  const aoa = Array.from({ length: rows }, (_, r) =>
    Array.from({ length: cols }, (_, c) => `r${r}c${c}`));
  const ws = utils.aoa_to_sheet([]);
  utils.sheet_add_aoa(ws, aoa, { origin });
  const wb = utils.book_new();
  utils.book_append_sheet(wb, ws, 'Data');
  return new Uint8Array(write(wb, { type: 'array', bookType: 'xlsx' }));
}

/** An .xlsx with one tab holding exactly the given cells, e.g. { A1: 'x' }. */
function cells(values: Record<string, string>): Uint8Array {
  const ws = utils.aoa_to_sheet([]);
  for (const [at, value] of Object.entries(values)) {
    utils.sheet_add_aoa(ws, [[value]], { origin: at });
  }
  const wb = utils.book_new();
  utils.book_append_sheet(wb, ws, 'Data');
  return new Uint8Array(write(wb, { type: 'array', bookType: 'xlsx' }));
}

/** The same workbook without the <dimension> tag, as some writers emit it. */
function withoutDimension(bytes: Uint8Array): Uint8Array {
  const parts = unzipSync(bytes);
  const sheet = 'xl/worksheets/sheet1.xml';
  parts[sheet] = strToU8(strFromU8(parts[sheet]).replace(/<dimension[^>]*\/>/, ''));
  return zipSync(parts);
}

beforeEach(() => {
  parses.length = 0;
  sheetJs.wholeTab = false;
});

/** parseXlsx as it was before the bounded parse: SheetJS reads every row. */
function parseWhole(bytes: Uint8Array, opts?: ParseOptions) {
  sheetJs.wholeTab = true;
  try {
    return parseXlsx(bytes, opts);
  } finally {
    sheetJs.wholeTab = false;
  }
}

/** What a caller can observe of a parse, for comparing two of them. */
const observed = (wb: ReturnType<typeof parseXlsx>) => ({
  truncated: wb.truncated,
  cells: wb.cells,
  chars: wb.chars,
  sheets: wb.sheets.map(({ data, rowCount, columnCount, truncated }) =>
    ({ data, rowCount, columnCount, truncated })),
});

describe('bounded .xlsx parsing', () => {
  it('parses only the rows the cell budget needs from a long tab', () => {
    const wb = parseXlsx(book(30_000, 5));
    const [sheet] = wb.sheets;

    expect(wb.truncated).toBe(true);
    expect(sheet.rowCount).toBe(XLSX_MAX_CELLS / 5);
    expect(sheet.data[9_999][4]).toEqual({ value: 'r9999c4' });
    // A 64-row pass to measure the width, then one sized to the budget
    expect(parses).toEqual([64, XLSX_MAX_CELLS / 5 + 1]);
  });

  it('sizes the second pass by width, so a wide tab reads few rows', () => {
    const wb = parseXlsx(book(400, 300));
    expect(wb.truncated).toBe(true);
    expect(wb.sheets[0].rowCount).toBe(Math.floor(XLSX_MAX_CELLS / 300));
    expect(parses[1]).toBeLessThan(200);
  });

  it('reads a tab that ends inside the first pass whole', () => {
    // SheetJS reports no dimension for a tab that fits the pass, so a short
    // tab looks like a sparse one and is read whole; it is short, so cheap.
    const wb = parseXlsx(book(10, 3));
    expect(wb.truncated).toBe(false);
    expect(wb.sheets[0].rowCount).toBe(10);
    expect(parses).toEqual([64, undefined]);
  });

  it('returns a tab that fits the budget whole, not cut at a pass boundary', () => {
    const wb = parseXlsx(book(5_000, 4));
    expect(wb.truncated).toBe(false);
    expect(wb.sheets[0].rowCount).toBe(5_000);
    expect(wb.sheets[0].data[4_999][3]).toEqual({ value: 'r4999c3' });
  });

  it('honors a caller-supplied cell cap when sizing the pass', () => {
    const wb = parseXlsx(book(2_000, 10), { maxCells: 1_000 });
    expect(wb.sheets[0].rowCount).toBe(100);
    expect(parses[1]).toBe(101);
  });

  it('finds data below the first pass through the declared dimension', () => {
    const wb = parseXlsx(book(3, 2, 'A200'));
    const [sheet] = wb.sheets;
    expect(wb.truncated).toBe(false);
    expect(sheet.rowCount).toBe(202);
    expect(sheet.data[199][0]).toEqual({ value: 'r0c0' });
    // The second pass holds the declared extent whole, so no third is needed
    expect(parses).toHaveLength(2);
    expect(parses[1]).toBeDefined();
  });

  it('reads the tab whole when nothing tells an empty pass from data further down', () => {
    const wb = parseXlsx(withoutDimension(book(3, 2, 'A200')));
    const [sheet] = wb.sheets;
    expect(sheet.rowCount).toBe(202);
    expect(sheet.data[201][1]).toEqual({ value: 'r2c1' });
    expect(parses).toEqual([64, undefined]);
  });

  it('bounds a long tab that declares no dimension by what it parsed', () => {
    const wb = parseXlsx(withoutDimension(book(30_000, 5)));
    expect(wb.truncated).toBe(true);
    expect(wb.sheets[0].rowCount).toBe(XLSX_MAX_CELLS / 5);
    expect(parses).toEqual([64, XLSX_MAX_CELLS / 5 + 1]);
  });

  it('keeps a value below a blank gap when the file declares no dimension', () => {
    // The first pass ends at A1, well before its last row, but the tab does
    // not: nothing short of reading on can tell a gap from the end.
    const wb = parseXlsx(withoutDimension(cells({ A1: 'header', A100: 'must survive' })));
    const [sheet] = wb.sheets;
    expect(sheet.truncated).toBe(false);
    expect(sheet.rowCount).toBe(100);
    expect(sheet.data[99][0]).toEqual({ value: 'must survive' });
  });

  it('reports a tab truncated when a declared wide cell lies below the pass', () => {
    // Column ALL is the 1,000th: the declared width sizes the pass at 51
    // rows, so the first 64-row pass is kept, and must be read at that width.
    const wb = parseXlsx(cells({ A1: 'header', ALL100: 'outside initial parse' }));
    const [sheet] = wb.sheets;
    expect(wb.truncated).toBe(true);
    expect(wb.cells).toBe(XLSX_MAX_CELLS);
    expect(sheet.rowCount).toBe(XLSX_MAX_CELLS / 1_000);

    const out: any = xlsxSheetOutput('id1', wb, undefined);
    expect(out.truncated).toBe(true);
    expect(out.message).toContain(`${XLSX_MAX_CELLS} cells`);
    expect(out.message).toContain('convert_to_google_sheet');
  });
});

describe('bounded parse matches the whole-tab parse', () => {
  const shapes: Record<string, () => Uint8Array> = {
    'long dense tab': () => book(12_000, 5),
    'long dense tab, no dimension': () => withoutDimension(book(12_000, 5)),
    'short tab': () => book(10, 3),
    'short tab, no dimension': () => withoutDimension(book(10, 3)),
    'wide tab': () => book(400, 300),
    'sparse tab': () => cells({ A1: 'header', A100: 'late' }),
    'sparse tab, no dimension': () => withoutDimension(cells({ A1: 'header', A100: 'late' })),
    'wide cell below the pass': () => cells({ A1: 'header', ALL100: 'late' }),
    'wide cell below the pass, no dimension': () =>
      withoutDimension(cells({ A1: 'header', ALL100: 'late' })),
    'data starting below the pass': () => book(3, 2, 'A200'),
    'data starting below the pass, no dimension': () => withoutDimension(book(3, 2, 'A200')),
    'full first pass, then a wide cell': () => withoutDimension(
      cells({ ...Object.fromEntries(Array.from({ length: 70 }, (_, r) => [`A${r + 1}`, `v${r}`])), ALL100: 'late' }),
    ),
    'empty tab': () => book(0, 0),
  };

  for (const [shape, make] of Object.entries(shapes)) {
    it(`reads the ${shape} exactly as a whole-tab parse does`, () => {
      const bytes = make();
      expect(observed(parseXlsx(bytes))).toEqual(observed(parseWhole(bytes)));
    });
  }

  it('matches under a caller-supplied cell cap', () => {
    const bytes = book(2_000, 10);
    expect(observed(parseXlsx(bytes, { maxCells: 1_000 })))
      .toEqual(observed(parseWhole(bytes, { maxCells: 1_000 })));
  });

  it('still reports truncation when a dimensionless tab widens past the sized pass', () => {
    // The one shape where the outputs differ. Without a declared dimension a
    // column first used past the sized pass is invisible to it, so the whole
    // parse pads its rows to that width and the bounded one does not. Both
    // stop at the budget and both say so.
    const ws = utils.aoa_to_sheet(Array.from({ length: 60_000 }, () => ['x']));
    utils.sheet_add_aoa(ws, [['late']], { origin: 'Z60001' });
    const wb = utils.book_new();
    utils.book_append_sheet(wb, ws, 'Data');
    const bytes = withoutDimension(new Uint8Array(write(wb, { type: 'array', bookType: 'xlsx' })));
    const bounded = parseXlsx(bytes);
    const whole = parseWhole(bytes);
    expect(bounded.truncated).toBe(true);
    expect(whole.truncated).toBe(true);
    expect(bounded.cells).toBe(XLSX_MAX_CELLS);
  });
});
