/**
 * The .xlsx reader must parse only the rows its cell budget can return. A
 * whole-tab parse of a 3.26M-cell upload took ~755MB of heap to return 50,000
 * cells, which crashed the 1GB machine. SheetJS is wrapped so the tests can
 * see how many rows each parse was allowed to read.
 */

import { describe, it, expect, vi, beforeEach } from 'vitest';
import { utils, write } from 'xlsx';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { parseXlsx, XLSX_MAX_CELLS } from '../lib/xlsx.js';

const parses = vi.hoisted(() => [] as Array<number | undefined>);

vi.mock('xlsx', async (importOriginal) => {
  const actual = await importOriginal<typeof import('xlsx')>();
  return {
    ...actual,
    read: (data: any, opts: any) => {
      if (!opts?.bookSheets) parses.push(opts?.sheetRows);
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

/** The same workbook without the <dimension> tag, as some writers emit it. */
function withoutDimension(bytes: Uint8Array): Uint8Array {
  const parts = unzipSync(bytes);
  const sheet = 'xl/worksheets/sheet1.xml';
  parts[sheet] = strToU8(strFromU8(parts[sheet]).replace(/<dimension[^>]*\/>/, ''));
  return zipSync(parts);
}

beforeEach(() => {
  parses.length = 0;
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

  it('reads a tab that ends inside the first pass in one parse', () => {
    const wb = parseXlsx(book(10, 3));
    expect(wb.truncated).toBe(false);
    expect(wb.sheets[0].rowCount).toBe(10);
    expect(parses).toEqual([64]);
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
});
