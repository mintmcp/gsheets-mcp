import { describe, it, expect } from 'vitest';
import { windowFor, MAX_RESPONSE_COLUMNS } from '../lib/sheetRead.js';
import { MAX_CELLS } from '../lib/sheetBudget.js';

/**
 * Deliberately runs against the real caps rather than injected ones, so a cap
 * change breaks this file.
 */
describe('windowFor over a whole tab', () => {
  it('covers a small tab entirely', () => {
    const w = windowFor({ rowCount: 100, columnCount: 10 }, undefined);
    expect(w).toEqual({
      startRow: 1, endRow: 100, startColumn: 0, columns: 10,
      columnsOmitted: 0, scopeEndRow: 100, a1: 'A1:J100', empty: false,
    });
  });

  it('clamps rows so the window fits the cell cap', () => {
    const w = windowFor({ rowCount: 500_000, columnCount: 26 }, undefined);
    expect(w.endRow).toBe(Math.floor(MAX_CELLS / 26));
    expect(w.a1).toBe('A1:Z192');
    expect(w.scopeEndRow).toBe(500_000);
  });

  it('clamps columns and reports how many were dropped', () => {
    const w = windowFor({ rowCount: 10, columnCount: 400 }, undefined);
    expect(w.columns).toBe(MAX_RESPONSE_COLUMNS);
    expect(w.columnsOmitted).toBe(400 - MAX_RESPONSE_COLUMNS);
    expect(w.a1).toBe('A1:IV10');
  });

  it('always yields at least one row and one column', () => {
    const w = windowFor({ rowCount: 0, columnCount: 0 }, undefined);
    expect(w).toEqual({
      startRow: 1, endRow: 1, startColumn: 0, columns: 1,
      columnsOmitted: 0, scopeEndRow: 1, a1: 'A1:A1', empty: false,
    });
  });

  it('never exceeds the cell cap, even at the column limit', () => {
    const w = windowFor({ rowCount: 500_000, columnCount: MAX_RESPONSE_COLUMNS }, undefined);
    expect(w.columns * (w.endRow - w.startRow + 1)).toBeLessThanOrEqual(MAX_CELLS);
  });
});

describe('windowFor over an explicit range', () => {
  const wide = { rowCount: 100_000, columnCount: 26 };

  it('passes through a range already within the caps', () => {
    const w = windowFor(wide, 'A1:C500');
    expect(w.a1).toBe('A1:C500');
    expect(w.columnsOmitted).toBe(0);
    expect(w.scopeEndRow).toBe(500);
  });

  it('clamps a range that exceeds the cell cap', () => {
    const w = windowFor(wide, 'A1:Z100000');
    expect(w.a1).toBe('A1:Z192');
    expect(w.scopeEndRow).toBe(100_000);
  });

  it('preserves a non-A start column', () => {
    const w = windowFor(wide, 'C5:E900');
    expect(w.startColumn).toBe(2);
    expect(w.a1).toBe('C5:E900');
  });

  it('pages from the start row of the range it is given', () => {
    // How paging actually works: nextRange comes back as the remaining scope
    // and is handed straight back in as the next request.
    const w = windowFor(wide, 'A193:Z100000');
    expect(w.a1).toBe('A193:Z384');
    expect(w.scopeEndRow).toBe(100_000);
  });

  it('does not count columns the tab never had as omitted', () => {
    const w = windowFor({ rowCount: 100, columnCount: 26 }, 'A1:ZZ10');
    expect(w.columns).toBe(26);
    expect(w.columnsOmitted).toBe(0);
    expect(w.a1).toBe('A1:Z10');
  });

  it('still reports omitted columns when the tab really is that wide', () => {
    const w = windowFor({ rowCount: 100, columnCount: 400 }, 'A1:OJ10');
    expect(w.columns).toBe(MAX_RESPONSE_COLUMNS);
    expect(w.columnsOmitted).toBe(400 - MAX_RESPONSE_COLUMNS);
    expect(w.a1).toBe('A1:IV10');
  });

  it('stops the scope at the end of the tab', () => {
    const w = windowFor({ rowCount: 3000, columnCount: 26 }, 'A1:Z500000');
    expect(w.scopeEndRow).toBe(3000);
  });

  it('never emits a reversed range for a scope past the end of the tab', () => {
    const w = windowFor({ rowCount: 2000, columnCount: 26 }, 'A2001:Z9000');
    expect(w.endRow).toBeGreaterThanOrEqual(w.startRow);
    expect(w.a1).toBe('A2001:Z2001');
  });

  it('marks a scope starting past the last column as empty', () => {
    // Sheets answers an out-of-grid range with a 400, so this must not be
    // fetched: "Range exceeds grid limits. Max rows: 3002, max columns: 28".
    const w = windowFor({ rowCount: 100, columnCount: 26 }, 'AA1:AB5');
    expect(w.empty).toBe(true);
  });

  it('marks a scope starting past the last row as empty', () => {
    const w = windowFor({ rowCount: 100, columnCount: 26 }, 'A500:Z600');
    expect(w.empty).toBe(true);
  });

  it('does not mark a scope that overlaps the grid as empty', () => {
    const w = windowFor({ rowCount: 100, columnCount: 26 }, 'A90:Z600');
    expect(w.empty).toBe(false);
    expect(w.scopeEndRow).toBe(100);
  });
});
