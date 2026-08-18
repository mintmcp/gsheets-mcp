import { describe, it, expect } from 'vitest';
import { columnIndexToLetter, windowFor } from '../lib/window.js';
import { columnLetterToIndex } from '../lib/a1.js';

const CAPS = { maxCells: 50_000, maxColumns: 256 };

describe('columnIndexToLetter', () => {
  it('maps 0-based indices to A1 column letters', () => {
    expect(columnIndexToLetter(0)).toBe('A');
    expect(columnIndexToLetter(25)).toBe('Z');
    expect(columnIndexToLetter(26)).toBe('AA');
    expect(columnIndexToLetter(51)).toBe('AZ');
    expect(columnIndexToLetter(52)).toBe('BA');
    expect(columnIndexToLetter(701)).toBe('ZZ');
  });

  it('round-trips with columnLetterToIndex', () => {
    for (const i of [0, 1, 25, 26, 27, 51, 52, 200, 701, 702, 16383]) {
      expect(columnLetterToIndex(columnIndexToLetter(i))).toBe(i);
    }
  });

  it('rejects negative or non-integer indices', () => {
    expect(() => columnIndexToLetter(-1)).toThrow(/non-negative/);
    expect(() => columnIndexToLetter(1.5)).toThrow(/non-negative/);
  });
});

describe('windowFor over a whole tab', () => {
  it('covers a small tab entirely', () => {
    const w = windowFor({ rowCount: 100, columnCount: 10 }, undefined, CAPS);
    expect(w).toEqual({
      startRow: 1, endRow: 100, startColumn: 0, columns: 10,
      columnsOmitted: 0, scopeEndRow: 100, a1: 'A1:J100',
    });
  });

  it('clamps rows so the window fits the cell cap', () => {
    const w = windowFor({ rowCount: 500_000, columnCount: 26 }, undefined, CAPS);
    expect(w.startRow).toBe(1);
    expect(w.endRow).toBe(1923);
    expect(w.a1).toBe('A1:Z1923');
    expect(w.scopeEndRow).toBe(500_000);
  });

  it('clamps columns and reports how many were dropped', () => {
    const w = windowFor({ rowCount: 10, columnCount: 400 }, undefined, CAPS);
    expect(w.columns).toBe(256);
    expect(w.columnsOmitted).toBe(144);
    expect(w.a1).toBe('A1:IV10');
  });

  it('always yields at least one row and one column', () => {
    const w = windowFor({ rowCount: 0, columnCount: 0 }, undefined, CAPS);
    expect(w).toEqual({
      startRow: 1, endRow: 1, startColumn: 0, columns: 1,
      columnsOmitted: 0, scopeEndRow: 1, a1: 'A1:A1',
    });
  });

  it('never exceeds the cell cap even at the column limit', () => {
    const w = windowFor({ rowCount: 500_000, columnCount: 256 }, undefined, CAPS);
    expect(w.columns * (w.endRow - w.startRow + 1)).toBeLessThanOrEqual(50_000);
  });
});

describe('windowFor over an explicit range', () => {
  const wide = { rowCount: 100_000, columnCount: 26 };

  it('passes through a range already within the caps', () => {
    const w = windowFor(wide, 'A1:C500', CAPS);
    expect(w.a1).toBe('A1:C500');
    expect(w.columnsOmitted).toBe(0);
    expect(w.scopeEndRow).toBe(500);
  });

  it('clamps a range that exceeds the cell cap', () => {
    const w = windowFor(wide, 'A1:Z100000', CAPS);
    expect(w.endRow).toBe(1923);
    expect(w.a1).toBe('A1:Z1923');
    expect(w.scopeEndRow).toBe(100_000);
  });

  it('preserves a non-A start column', () => {
    const w = windowFor(wide, 'C5:E900', CAPS);
    expect(w.startColumn).toBe(2);
    expect(w.a1).toBe('C5:E900');
  });

  it('pages from the start row of the range it is given', () => {
    // How paging actually works: nextRange comes back as the remaining
    // scope and is handed straight back in as the next request.
    const w = windowFor(wide, 'A1924:Z100000', CAPS);
    expect(w.a1).toBe('A1924:Z3846');
    expect(w.scopeEndRow).toBe(100_000);
  });

  it('does not count columns the tab never had as omitted', () => {
    // Regression: sizing off the rectangle alone reported 446 omitted
    // columns here and flagged a fully satisfied request as truncated.
    const w = windowFor({ rowCount: 100, columnCount: 26 }, 'A1:ZZ10', CAPS);
    expect(w.columns).toBe(26);
    expect(w.columnsOmitted).toBe(0);
    expect(w.a1).toBe('A1:Z10');
  });

  it('still reports omitted columns when the tab really is that wide', () => {
    const w = windowFor({ rowCount: 100, columnCount: 400 }, 'A1:OJ10', CAPS);
    expect(w.columns).toBe(256);
    expect(w.columnsOmitted).toBe(144);
    expect(w.a1).toBe('A1:IV10');
  });

  it('stops the scope at the end of the tab', () => {
    const w = windowFor({ rowCount: 3000, columnCount: 26 }, 'A1:Z500000', CAPS);
    expect(w.scopeEndRow).toBe(3000);
  });

  it('never emits a reversed range for a scope past the end of the tab', () => {
    const w = windowFor({ rowCount: 2000, columnCount: 26 }, 'A2001:Z9000', CAPS);
    expect(w.endRow).toBeGreaterThanOrEqual(w.startRow);
    expect(w.a1).toBe('A2001:Z2001');
  });

  it('yields a valid range for a scope starting past the last column', () => {
    const w = windowFor({ rowCount: 100, columnCount: 26 }, 'AA1:AB5', CAPS);
    expect(w.startColumn).toBe(26);
    expect(w.columns).toBe(1);
    expect(w.a1).toBe('AA1:AA5');
  });
});
