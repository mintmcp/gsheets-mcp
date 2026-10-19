import { describe, it, expect } from 'vitest';
import { windowFor, collectPage, MAX_RESPONSE_COLUMNS, MAX_FETCH_CELLS, MAX_LOOKAHEAD_CELLS, MAX_LOOKAHEAD_ROWS, type SheetWindow } from '../lib/sheetRead.js';
import { MAX_CELLS } from '../lib/sheetBudget.js';
import type { RawRow } from '../lib/cells.js';

/**
 * Deliberately runs against the real caps rather than injected ones, so a cap
 * change breaks this file.
 */
describe('windowFor over a whole tab', () => {
  it('covers a small tab entirely', () => {
    const w = windowFor({ rowCount: 100, columnCount: 10 }, undefined);
    expect(w).toEqual({
      startRow: 1, scopeEndRow: 100, startColumn: 0, columns: 10, columnsOmitted: 0, empty: false,
    });
  });

  it('keeps the whole tab as the scope, however tall', () => {
    const w = windowFor({ rowCount: 500_000, columnCount: 26 }, undefined);
    expect(w.scopeEndRow).toBe(500_000);
  });

  it('clamps columns and reports how many were dropped', () => {
    const w = windowFor({ rowCount: 10, columnCount: 400 }, undefined);
    expect(w.columns).toBe(MAX_RESPONSE_COLUMNS);
    expect(w.columnsOmitted).toBe(400 - MAX_RESPONSE_COLUMNS);
  });

  it('always yields at least one row and one column', () => {
    const w = windowFor({ rowCount: 0, columnCount: 0 }, undefined);
    expect(w).toEqual({
      startRow: 1, scopeEndRow: 1, startColumn: 0, columns: 1, columnsOmitted: 0, empty: false,
    });
  });
});

describe('windowFor over an explicit range', () => {
  const wide = { rowCount: 100_000, columnCount: 26 };

  it('passes through a range already within the tab', () => {
    const w = windowFor(wide, 'A1:C500');
    expect(w).toMatchObject({ startRow: 1, scopeEndRow: 500, startColumn: 0, columns: 3, columnsOmitted: 0 });
  });

  it('preserves a non-A start column', () => {
    const w = windowFor(wide, 'C5:E900');
    expect(w.startColumn).toBe(2);
    expect(w.startRow).toBe(5);
    expect(w.columns).toBe(3);
  });

  it('does not count columns the tab never had as omitted', () => {
    const w = windowFor({ rowCount: 100, columnCount: 26 }, 'A1:ZZ10');
    expect(w.columns).toBe(26);
    expect(w.columnsOmitted).toBe(0);
  });

  it('still reports omitted columns when the tab really is that wide', () => {
    const w = windowFor({ rowCount: 100, columnCount: 400 }, 'A1:OJ10');
    expect(w.columns).toBe(MAX_RESPONSE_COLUMNS);
    expect(w.columnsOmitted).toBe(400 - MAX_RESPONSE_COLUMNS);
  });

  it('stops the scope at the end of the tab', () => {
    const w = windowFor({ rowCount: 3000, columnCount: 26 }, 'A1:Z500000');
    expect(w.scopeEndRow).toBe(3000);
  });

  it('never emits a reversed scope for a range past the end of the tab', () => {
    const w = windowFor({ rowCount: 2000, columnCount: 26 }, 'A2001:Z9000');
    expect(w.scopeEndRow).toBeGreaterThanOrEqual(w.startRow);
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

describe('collectPage', () => {
  /** A tab whose rows `populated` hold `width` cells each, trimmed like Google trims. */
  function tab(populated: (row: number) => boolean, width: number) {
    const requests: Array<[number, number]> = [];
    const lookaheads: Array<[number, number]> = [];
    const read = (log: Array<[number, number]>, cell: object) => async (from: number, to: number): Promise<RawRow[]> => {
      log.push([from, to]);
      const rows: RawRow[] = [];
      for (let r = from; r <= to; r++) {
        rows.push(populated(r) ? { values: Array.from({ length: width }, () => cell) } : {});
      }
      while (rows.length > 0 && !rows[rows.length - 1].values) rows.pop();
      return rows;
    };
    const source = { fetchRows: read(requests, { formattedValue: 'x' }), fetchPresence: read(lookaheads, {}) };
    return { requests, lookaheads, source };
  }

  const whole = (rows: number, columns: number): SheetWindow => ({
    startRow: 1, scopeEndRow: rows, startColumn: 0, columns, columnsOmitted: 0, empty: false,
  });

  it('fetches nothing for an empty window', async () => {
    const { requests, source } = tab(() => true, 3);
    const page = await collectPage({ ...whole(10, 3), empty: true }, source);
    expect(page).toEqual({ rows: [] });
    expect(requests).toEqual([]);
  });

  it('fills the page to the cell cap at the width the data has', async () => {
    const { requests, source } = tab(() => true, 3);
    const page = await collectPage(whole(100_000, 26), source);
    expect(page.rows).toHaveLength(Math.floor(MAX_CELLS / 3));
    expect(page.nextRow).toBe(1667);
    // First chunk at the allocated width plus its probe row, then the rest
    expect(requests).toEqual([[1, 193], [193, 1667]]);
  });

  it('never asks for more than MAX_FETCH_CELLS at the allocated width', async () => {
    const { requests, source } = tab(() => true, 1);
    await collectPage(whole(100_000, 256), source);
    for (const [from, to] of requests) {
      expect((to - from + 1 - 1) * 256).toBeLessThanOrEqual(MAX_FETCH_CELLS);
    }
  });

  it('stops at the end of the scope without a probe past it', async () => {
    const { requests, source } = tab(() => true, 2);
    const page = await collectPage(whole(50, 2), source);
    expect(page.nextRow).toBeUndefined();
    expect(page.rows).toHaveLength(50);
    expect(requests).toEqual([[1, 50]]);
  });

  it('settles a short chunk with presence-only look-aheads over growing, capped spans', async () => {
    const { requests, lookaheads, source } = tab((r) => r <= 40, 26);
    const page = await collectPage(whole(1_000_000, 26), source);
    expect(page.rows).toHaveLength(40);
    expect(page.nextRow).toBeUndefined();
    expect(requests).toEqual([[1, 193]]);
    const spans = lookaheads.map(([from, to]) => to - from + 1);
    expect(spans[0]).toBe(Math.floor(MAX_FETCH_CELLS / 26));
    expect(spans.slice(1, -1).every((s, i) => s >= spans[i])).toBe(true);
    expect(Math.max(...spans) * 26).toBeLessThanOrEqual(MAX_LOOKAHEAD_CELLS);
    expect(lookaheads[lookaheads.length - 1][1]).toBe(1_000_000);
  });

  it('caps a look-ahead by rows as well as cells on a narrow tab', async () => {
    const { lookaheads, source } = tab((r) => r <= 2, 1);
    await collectPage(whole(5_000_000, 1), source);
    for (const [from, to] of lookaheads) expect(to - from + 1).toBeLessThanOrEqual(MAX_LOOKAHEAD_ROWS);
  });

  it('ends the page at a long blank run and says where data resumes', async () => {
    const { source } = tab((r) => r === 1 || r === 10_000, 26);
    const page = await collectPage(whole(50_000, 26), source);
    expect(page.rows).toHaveLength(1);
    expect(page.nextRow).toBe(10_000);
  });

  it('keeps blank rows that fall inside a chunk, so paging stays row-aligned', async () => {
    const { source } = tab((r) => r === 1 || r === 5, 2);
    const page = await collectPage(whole(5, 2), source);
    expect(page.rows).toHaveLength(5);
    expect(page.rows[2]).toEqual({});
  });
});
