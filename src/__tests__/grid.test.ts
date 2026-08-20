import { describe, it, expect } from 'vitest';
import { maxRowLength, padRaggedRows, MAX_PADDED_CELLS } from '../lib/grid.js';

describe('maxRowLength', () => {
  it('returns 0 for no rows', () => {
    expect(maxRowLength([])).toBe(0);
  });

  it('returns the widest row length', () => {
    expect(maxRowLength([['a'], ['b', 'c', 'd'], ['e', 'f']])).toBe(3);
  });

  it('handles more rows than the call-stack spread limit', () => {
    const rows = Array.from({ length: 200_000 }, () => ['a']);
    expect(maxRowLength(rows)).toBe(1);
  });
});

describe('padRaggedRows on very tall input', () => {
  it('does not overflow the call stack', () => {
    const rows: string[][] = Array.from({ length: 200_000 }, (_, i) =>
      i === 0 ? ['a', 'b'] : ['a'],
    );
    const padded = padRaggedRows(rows);
    expect(padded).toHaveLength(200_000);
    expect(padded[1]).toEqual(['a', '']);
  });
});

describe('padRaggedRows explosion guard', () => {
  it('rejects a sparse matrix small enough to send but huge once padded', () => {
    // 50,000 cells in a body under 1MB, so nothing upstream of here objects,
    // and padding turns it into 625 million entries.
    const rows = [...Array.from({ length: 25_000 }, () => ['a']),
      Array.from({ length: 25_000 }, () => 'b')];
    expect(rows.reduce((n, r) => n + r.length, 0)).toBe(50_000);
    expect(() => padRaggedRows(rows)).toThrow(/pads out to 625025000 cells/);
  });

  it('leaves an ordinary ragged write alone', () => {
    // 8,000 rows averaging 5 cells with one 8-wide header: 64,000 padded,
    // which used to be rejected and is nowhere near the memory cliff.
    const rows = [Array.from({ length: 8 }, () => 'h'),
      ...Array.from({ length: 7_999 }, () => ['a', 'b', 'c', 'd', 'e'])];
    expect(rows.length * 8).toBeLessThan(MAX_PADDED_CELLS);
    expect(padRaggedRows(rows)).toHaveLength(8_000);
  });
});

describe('padRaggedRows', () => {
  it('passes through uniform rows untouched', () => {
    const input = [['a', 'b'], ['c', 'd']];
    expect(padRaggedRows(input)).toEqual(input);
  });

  it('pads short rows with empty strings to match the widest row', () => {
    const input = [['a', 'b', 'c'], ['d']];
    expect(padRaggedRows(input)).toEqual([
      ['a', 'b', 'c'],
      ['d', '', ''],
    ]);
  });

  it('does not mutate the input', () => {
    const input = [['a'], ['b', 'c']];
    const snapshot = JSON.parse(JSON.stringify(input));
    padRaggedRows(input);
    expect(input).toEqual(snapshot);
  });

  it('throws on empty input', () => {
    expect(() => padRaggedRows([])).toThrow(/at least one row/);
  });

  it('throws on non-array input', () => {
    expect(() => padRaggedRows(undefined)).toThrow(/at least one row/);
    expect(() => padRaggedRows('not-array')).toThrow(/at least one row/);
  });

  it('throws when no row has any cells', () => {
    expect(() => padRaggedRows([[]])).toThrow(/at least one cell/);
    expect(() => padRaggedRows([[], [], []])).toThrow(/at least one cell/);
  });

  it('throws when a row is not an array', () => {
    expect(() => padRaggedRows([['a'], 'not-a-row'])).toThrow(/2D array/);
  });
});
