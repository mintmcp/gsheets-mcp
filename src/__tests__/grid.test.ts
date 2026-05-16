import { describe, it, expect } from 'vitest';
import { padRaggedRows } from '../lib/grid.js';

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
