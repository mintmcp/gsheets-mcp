import { describe, it, expect } from 'vitest';
import { decodeGrid, type RawRow } from '../lib/cells.js';
import { MAX_CELL_CHARS } from '../lib/sheetBudget.js';

describe('decodeGrid', () => {
  it('classifies each cell type and prefers formattedValue for display', () => {
    const result = decodeGrid([
      { values: [
        { effectiveValue: { stringValue: 'hi' }, formattedValue: 'hi' },
        { effectiveValue: { numberValue: 1234.5 }, formattedValue: '1,234.50' },
        { effectiveValue: { boolValue: true }, formattedValue: 'TRUE' },
        { userEnteredValue: { formulaValue: '=SUM(A1:A2)' }, effectiveValue: { numberValue: 3 }, formattedValue: '3' },
        {},
      ] },
    ]);
    expect(result.data[0]).toEqual([
      { value: 'hi' },
      { value: '1,234.50', type: 'number' },
      { value: 'TRUE', type: 'boolean' },
      { value: '3', type: 'number', formula: '=SUM(A1:A2)' },
      { value: '', type: 'empty' },
    ]);
    expect(result.truncated).toBe(false);
  });

  it('reads a spilled cell that has no userEnteredValue', () => {
    // ARRAYFORMULA/QUERY spill into cells that carry only the computed value
    const result = decodeGrid([
      { values: [{ effectiveValue: { numberValue: 20 }, formattedValue: '20' }] },
    ]);
    expect(result.data[0][0]).toEqual({ value: '20', type: 'number' });
  });

  it('types a formula error as error, not as the text of the error', () => {
    const result = decodeGrid([
      { values: [{
        userEnteredValue: { formulaValue: '=1/0' },
        effectiveValue: { errorValue: { type: 'DIVIDE_BY_ZERO', message: 'Function DIVIDE parameter 2 cannot be zero.' } },
        formattedValue: '#DIV/0!',
      }] },
    ]);
    expect(result.data[0][0]).toEqual({ value: '#DIV/0!', type: 'error', formula: '=1/0' });
  });

  it('keeps the formula on a cell whose result is empty', () => {
    const result = decodeGrid([
      { values: [{ userEnteredValue: { formulaValue: '=IF(TRUE,,)' } }] },
    ]);
    expect(result.data[0][0]).toEqual({ value: '', type: 'empty', formula: '=IF(TRUE,,)' });
  });

  it('measures a formula cell hyperlink over the result, not the formula text', () => {
    const result = decodeGrid([
      { values: [{
        userEnteredValue: { formulaValue: '=HYPERLINK("https://g.co","Google")' },
        effectiveValue: { stringValue: 'Google' },
        formattedValue: 'Google',
        hyperlink: 'https://g.co',
      }] },
    ]);
    expect(result.data[0][0]).toEqual({
      value: 'Google',
      formula: '=HYPERLINK("https://g.co","Google")',
      hyperlinks: [{ url: 'https://g.co', start: 0, end: 6 }],
    });
  });

  it('extracts whole-cell hyperlinks', () => {
    const result = decodeGrid([
      { values: [{ effectiveValue: { stringValue: 'Google' }, formattedValue: 'Google', hyperlink: 'https://g.co' }] },
    ]);
    expect(result.data[0][0].hyperlinks).toEqual([{ url: 'https://g.co', start: 0, end: 6 }]);
  });

  it('extracts mixed-content hyperlinks from textFormatRuns', () => {
    const result = decodeGrid([
      { values: [{
        effectiveValue: { stringValue: 'Visit Google today' },
        formattedValue: 'Visit Google today',
        textFormatRuns: [
          { startIndex: 0 },
          { startIndex: 6, format: { link: { uri: 'https://g.co' } } },
          { startIndex: 12 },
        ],
      }] },
    ]);
    expect(result.data[0][0].hyperlinks).toEqual([{ url: 'https://g.co', start: 6, end: 12 }]);
  });

  it('stops at the cell budget and reports truncated', () => {
    const rowData: RawRow[] = Array.from({ length: 100 }, () => ({
      values: Array.from({ length: 10 }, () => ({ effectiveValue: { stringValue: 'x' } })),
    }));
    const result = decodeGrid(rowData, { maxCells: 25 });
    expect(result.truncated).toBe(true);
    expect(result.data.flat().length).toBeLessThanOrEqual(25);
  });

  it('stops at the char budget and reports truncated', () => {
    const big = 'y'.repeat(1000);
    const rowData: RawRow[] = Array.from({ length: 50 }, () => ({
      values: [{ effectiveValue: { stringValue: big }, formattedValue: big }],
    }));
    const result = decodeGrid(rowData, { maxChars: 5_000 });
    expect(result.truncated).toBe(true);
    expect(result.data.length).toBeLessThan(50);
  });

  it('clips a single oversized cell value', () => {
    const huge = 'z'.repeat(MAX_CELL_CHARS + 500);
    const result = decodeGrid([{ values: [{ effectiveValue: { stringValue: huge }, formattedValue: huge }] }]);
    expect(result.data[0][0].value).toHaveLength(MAX_CELL_CHARS);
  });

  it('flags a clipped value so it cannot read as complete', () => {
    const huge = 'z'.repeat(MAX_CELL_CHARS + 500);
    const result = decodeGrid([{ values: [{ effectiveValue: { stringValue: huge }, formattedValue: huge }] }]);
    expect(result.data[0][0].valueShortened).toBe(true);
  });

  it('leaves the flag off a value that fit', () => {
    const result = decodeGrid([{ values: [{ effectiveValue: { stringValue: 'ok' }, formattedValue: 'ok' }] }]);
    expect(result.data[0][0].valueShortened).toBeUndefined();
  });

  it('reports dimensions and handles empty input', () => {
    expect(decodeGrid([])).toEqual({
      data: [], rowCount: 0, columnCount: 0, truncated: false, partialRow: false,
    });
    const result = decodeGrid([{ values: [{}, {}, {}] }, { values: [{}] }]);
    expect(result.rowCount).toBe(2);
    expect(result.columnCount).toBe(3);
  });

  it('never reports a hyperlink range past the clipped value', () => {
    const huge = 'z'.repeat(MAX_CELL_CHARS + 500);
    const result = decodeGrid([
      { values: [{ effectiveValue: { stringValue: huge }, formattedValue: huge, hyperlink: 'https://g.co' }] },
    ]);
    const cell = result.data[0][0];
    expect(cell.value).toHaveLength(MAX_CELL_CHARS);
    for (const link of cell.hyperlinks ?? []) {
      expect(link.end).toBeLessThanOrEqual(cell.value.length);
      expect(link.start).toBeLessThanOrEqual(link.end);
    }
  });

  it('drops textFormatRuns links that fall entirely past the text', () => {
    const result = decodeGrid([
      { values: [{
        effectiveValue: { stringValue: 'short' },
        formattedValue: 'short',
        textFormatRuns: [
          { startIndex: 0, format: { link: { uri: 'https://in.co' } } },
          { startIndex: 900, format: { link: { uri: 'https://out.co' } } },
        ],
      }] },
    ]);
    const urls = (result.data[0][0].hyperlinks ?? []).map((l) => l.url);
    expect(urls).toEqual(['https://in.co']);
  });

  it('counts hyperlink structure against the char budget', () => {
    const linked = {
      effectiveValue: { stringValue: 'a' },
      formattedValue: 'a',
      hyperlink: 'https://example.com/' + 'p'.repeat(200),
    };
    const rowData: RawRow[] = Array.from({ length: 100 }, () => ({ values: [linked] }));
    const result = decodeGrid(rowData, { maxChars: 2_000 });
    expect(result.truncated).toBe(true);
    expect(result.data.length).toBeLessThan(20);
  });

  it('counts the formula text against the char budget', () => {
    const cell = {
      userEnteredValue: { formulaValue: '=' + 'A1+'.repeat(300) + 'A1' },
      effectiveValue: { numberValue: 301 },
      formattedValue: '301',
    };
    const rowData: RawRow[] = Array.from({ length: 100 }, () => ({ values: [cell] }));
    const result = decodeGrid(rowData, { maxChars: 5_000 });
    expect(result.truncated).toBe(true);
    expect(result.data.length).toBeLessThan(10);
  });

  it('omits a formula longer than the cell cap rather than emit a truncated one', () => {
    // A clipped formula written back would corrupt the sheet; the value still stands on its own
    const result = decodeGrid([
      { values: [{
        userEnteredValue: { formulaValue: '=' + 'A1+'.repeat(20) + 'A1' },
        effectiveValue: { numberValue: 21 },
        formattedValue: '21',
      }] },
    ], { maxCellChars: 30 });
    expect(result.data[0][0]).toEqual({ value: '21', type: 'number' });
  });

  it('handles rows with no values array', () => {
    const result = decodeGrid([{}, { values: [{ effectiveValue: { stringValue: 'a' } }] }]);
    expect(result.data).toEqual([[], [{ value: 'a' }]]);
    expect(result.rowCount).toBe(2);
  });
});

describe('decodeGrid row alignment under truncation', () => {
  it('drops a partial row so paging can resume cleanly at the next row', () => {
    const rowData: RawRow[] = Array.from({ length: 10 }, () => ({
      values: Array.from({ length: 4 }, () => ({ effectiveValue: { stringValue: 'x' } })),
    }));
    // 10 cells = two full rows (8 cells) plus a partial third.
    const result = decodeGrid(rowData, { maxCells: 10 });
    expect(result.truncated).toBe(true);
    expect(result.partialRow).toBe(false);
    for (const row of result.data) expect(row).toHaveLength(4);
  });

  it('keeps a single oversized row rather than returning nothing', () => {
    const big = 'b'.repeat(30_000);
    const rowData: RawRow[] = [{
      values: Array.from({ length: 50 }, () => ({ effectiveValue: { stringValue: big }, formattedValue: big })),
    }];
    const result = decodeGrid(rowData, { maxChars: 100_000 });
    expect(result.rowCount).toBe(1);
    expect(result.partialRow).toBe(true);
    expect(result.truncated).toBe(true);
  });
});

describe('decodeGrid stays inside the character budget', () => {
  const row = (cells: Array<{ v: string; num?: boolean }>) => ({
    values: cells.map((c) => ({
      effectiveValue: c.num ? { numberValue: Number(c.v) } : { stringValue: c.v },
      formattedValue: c.v,
    })),
  });

  it('never returns more characters than the budget allows', () => {
    const rows = Array.from({ length: 500 }, () =>
      row(Array.from({ length: 20 }, (_, c) => ({ v: String(c), num: true }))));
    const result = decodeGrid(rows, { maxChars: 20_000 });

    expect(JSON.stringify(result.data).length).toBeLessThanOrEqual(20_000);
    expect(result.truncated).toBe(true);
  });

  it('does not let one oversized cell overshoot the budget', () => {
    const rows = [row([{ v: 'a'.repeat(50) }]), row([{ v: 'b'.repeat(5_000) }])];
    const result = decodeGrid(rows, { maxChars: 200 });

    expect(JSON.stringify(result.data).length).toBeLessThanOrEqual(200);
  });

  it('returns the first cell even when it alone exceeds the budget', () => {
    const result = decodeGrid([row([{ v: 'c'.repeat(5_000) }])], { maxChars: 100 });
    expect(result.rowCount).toBe(1);
    expect(result.data[0]).toHaveLength(1);
  });
});
