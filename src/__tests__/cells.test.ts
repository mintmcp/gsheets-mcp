import { describe, it, expect } from 'vitest';
import { decodeGrid, type RawRow } from '../lib/cells.js';
import { MAX_CELL_CHARS } from '../lib/sheetBudget.js';

describe('decodeGrid', () => {
  it('classifies each cell type and prefers formattedValue for display', () => {
    const result = decodeGrid([
      { values: [
        { userEnteredValue: { stringValue: 'hi' }, formattedValue: 'hi' },
        { userEnteredValue: { numberValue: 1234.5 }, formattedValue: '1,234.50' },
        { userEnteredValue: { boolValue: true }, formattedValue: 'TRUE' },
        { userEnteredValue: { formulaValue: '=SUM(A1:A2)' } },
        {},
      ] },
    ]);
    expect(result.data[0]).toEqual([
      { value: 'hi', type: 'string' },
      { value: '1,234.50', type: 'number' },
      { value: 'TRUE', type: 'boolean' },
      { value: '=SUM(A1:A2)', type: 'formula' },
      { value: '', type: 'empty' },
    ]);
    expect(result.truncated).toBe(false);
  });

  it('extracts whole-cell hyperlinks', () => {
    const result = decodeGrid([
      { values: [{ userEnteredValue: { stringValue: 'Google' }, formattedValue: 'Google', hyperlink: 'https://g.co' }] },
    ]);
    expect(result.data[0][0].hyperlinks).toEqual([{ url: 'https://g.co', start: 0, end: 6 }]);
  });

  it('extracts mixed-content hyperlinks from textFormatRuns', () => {
    const result = decodeGrid([
      { values: [{
        userEnteredValue: { stringValue: 'Visit Google today' },
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
      values: Array.from({ length: 10 }, () => ({ userEnteredValue: { stringValue: 'x' } })),
    }));
    const result = decodeGrid(rowData, { maxCells: 25 });
    expect(result.truncated).toBe(true);
    expect(result.data.flat().length).toBeLessThanOrEqual(25);
  });

  it('stops at the char budget and reports truncated', () => {
    const big = 'y'.repeat(1000);
    const rowData: RawRow[] = Array.from({ length: 50 }, () => ({
      values: [{ userEnteredValue: { stringValue: big }, formattedValue: big }],
    }));
    const result = decodeGrid(rowData, { maxChars: 5_000 });
    expect(result.truncated).toBe(true);
    expect(result.data.length).toBeLessThan(50);
  });

  it('clips a single oversized cell value', () => {
    const huge = 'z'.repeat(MAX_CELL_CHARS + 500);
    const result = decodeGrid([{ values: [{ userEnteredValue: { stringValue: huge }, formattedValue: huge }] }]);
    expect(result.data[0][0].value).toHaveLength(MAX_CELL_CHARS);
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
      { values: [{ userEnteredValue: { stringValue: huge }, formattedValue: huge, hyperlink: 'https://g.co' }] },
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
        userEnteredValue: { stringValue: 'short' },
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
      userEnteredValue: { stringValue: 'a' },
      formattedValue: 'a',
      hyperlink: 'https://example.com/' + 'p'.repeat(200),
    };
    const rowData: RawRow[] = Array.from({ length: 100 }, () => ({ values: [linked] }));
    const result = decodeGrid(rowData, { maxChars: 2_000 });
    expect(result.truncated).toBe(true);
    expect(result.data.length).toBeLessThan(20);
  });

  it('handles rows with no values array', () => {
    const result = decodeGrid([{}, { values: [{ userEnteredValue: { stringValue: 'a' } }] }]);
    expect(result.data).toEqual([[], [{ value: 'a', type: 'string' }]]);
    expect(result.rowCount).toBe(2);
  });
});

describe('decodeGrid row alignment under truncation', () => {
  it('drops a partial row so paging can resume cleanly at the next row', () => {
    const rowData: RawRow[] = Array.from({ length: 10 }, () => ({
      values: Array.from({ length: 4 }, () => ({ userEnteredValue: { stringValue: 'x' } })),
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
      values: Array.from({ length: 50 }, () => ({ userEnteredValue: { stringValue: big }, formattedValue: big })),
    }];
    const result = decodeGrid(rowData, { maxChars: 100_000 });
    expect(result.rowCount).toBe(1);
    expect(result.partialRow).toBe(true);
    expect(result.truncated).toBe(true);
  });
});
