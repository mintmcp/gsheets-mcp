import { describe, it, expect } from 'vitest';
import { parseXlsx, toCell } from '../lib/xlsx.js';
import { fixture } from './__fixtures__/index.js';

describe('toCell', () => {
  it('maps each SheetJS cell type to the native type enum', () => {
    expect(toCell({ t: 's', v: 'Region', w: 'Region' })).toEqual({ value: 'Region' });
    expect(toCell({ t: 'n', v: 1284000, w: '$1,284,000.00' }))
      .toEqual({ value: '$1,284,000.00', type: 'number' });
    expect(toCell({ t: 'b', v: true, w: 'TRUE' })).toEqual({ value: 'TRUE', type: 'boolean' });
    expect(toCell({ t: 'e', v: 0x17, w: '#REF!' })).toEqual({ value: '#REF!' });
    expect(toCell(undefined)).toEqual({ value: '', type: 'empty' });
  });

  it('returns formulas with a leading = like the Sheets API does', () => {
    expect(toCell({ t: 'n', v: 4285400, w: '$4,285,400.00', f: 'SUM(B2:B4)' }))
      .toEqual({ value: '=SUM(B2:B4)', type: 'formula' });
  });

  it('prefers the rendered text so a date is never a raw serial or a local-time Date', () => {
    expect(toCell({ t: 'd', v: new Date(Date.UTC(2026, 7, 14)), w: '8/14/26' }))
      .toEqual({ value: '8/14/26', type: 'number' });
  });

  it('attaches a hyperlink over the display text', () => {
    expect(toCell({
      t: 's', v: 'Northwind contract', w: 'Northwind contract',
      l: { Target: 'https://example.com/x' },
    })).toEqual({
      value: 'Northwind contract',
      hyperlinks: [{ url: 'https://example.com/x', start: 0, end: 18 }],
    });
  });

  it('measures a formula cell hyperlink over the cached result, not the formula text', () => {
    const cell = toCell({
      t: 's', v: 'hi', w: 'hi', f: 'CONCAT(B1,C1)',
      l: { Target: 'https://example.com' },
    });
    expect(cell.value).toBe('=CONCAT(B1,C1)');
    expect(cell.hyperlinks).toEqual([{ url: 'https://example.com', start: 0, end: 2 }]);
  });
});

describe('parseXlsx', () => {
  it('reads sheet names and cell values from a real workbook', () => {
    const wb = parseXlsx(fixture('basic.xlsx'));
    expect(wb.sheets).toHaveLength(1);
    expect(wb.sheets[0].name).toBe('Q3 Revenue');
    expect(wb.sheets[0].data[0][0]).toEqual({ value: 'Region' });
    expect(wb.truncated).toBe(false);
  });

  it('applies number formats and formulas from the formats fixture', () => {
    const [sheet] = parseXlsx(fixture('formats.xlsx')).sheets;
    expect(sheet.data[1][1]).toEqual({ value: '$1,284,000.00', type: 'number' });  // B2
    expect(sheet.data[1][2]).toEqual({ value: '12.0%', type: 'number' });          // C2
    expect(sheet.data[4][1]).toEqual({ value: '=SUM(B2:B4)', type: 'formula' });   // B5
  });

  it('renders a date through the file own format code, never a serial', () => {
    const [sheet] = parseXlsx(fixture('formats.xlsx')).sheets;
    const d2 = sheet.data[1][3];
    expect(d2.type).toBe('number');
    expect(d2.value).not.toBe('46248');
    expect(d2.value).toMatch(/8.14.26/);   // the fixture's own mm-dd-yy code
  });

  it('resolves whole-cell hyperlinks', () => {
    const [sheet] = parseXlsx(fixture('formats.xlsx')).sheets;
    expect(sheet.data[1][4].hyperlinks).toEqual([
      { url: 'https://example.com/northwind-contract', start: 0, end: 18 },
    ]);
  });

  it('reports rowCount and columnCount', () => {
    const [sheet] = parseXlsx(fixture('formats.xlsx')).sheets;
    expect(sheet.rowCount).toBe(sheet.data.length);
    expect(sheet.columnCount).toBe(5);
  });

  it('keeps row numbers aligned with the spreadsheet the user sees', () => {
    const [sheet] = parseXlsx(fixture('formats.xlsx')).sheets;
    // Row 7 in the file carries the marker, so it must land at index 6.
    expect(sheet.data[6][0].value).toContain('MAGIC-PHRASE-XLSX');
  });

  it('flags truncation when the cell cap is hit', () => {
    const wb = parseXlsx(fixture('formats.xlsx'), { maxCells: 3 });
    expect(wb.truncated).toBe(true);
    expect(wb.cells).toBeLessThanOrEqual(3);
  });

  it('reads only tab names in namesOnly mode', () => {
    const wb = parseXlsx(fixture('formats.xlsx'), { namesOnly: true });
    expect(wb.sheets.map((s) => s.name)).toEqual(['Q3 Revenue']);
    expect(wb.cells).toBe(0);
    expect(wb.sheets[0].data).toEqual([]);
  });

  it('reports the totals it emitted', () => {
    const wb = parseXlsx(fixture('formats.xlsx'));
    expect(wb.cells).toBeGreaterThan(0);
    expect(wb.chars).toBeGreaterThan(0);
  });
});
