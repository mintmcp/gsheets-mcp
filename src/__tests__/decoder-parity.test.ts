import { describe, it, expect } from 'vitest';
import { decodeGrid } from '../lib/cells.js';
import { toCell } from '../lib/xlsx.js';
import {
  createBudget,
  chargeCell,
  MAX_CELLS,
  MAX_CELL_CHARS,
  MAX_OUTPUT_CHARS,
} from '../lib/sheetBudget.js';

/**
 * `get_sheet_data` decodes native grids and .xlsx workbooks through separate
 * modules that feed one output schema. These previously drifted — the
 * character envelope disagreed (32 vs 50), so the same documented cap meant
 * two different things. Guard the contract they must share.
 */

describe('native and xlsx decoders share one contract', () => {
  it('produce the same cell shape for equivalent content', () => {
    const native = decodeGrid([
      { values: [
        { userEnteredValue: { stringValue: 'hi' }, formattedValue: 'hi' },
        { userEnteredValue: { boolValue: true }, formattedValue: 'TRUE' },
        { userEnteredValue: { formulaValue: '=A1' }, formattedValue: '3' },
      ] },
    ]).data[0];

    const xlsx = [
      toCell({ t: 's', v: 'hi', w: 'hi' } as any),
      toCell({ t: 'b', v: true, w: 'TRUE' } as any),
      toCell({ t: 'n', v: 3, w: '3', f: 'A1' } as any),
    ];

    expect(native.map((c) => Object.keys(c).sort()))
      .toEqual(xlsx.map((c) => Object.keys(c).sort()));
    expect(native.map((c) => c.type)).toEqual(xlsx.map((c) => c.type));
    expect(native[0].value).toBe(xlsx[0].value);
    expect(native[1].value).toBe(xlsx[1].value);
  });

  it('clip an oversized value to the same length', () => {
    const huge = 'z'.repeat(MAX_CELL_CHARS + 1_000);
    const native = decodeGrid([
      { values: [{ userEnteredValue: { stringValue: huge }, formattedValue: huge }] },
    ]).data[0][0];
    const xlsx = toCell({ t: 's', v: huge, w: huge } as any);

    expect(native.value.length).toBe(MAX_CELL_CHARS);
    expect(xlsx.value.length).toBe(MAX_CELL_CHARS);
  });

  it('charge an identical cell identically against the budget', () => {
    const cell = { value: 'abc', type: 'string' as const };
    const a = createBudget();
    const b = createBudget();
    chargeCell(a, cell);
    chargeCell(b, cell);
    expect(a.chars).toBe(b.chars);
    expect(a.chars).toBeGreaterThan(cell.value.length);
  });

  it.each([
    'javascript:alert(1)',
    'JaVaScRiPt:alert(1)',
    'data:text/html;base64,PHNjcmlwdD4=',
    'vbscript:msgbox(1)',
    'file:///etc/passwd',
  ])('drop unsafe hyperlink target %s on BOTH paths', (url) => {
    const native = decodeGrid([
      { values: [{ userEnteredValue: { stringValue: 'x' }, formattedValue: 'x', hyperlink: url }] },
    ]).data[0][0];
    const xlsx = toCell({ t: 's', v: 'x', w: 'x', l: { Target: url } } as any);

    expect(native.hyperlinks).toBeUndefined();
    expect(xlsx.hyperlinks).toBeUndefined();
  });

  it.each(['https://example.com', 'http://example.com', 'mailto:a@b.co'])(
    'keep safe hyperlink target %s on BOTH paths',
    (url) => {
      const native = decodeGrid([
        { values: [{ userEnteredValue: { stringValue: 'x' }, formattedValue: 'x', hyperlink: url }] },
      ]).data[0][0];
      const xlsx = toCell({ t: 's', v: 'x', w: 'x', l: { Target: url } } as any);

      expect(native.hyperlinks?.[0].url).toBe(url);
      expect(xlsx.hyperlinks?.[0].url).toBe(url);
    },
  );

  it('drop an unsafe target hidden in a native textFormatRun', () => {
    const cell = decodeGrid([
      { values: [{
        userEnteredValue: { stringValue: 'click me' },
        formattedValue: 'click me',
        textFormatRuns: [
          { startIndex: 0, format: { link: { uri: 'javascript:alert(1)' } } },
          { startIndex: 5, format: { link: { uri: 'https://ok.example' } } },
        ],
      }] },
    ]).data[0][0];

    expect((cell.hyperlinks ?? []).map((l) => l.url)).toEqual(['https://ok.example']);
  });

  it('clip an over-long hyperlink target to the same length on BOTH paths', () => {
    const url = `https://example.com/${'u'.repeat(MAX_CELL_CHARS + 1_000)}`;
    const native = decodeGrid([
      { values: [{ userEnteredValue: { stringValue: 'x' }, formattedValue: 'x', hyperlink: url }] },
    ]).data[0][0];
    const xlsx = toCell({ t: 's', v: 'x', w: 'x', l: { Target: url } } as any);

    expect(native.hyperlinks?.[0].url).toHaveLength(MAX_CELL_CHARS);
    expect(xlsx.hyperlinks?.[0].url).toHaveLength(MAX_CELL_CHARS);
  });

  it('omit type for plain strings and keep it everywhere else, on BOTH paths', () => {
    // `string` is the fallback branch of both decoders, so emitting it says
    // nothing while costing ~17 bytes on the cells that dominate a sheet.
    const native = decodeGrid([
      { values: [
        { userEnteredValue: { stringValue: 'text' }, formattedValue: 'text' },
        { userEnteredValue: { numberValue: 1234 }, formattedValue: '$1,234.00' },
        { userEnteredValue: { formulaValue: '=A1' }, formattedValue: '3' },
        { userEnteredValue: { boolValue: true }, formattedValue: 'TRUE' },
        {},
      ] },
    ]).data[0];

    expect(native.map((c) => c.type)).toEqual([
      undefined, 'number', 'formula', 'boolean', 'empty',
    ]);

    expect(toCell({ t: 's', v: 'text', w: 'text' } as any).type).toBeUndefined();
    expect(toCell({ t: 'n', v: 1234, w: '$1,234.00' } as any).type).toBe('number');
    expect(toCell(undefined).type).toBe('empty');
  });

  it('resolve the same default limits', () => {
    const budget = createBudget();
    expect(budget.maxCells).toBe(MAX_CELLS);
    expect(budget.maxChars).toBe(MAX_OUTPUT_CHARS);
    expect(budget.maxCellChars).toBe(MAX_CELL_CHARS);
  });
});
