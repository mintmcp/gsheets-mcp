import { describe, it, expect } from 'vitest';
import {
  assertBareA1Range,
  assertSingleCell,
  parseA1Range,
  columnLetterToIndex,
  quoteSheetName,
} from '../lib/a1.js';

describe('columnLetterToIndex', () => {
  it('maps single letters', () => {
    expect(columnLetterToIndex('A')).toBe(0);
    expect(columnLetterToIndex('B')).toBe(1);
    expect(columnLetterToIndex('Z')).toBe(25);
  });
  it('maps double letters', () => {
    expect(columnLetterToIndex('AA')).toBe(26);
    expect(columnLetterToIndex('AZ')).toBe(51);
  });
  it('is case-insensitive', () => {
    expect(columnLetterToIndex('a')).toBe(0);
    expect(columnLetterToIndex('aa')).toBe(26);
  });
});

describe('quoteSheetName', () => {
  it('wraps in single quotes', () => {
    expect(quoteSheetName('Sheet1')).toBe("'Sheet1'");
  });
  it('doubles existing single quotes', () => {
    expect(quoteSheetName("Bob's Sheet")).toBe("'Bob''s Sheet'");
  });
});

describe('assertBareA1Range', () => {
  it('accepts a single cell', () => {
    expect(assertBareA1Range('A1')).toBe('A1');
    expect(assertBareA1Range('Z99')).toBe('Z99');
    expect(assertBareA1Range('AA1')).toBe('AA1');
  });
  it('accepts a bounded range', () => {
    expect(assertBareA1Range('A1:C3')).toBe('A1:C3');
    expect(assertBareA1Range('B2:B2')).toBe('B2:B2');
  });
  it('trims whitespace', () => {
    expect(assertBareA1Range('  A1:C3  ')).toBe('A1:C3');
  });
  it('rejects non-string input', () => {
    expect(() => assertBareA1Range(undefined)).toThrow(/string/i);
    expect(() => assertBareA1Range(42)).toThrow(/string/i);
    expect(() => assertBareA1Range(null)).toThrow(/string/i);
  });
  it('rejects empty / whitespace-only', () => {
    expect(() => assertBareA1Range('')).toThrow(/non-empty/i);
    expect(() => assertBareA1Range('   ')).toThrow(/non-empty/i);
  });
  it('rejects sheet-qualified strings', () => {
    expect(() => assertBareA1Range('Sheet1!A1:C3')).toThrow(/bare A1/i);
    expect(() => assertBareA1Range("'My Sheet'!A1")).toThrow(/bare A1/i);
  });
  it('rejects open-ended forms', () => {
    expect(() => assertBareA1Range('A:C')).toThrow(/bounded A1/i);
    expect(() => assertBareA1Range('1:3')).toThrow(/bounded A1/i);
    expect(() => assertBareA1Range('A1:C')).toThrow(/bounded A1/i);
    expect(() => assertBareA1Range('A:C3')).toThrow(/bounded A1/i);
  });
  it('rejects row 0', () => {
    expect(() => assertBareA1Range('A0')).toThrow(/row 0/);
    expect(() => assertBareA1Range('A0:B5')).toThrow(/row 0/);
    expect(() => assertBareA1Range('A1:B0')).toThrow(/row 0/);
  });
  it('rejects reversed endpoints', () => {
    expect(() => assertBareA1Range('B2:A1')).toThrow(/reversed/);
    expect(() => assertBareA1Range('C3:A1')).toThrow(/reversed/);
    expect(() => assertBareA1Range('A5:A1')).toThrow(/reversed/);
    expect(() => assertBareA1Range('C1:A1')).toThrow(/reversed/);
  });
  it('rejects garbage', () => {
    expect(() => assertBareA1Range('not-a-range')).toThrow(/bounded A1/i);
    expect(() => assertBareA1Range('1A')).toThrow(/bounded A1/i);
    expect(() => assertBareA1Range('A1:B2:C3')).toThrow(/bounded A1/i);
  });
  it('uses the paramName in error messages', () => {
    expect(() => assertBareA1Range('A:C', 'ranges[]')).toThrow(/ranges\[\]/);
  });
});

describe('assertSingleCell', () => {
  it('accepts single cells', () => {
    expect(assertSingleCell('A1')).toBe('A1');
    expect(assertSingleCell('AA1')).toBe('AA1');
    expect(assertSingleCell('B3')).toBe('B3');
  });
  it('rejects ranges', () => {
    expect(() => assertSingleCell('A1:B2')).toThrow(/single cell/i);
  });
  it('rejects non-string', () => {
    expect(() => assertSingleCell(undefined)).toThrow(/single cell/i);
  });
  it('rejects malformed cells', () => {
    expect(() => assertSingleCell('1A')).toThrow(/Invalid A1 cell/);
    expect(() => assertSingleCell('not-a-cell')).toThrow(/Invalid A1 cell/);
    expect(() => assertSingleCell('')).toThrow(/Invalid A1 cell/);
  });
});

describe('parseA1Range', () => {
  it('parses a single cell', () => {
    expect(parseA1Range('A1')).toEqual({
      startRowIndex: 0,
      endRowIndex: 1,
      startColumnIndex: 0,
      endColumnIndex: 1,
    });
  });
  it('parses a range', () => {
    expect(parseA1Range('A1:C3')).toEqual({
      startRowIndex: 0,
      endRowIndex: 3,
      startColumnIndex: 0,
      endColumnIndex: 3,
    });
  });
  it('parses double-letter columns', () => {
    expect(parseA1Range('AA1:AB2')).toEqual({
      startRowIndex: 0,
      endRowIndex: 2,
      startColumnIndex: 26,
      endColumnIndex: 28,
    });
  });
});
