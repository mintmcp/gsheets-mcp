import { describe, it, expect } from 'vitest';
import { normalizeRange } from '../lib/a1.js';

/**
 * The official connector takes "Sheet1!A1:B2" and R1C1 wherever it takes a
 * range. normalizeRange accepts both and reduces them to the bare bounded A1
 * the rest of the server works in, plus the tab the prefix named.
 */
describe('normalizeRange', () => {
  it('passes a bare bounded range through', () => {
    expect(normalizeRange('A1:C3')).toEqual({ a1: 'A1:C3' });
    expect(normalizeRange(' B2 ')).toEqual({ a1: 'B2' });
  });

  it('splits off an unquoted sheet prefix', () => {
    expect(normalizeRange('Data!A1:C3')).toEqual({ sheetName: 'Data', a1: 'A1:C3' });
  });

  it('unquotes a quoted sheet prefix, including doubled quotes', () => {
    expect(normalizeRange("'Q3 Sales'!A1")).toEqual({ sheetName: 'Q3 Sales', a1: 'A1' });
    expect(normalizeRange("'Bob''s'!A1:B2")).toEqual({ sheetName: "Bob's", a1: 'A1:B2' });
  });

  it('converts absolute R1C1 to A1', () => {
    expect(normalizeRange('R1C1:R4C4')).toEqual({ a1: 'A1:D4' });
    expect(normalizeRange('Data!R2C27')).toEqual({ sheetName: 'Data', a1: 'AA2' });
  });

  it('rejects relative R1C1, which has no fixed meaning outside a cell', () => {
    expect(() => normalizeRange('R[1]C[1]')).toThrow(/relative/i);
  });

  it('rejects a prefix naming a tab other than the one given', () => {
    expect(() => normalizeRange('Other!A1', { sheetName: 'Data' })).toThrow(/Other.*Data/);
  });

  it('accepts a prefix that matches the tab given', () => {
    expect(normalizeRange('Data!A1', { sheetName: 'Data' })).toEqual({ sheetName: 'Data', a1: 'A1' });
  });

  it('rejects a bare whole sheet reference without a prefix form', () => {
    expect(() => normalizeRange('Data!')).toThrow(/A1/);
  });

  it('keeps bounded-only rules by default', () => {
    expect(() => normalizeRange('A:C')).toThrow(/whole-column/i);
    expect(() => normalizeRange('Data!1:3')).toThrow(/whole-row/i);
  });

  it('accepts whole-column and whole-row forms when allowed', () => {
    expect(normalizeRange('A:C', { allowOpen: true })).toEqual({ a1: 'A:C' });
    expect(normalizeRange('Data!1:3', { allowOpen: true })).toEqual({ sheetName: 'Data', a1: '1:3' });
    // column R, not whole rows: A1 wins over R1C1 for the ambiguous pair
    expect(normalizeRange('R1:R3', { allowOpen: true })).toEqual({ a1: 'R1:R3' });
    expect(normalizeRange('C1:C3')).toEqual({ a1: 'C1:C3' });
    expect(normalizeRange('C1:C3', { allowOpen: true })).toEqual({ a1: 'C1:C3' });
    expect(normalizeRange('B2:D', { allowOpen: true })).toEqual({ a1: 'B2:D' });
  });

  it('names the parameter in errors', () => {
    expect(() => normalizeRange('', { paramName: 'ranges[0]' })).toThrow(/ranges\[0\]/);
    expect(() => normalizeRange(5, { paramName: 'ranges[0]' })).toThrow(/ranges\[0\]/);
  });

  it('rejects reversed endpoints and row 0 as before', () => {
    expect(() => normalizeRange('C3:A1')).toThrow(/reversed/);
    expect(() => normalizeRange('A0')).toThrow(/row 0/);
  });

  it('rejects reversed half-open ranges too', () => {
    expect(() => normalizeRange('D2:B', { allowOpen: true })).toThrow(/reversed/);
    expect(() => normalizeRange('B9:3', { allowOpen: true })).toThrow(/reversed/);
    expect(normalizeRange('B2:9', { allowOpen: true })).toEqual({ a1: 'B2:9' });
  });
});
