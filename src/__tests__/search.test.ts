import { describe, it, expect } from 'vitest';
import { buildDriveSearchQuery } from '../lib/search.js';

const BASE = "mimeType = 'application/vnd.google-apps.spreadsheet'";
const TRASHED = ' and trashed = false';

describe('buildDriveSearchQuery', () => {
  it('omits the name clause when name is undefined', () => {
    expect(buildDriveSearchQuery(undefined)).toBe(BASE + TRASHED);
  });

  it('omits the name clause when name is blank after trim', () => {
    expect(buildDriveSearchQuery('   ')).toBe(BASE + TRASHED);
  });

  it('includes a contains clause when name is provided', () => {
    expect(buildDriveSearchQuery('budget')).toBe(
      BASE + " and name contains 'budget'" + TRASHED,
    );
  });

  it('trims the name', () => {
    expect(buildDriveSearchQuery('  budget  ')).toBe(
      BASE + " and name contains 'budget'" + TRASHED,
    );
  });

  it("escapes single quotes inside the name", () => {
    expect(buildDriveSearchQuery("Bob's budget")).toBe(
      BASE + " and name contains 'Bob\\'s budget'" + TRASHED,
    );
  });

  it('escapes backslashes BEFORE quotes (order matters)', () => {
    // Input contains a single backslash and a single quote.
    // Expected: backslash becomes \\, then ' becomes \', and the doubled
    // backslash from step 1 is NOT touched by step 2.
    const q = buildDriveSearchQuery("a\\b'c");
    expect(q).toBe(BASE + " and name contains 'a\\\\b\\'c'" + TRASHED);
  });

  it('rejects control characters (newline)', () => {
    expect(() => buildDriveSearchQuery('a\nb')).toThrow(/control character/i);
  });

  it('rejects control characters (carriage return)', () => {
    expect(() => buildDriveSearchQuery('a\rb')).toThrow(/control character/i);
  });

  it('rejects ASCII NUL', () => {
    expect(() => buildDriveSearchQuery('a\x00b')).toThrow(/control character/i);
  });

  it('rejects other C0 controls', () => {
    expect(() => buildDriveSearchQuery('a\x07b')).toThrow(/control character/i);
    expect(() => buildDriveSearchQuery('a\x1Fb')).toThrow(/control character/i);
  });

  it('allows TAB (0x09)', () => {
    // Tab is not in the rejected control range, even though Drive may not
    // do anything useful with it — we don't want to over-restrict.
    expect(buildDriveSearchQuery('a\tb')).toContain("'a\tb'");
  });

  it('ignores non-string name inputs (no contains clause)', () => {
    expect(buildDriveSearchQuery(42 as unknown)).toBe(BASE + TRASHED);
    expect(buildDriveSearchQuery(null as unknown)).toBe(BASE + TRASHED);
  });
});
