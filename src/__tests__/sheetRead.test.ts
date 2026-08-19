import { describe, it, expect } from 'vitest';
import { describeRead } from '../lib/sheetRead.js';
import { windowFor } from '../lib/window.js';
import type { DecodeResult } from '../lib/cells.js';

/**
 * The paging contract as pure arithmetic. read.test.ts drives the same rules
 * through a stubbed Sheets API; this pins them without one, so a change in
 * how a response describes itself fails here first.
 */
const decoded = (over: Partial<DecodeResult> = {}): DecodeResult => ({
  data: [], rowCount: 0, columnCount: 0, truncated: false, partialRow: false, ...over,
});

const TAB = { rowCount: 500_000, columnCount: 26 };

describe('describeRead', () => {
  it('says nothing when the window was fully satisfied', () => {
    const w = windowFor({ rowCount: 40, columnCount: 26 }, undefined);
    const out = describeRead(w, decoded({ rowCount: 40, columnCount: 26 }), false);

    expect(out).toEqual({ returnedRange: 'A1:Z40' });
    expect(out.truncated).toBeUndefined();
    expect(out.nextRange).toBeUndefined();
  });

  it('omits returnedRange when nothing came back', () => {
    const w = windowFor(TAB, undefined);
    expect(describeRead(w, decoded(), false).returnedRange).toBeUndefined();
  });

  it('points nextRange at the remaining scope, not the next window', () => {
    const w = windowFor(TAB, undefined);
    const out = describeRead(w, decoded({ rowCount: 192, columnCount: 26 }), true);

    expect(out.returnedRange).toBe('A1:Z192');
    expect(out.nextRange).toBe('A193:Z500000');
    expect(out.truncated).toBe(true);
  });

  it('resumes after the last row the decoder returned, not the window end', () => {
    // A budget tripped mid-window: paging must not skip the undecoded rows.
    const w = windowFor(TAB, undefined);
    const out = describeRead(w, decoded({ rowCount: 80, columnCount: 26, truncated: true }), false);

    expect(out.nextRange).toBe('A81:Z500000');
    expect(out.message).toContain('Output capped at');
  });

  it('keeps truncated and message in lockstep', () => {
    const w = windowFor({ rowCount: 10, columnCount: 400 }, undefined);
    const out = describeRead(w, decoded({ rowCount: 10, columnCount: 256 }), false);

    // Columns were dropped, so the flag is set and says why.
    expect(out.truncated).toBe(true);
    expect(out.message).toContain('column(s) beyond');
    // Nothing is flagged without an explanation, or explained without a flag.
    expect(Boolean(out.truncated)).toBe(Boolean(out.message));
  });

  it('reports a narrower returnedRange than the window when rows are trimmed', () => {
    const w = windowFor(TAB, 'A2900:B3002');
    const out = describeRead(w, decoded({ rowCount: 101, columnCount: 2 }), false);

    expect(out.returnedRange).toBe('A2900:B3000');
    expect(out.truncated).toBeUndefined();
  });
});
