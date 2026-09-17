import { describe, it, expect, beforeEach, vi } from 'vitest';
import { readTools } from '../tools/read.js';
import { writeTools } from '../tools/write.js';
import { formatTools } from '../tools/format.js';
import { chartTools } from '../tools/charts.js';
import { tools } from '../tools/index.js';
import { requestContext } from '../auth.js';
import { MAX_CELL_CHARS } from '../lib/sheetBudget.js';

/**
 * Drives get_sheet_data against a stubbed Sheets API so the windowing and
 * paging contract is exercised without network access.
 */

interface Stub {
  rowCount: number;
  columnCount: number;
  /** How many rows of the tab actually hold data. */
  rowsReturned: number;
  title?: string;
  /** Cell text, so a test can make the character budget bind before the cell cap. */
  cellText?: string;
  /** Exact rows holding data, for sparse tabs. Overrides rowsReturned. */
  dataRows?: number[];
}

let calls: string[] = [];

/** Mirrors the API: never returns more rows than the requested range covers. */
function stubSheets({ rowCount, columnCount, rowsReturned, title = 'Sheet1', cellText = 'x', dataRows }: Stub) {
  vi.stubGlobal('fetch', async (input: any) => {
    const url = typeof input === 'string' ? input : String(input?.url ?? input);
    calls.push(url);

    if (url.includes('gridProperties')) {
      return jsonResponse({
        sheets: [{ properties: { title, gridProperties: { rowCount, columnCount } } }],
      });
    }

    const asked = decodeURIComponent(url).match(/!([A-Z]+)(\d+):([A-Z]+)(\d+)/);
    const startRow = asked ? Number(asked[2]) : 1;
    const endRow = asked ? Number(asked[4]) : rowCount;
    const askedCols = asked ? letterToIndex(asked[3]) - letterToIndex(asked[1]) + 1 : columnCount;

    const cells = (n: number) => Array.from({ length: n }, () => ({
      effectiveValue: { stringValue: cellText },
      formattedValue: cellText,
    }));

    if (dataRows) {
      // Sheets drops trailing blank rows but keeps blanks between populated ones
      const inRange = dataRows.filter((r) => r >= startRow && r <= endRow);
      const last = inRange.length ? Math.max(...inRange) : startRow - 1;
      const sparse = Array.from({ length: Math.max(0, last - startRow + 1) }, (_, i) =>
        dataRows.includes(startRow + i)
          ? { values: cells(Math.min(askedCols, columnCount)) }
          : {});
      return jsonResponse({ sheets: [{ data: [{ rowData: sparse }] }] });
    }

    const available = Math.max(0, Math.min(rowsReturned, endRow) - startRow + 1);
    const rowData = Array.from({ length: available }, () => ({
      values: Array.from({ length: Math.min(askedCols, columnCount) }, () => ({
        effectiveValue: { stringValue: cellText },
        formattedValue: cellText,
      })),
    }));
    return jsonResponse({ sheets: [{ data: [{ rowData }] }] });
  });
}

function letterToIndex(letter: string): number {
  let index = 0;
  for (const ch of letter) index = index * 26 + (ch.charCodeAt(0) - 64);
  return index - 1;
}

function jsonResponse(body: unknown) {
  return new Response(JSON.stringify(body), {
    status: 200,
    headers: { 'content-type': 'application/json' },
  });
}

const run = (args: Record<string, unknown>) =>
  requestContext.run({ accessToken: 'test-token' }, () =>
    (readTools.get_sheet_data.handler as any)(args),
  );

const payload = (res: any) => res.structuredContent ?? JSON.parse(res.content[0].text);

beforeEach(() => {
  calls = [];
});

describe('get_sheet_data windowing', () => {
  it('requests a bounded A1 window instead of the whole tab', async () => {
    stubSheets({ rowCount: 500_000, columnCount: 26, rowsReturned: 500_000 });
    const out = payload(await run({ spreadsheet_id: 'abc' }));

    // 5,000 cells over 26 columns is a 192-row window; the request adds one
    // probe row to detect whether more data exists beyond it.
    const gridCall = calls.find((c) => c.includes('ranges='))!;
    expect(decodeURIComponent(gridCall)).toContain("'Sheet1'!A1:Z193");
    expect(out.returnedRange).toBe('A1:Z192');
  });

  it('asks Google for the computed value of every cell', async () => {
    // Formula results and spilled cells exist only in effectiveValue
    stubSheets({ rowCount: 10, columnCount: 2, rowsReturned: 10 });
    await run({ spreadsheet_id: 'abc' });
    const gridCall = calls.find((c) => c.includes('ranges='))!;
    const fields = decodeURIComponent(gridCall);
    expect(fields).toContain('effectiveValue');
    expect(fields).toContain('userEnteredValue.formulaValue');
  });

  it('reports truncated and a usable nextRange on a large tab', async () => {
    stubSheets({ rowCount: 500_000, columnCount: 26, rowsReturned: 500_000 });
    const out = payload(await run({ spreadsheet_id: 'abc' }));

    expect(out.rowCount).toBe(192);
    expect(out.truncated).toBe(true);
    // The remaining scope, re-clamped to a window when passed back.
    expect(out.nextRange).toBe('A193:Z500000');
    expect(out.message).toContain('pass nextRange');
  });

  it('paging with nextRange continues from the following row', async () => {
    stubSheets({ rowCount: 500_000, columnCount: 26, rowsReturned: 500_000 });
    const out = payload(await run({ spreadsheet_id: 'abc', range: 'A193:Z500000' }));

    const gridCall = decodeURIComponent(calls.find((c) => c.includes('ranges='))!);
    expect(gridCall).toContain("'Sheet1'!A193:Z");
    expect(out.returnedRange).toBe('A193:Z384');
    expect(out.rowCount).toBe(192);
  });

  it('truncates on the character budget before the cell cap on fat content', async () => {
    // Short values fill the whole 192-row window; 200-character cells trip
    // the character budget long before it, so the page shrinks with density.
    stubSheets({
      rowCount: 100_000, columnCount: 26, rowsReturned: 100_000,
      cellText: 'y'.repeat(200),
    });
    const out = payload(await run({ spreadsheet_id: 'abc' }));

    expect(out.rowCount).toBeLessThan(192);
    expect(out.rowCount).toBeGreaterThan(0);
    expect(out.truncated).toBe(true);
    expect(out.message).toContain('Output capped at');
    expect(out.nextRange).toBe(`A${out.rowCount + 1}:Z100000`);
  });

  it('blames characters when prose stops the page well short of the cell cap', async () => {
    stubSheets({
      rowCount: 100_000, columnCount: 8, rowsReturned: 100_000,
      cellText: 'y'.repeat(2_000),
    });
    const out = payload(await run({ spreadsheet_id: 'abc' }));

    expect(out.truncated).toBe(true);
    expect(out.message).toContain('250000 characters');
    expect(out.message).not.toContain('cells');
  });

  it('never blames the cell cap, which bounds the fetch window not the page', async () => {
    // floor(MAX_CELLS / columns) rows means a window holds at most MAX_CELLS
    // cells, so admitCell's cell ceiling is unreachable on this path. Short
    // values end a page by exhausting the window, which nextRange reports.
    stubSheets({
      rowCount: 100_000, columnCount: 25, rowsReturned: 100_000, cellText: 'v',
    });
    const out = payload(await run({ spreadsheet_id: 'abc' }));

    expect(out.message).not.toContain('Output capped');
    expect(out.nextRange).toBeDefined();
  });

  it('returns a whole small tab without invoking either budget', async () => {
    stubSheets({ rowCount: 40, columnCount: 26, rowsReturned: 40 });
    const out = payload(await run({ spreadsheet_id: 'abc' }));

    expect(out.rowCount).toBe(40);
    expect(out.truncated).toBeUndefined();
    expect(out.message).toBeUndefined();
  });

  it('finds data past a blank run longer than one window', async () => {
    stubSheets({ rowCount: 1_000, columnCount: 26, rowsReturned: 0, dataRows: [1, 300] });

    const first = payload(await run({ spreadsheet_id: 'abc' }));
    expect(first.truncated).toBe(true);
    expect(first.nextRange).toBeDefined();

    const second = payload(await run({ spreadsheet_id: 'abc', range: first.nextRange }));
    const lastRow = Number(second.returnedRange.match(/(\d+)$/)![1]);
    expect(lastRow).toBe(300);
  });

  it('emits no nextRange when a partial row ends the scope', async () => {
    // A row whose cells alone blow the character budget comes back partial and
    // still signals more pages. On a single-row scope that would make a reversed
    // range, which assertBareA1Range rejects when passed back.
    stubSheets({
      rowCount: 1_000, columnCount: 8, rowsReturned: 1_000,
      cellText: 'x'.repeat(MAX_CELL_CHARS),
    });
    const out = payload(await run({ spreadsheet_id: 'abc', range: 'A1:H1' }));

    expect(out.rowCount).toBe(1);
    expect(out.data[0].length).toBeLessThan(8);
    expect(out.truncated).toBe(true);
    expect(out.message).toContain('final row is incomplete');
    expect(out.nextRange).toBeUndefined();
  });

  it('still pages when a partial row leaves rows behind it', async () => {
    stubSheets({
      rowCount: 1_000, columnCount: 5, rowsReturned: 1_000,
      cellText: 'x'.repeat(MAX_CELL_CHARS),
    });
    const out = payload(await run({ spreadsheet_id: 'abc', range: 'A1:E50' }));

    expect(out.nextRange).toBe('A2:E50');
  });

  it('omits returnedRange when the window came back empty', async () => {
    // An explicit range past the end of the data returns no cells at all;
    // reporting "A5000:A5000" would claim one cell was returned.
    stubSheets({ rowCount: 100_000, columnCount: 26, rowsReturned: 50 });
    const out = payload(await run({ spreadsheet_id: 'abc', range: 'A5000:Z5010' }));

    expect(out.rowCount).toBe(0);
    expect(out.returnedRange).toBeUndefined();
    expect(out.nextRange).toBeUndefined();
  });

  it('does not report columns the tab never had as omitted', async () => {
    // A generous rectangle over a narrow tab is fully satisfied, so nothing
    // was omitted and nothing should be flagged as truncated.
    stubSheets({ rowCount: 100, columnCount: 26, rowsReturned: 100 });
    const out = payload(await run({ spreadsheet_id: 'abc', range: 'A1:ZZ10' }));

    expect(out.truncated).toBeUndefined();
    expect(out.message).toBeUndefined();
    expect(decodeURIComponent(calls.find((c) => c.includes('ranges='))!)).toContain("'Sheet1'!A1:Z10");
  });

  it('leaves truncated and nextRange absent when the tab fits', async () => {
    stubSheets({ rowCount: 100, columnCount: 10, rowsReturned: 100 });
    const out = payload(await run({ spreadsheet_id: 'abc' }));

    expect(out.truncated).toBeUndefined();
    expect(out.nextRange).toBeUndefined();
    expect(out.rowCount).toBe(100);
  });

  it('does not flag truncation when the allocated grid exceeds the used range', async () => {
    // 50k allocated rows but only 20 rows of real data: the window asks for
    // more than exists, and fewer rows come back, so nothing was clipped.
    stubSheets({ rowCount: 50_000, columnCount: 26, rowsReturned: 20 });
    const out = payload(await run({ spreadsheet_id: 'abc' }));

    expect(out.truncated).toBeUndefined();
    expect(out.nextRange).toBeUndefined();
    expect(out.rowCount).toBe(20);
  });

  it('clamps the upstream request to 256 columns on a very wide tab', async () => {
    stubSheets({ rowCount: 10, columnCount: 400, rowsReturned: 10 });
    const out = payload(await run({ spreadsheet_id: 'abc' }));

    // IV is the 256th column: prove the clamp reached the wire, not just the message.
    const gridCall = decodeURIComponent(calls.find((c) => c.includes('ranges='))!);
    expect(gridCall).toContain('!A1:IV');
    expect(gridCall).not.toMatch(/!A1:O[A-Z]/);
    expect(out.truncated).toBe(true);
    expect(out.message).toContain('column(s) past');
  });

  it('clamps an oversized explicit range instead of forwarding it upstream', async () => {
    stubSheets({ rowCount: 500_000, columnCount: 400, rowsReturned: 1923 });
    const out = payload(await run({ spreadsheet_id: 'abc', range: 'A1:OJ100000' }));

    const gridCall = decodeURIComponent(calls.find((c) => c.includes('ranges='))!);
    expect(gridCall).not.toContain('OJ100000');
    expect(gridCall).toContain('!A1:IV');
    expect(out.truncated).toBe(true);
    expect(out.nextRange).toBeDefined();
  });

  it('offers a nextRange when the char budget trips before the row count fills', async () => {
    // One 20k-char cell per row blows MAX_OUTPUT_CHARS (4M) long before the
    // 50k-cell cap. Kept well under the 25MB upstream read budget so this
    // exercises the decoder, not the byte guard.
    const wide = 'w'.repeat(20_000);
    vi.stubGlobal('fetch', async (input: any) => {
      const url = typeof input === 'string' ? input : String(input?.url ?? input);
      calls.push(url);
      if (url.includes('gridProperties')) {
        return jsonResponse({ sheets: [{ properties: { title: 'Sheet1', gridProperties: { rowCount: 500_000, columnCount: 1 } } }] });
      }
      const rowData = Array.from({ length: 300 }, () => ({
        values: [{ effectiveValue: { stringValue: wide }, formattedValue: wide }],
      }));
      return jsonResponse({ sheets: [{ data: [{ rowData }] }] });
    });

    const out = payload(await run({ spreadsheet_id: 'abc' }));

    expect(out.truncated).toBe(true);
    expect(out.nextRange).toBeDefined();
    expect(out.nextRange).toBe(`A${out.rowCount + 1}:A500000`);
  });

  it('does not flag truncation when the used range ends exactly on the window boundary', async () => {
    // Allocated grid is taller than the window, but the probe row comes back
    // empty, so there is genuinely nothing more to fetch.
    stubSheets({ rowCount: 500_000, columnCount: 26, rowsReturned: 102 });
    const out = payload(await run({ spreadsheet_id: 'abc' }));
    expect(out.rowCount).toBe(102);
    expect(out.truncated).toBeUndefined();
    expect(out.nextRange).toBeUndefined();
  });

  it('walks an oversized explicit range to completion without stranding rows', async () => {
    stubSheets({ rowCount: 500_000, columnCount: 26, rowsReturned: 500_000 });

    const seen: Array<[number, number]> = [];
    let cursor: string | undefined = 'A1:Z1000';
    for (let page = 0; page < 20 && cursor; page++) {
      const out = payload(await run({ spreadsheet_id: 'abc', range: cursor }));
      const [, first, , last] = out.returnedRange.match(/^A(\d+):([A-Z]+)(\d+)$/)!.map(Number);
      seen.push([first, last]);
      cursor = out.nextRange;
    }

    // Contiguous, no gaps, no overlap, terminating exactly at the scope end.
    expect(seen[0][0]).toBe(1);
    for (let i = 1; i < seen.length; i++) {
      expect(seen[i][0]).toBe(seen[i - 1][1] + 1);
    }
    expect(seen[seen.length - 1][1]).toBe(1000);
    expect(cursor).toBeUndefined();
  });

  it('walks a whole tab to completion from a default read', async () => {
    stubSheets({ rowCount: 1_000, columnCount: 26, rowsReturned: 1_000 });

    const seen: Array<[number, number]> = [];
    let cursor: string | undefined;
    for (let page = 0; page < 20; page++) {
      const args: Record<string, unknown> = { spreadsheet_id: 'abc' };
      if (cursor) args.range = cursor;
      const out = payload(await run(args));
      const [, first, , last] = out.returnedRange.match(/^A(\d+):([A-Z]+)(\d+)$/)!.map(Number);
      seen.push([first, last]);
      cursor = out.nextRange;
      if (!cursor) break;
    }

    expect(seen[0][0]).toBe(1);
    for (let i = 1; i < seen.length; i++) {
      expect(seen[i][0]).toBe(seen[i - 1][1] + 1);
    }
    expect(seen[seen.length - 1][1]).toBe(1_000);
    expect(cursor).toBeUndefined();
  });

  it('does not report omitted columns for a narrow explicit range', async () => {
    // The tab is 26 columns wide; asking for one column omits nothing.
    stubSheets({ rowCount: 1000, columnCount: 26, rowsReturned: 20 });
    const out = payload(await run({ spreadsheet_id: 'abc', range: 'E1:E2' }));

    expect(out.truncated).toBeUndefined();
    expect(out.message).toBeUndefined();
    expect(out.returnedRange).toMatch(/^E1:/);
  });

  it('still reports omitted columns when the window itself clamped them', async () => {
    stubSheets({ rowCount: 10, columnCount: 400, rowsReturned: 10 });
    const out = payload(await run({ spreadsheet_id: 'abc' }));

    expect(out.truncated).toBe(true);
    expect(out.message).toContain('column(s) past');
  });

  it('returns a clean error for an unknown sheet name', async () => {
    stubSheets({ rowCount: 10, columnCount: 5, rowsReturned: 10 });
    const res: any = await run({ spreadsheet_id: 'abc', sheet_name: 'Nope' });

    expect(res.isError).toBe(true);
    expect(JSON.parse(res.content[0].text).error).toBe('Sheet tab "Nope" not found');
  });

  it('rejects an unbounded range argument', async () => {
    stubSheets({ rowCount: 10, columnCount: 5, rowsReturned: 10 });
    const res: any = await run({ spreadsheet_id: 'abc', range: 'A:C' });

    expect(res.isError).toBe(true);
    expect(JSON.parse(res.content[0].text).error).toMatch(/not a bounded A1 range/);
  });
});

describe('get_metadata structure', () => {
  function stubMetadata(tabs: Array<{ title: string; rows?: number; cols?: number }>) {
    vi.stubGlobal('fetch', async (input: any) => {
      const url = decodeURIComponent(typeof input === 'string' ? input : String(input?.url ?? input));
      calls.push(url);
      if (url.includes('values:batchGet')) {
        const requested = (url.match(/ranges=/g) || []).length;
        return jsonResponse({
          valueRanges: Array.from({ length: requested }, (_, i) => ({
            values: [[`h1_${i}`, `h2_${i}`]],
          })),
        });
      }
      return jsonResponse({
        properties: { title: 'Book' },
        spreadsheetUrl: 'https://x',
        sheets: tabs.map((t, i) => ({
          properties: {
            title: t.title,
            index: i,
            // Object sheets (charts) carry no gridProperties at all.
            ...(t.cols === undefined
              ? {}
              : { gridProperties: { rowCount: t.rows, columnCount: t.cols } }),
          },
        })),
      });
    });
  }

  const runMeta = (args: Record<string, unknown>) =>
    requestContext.run({ accessToken: 't' }, () =>
      (readTools.get_metadata.handler as any)(args),
    );

  it('returns per-tab dimensions without an extra API call', async () => {
    stubMetadata([{ title: 'Sales', rows: 3000, cols: 26 }]);
    const out = payload(await runMeta({ spreadsheet_id: 'abc' }));

    expect(out.sheets[0]).toMatchObject({ title: 'Sales', rowCount: 3000, columnCount: 26 });
    expect(out.sheets[0].headers).toBeUndefined();
    expect(calls.filter((c) => c.includes('values:batchGet'))).toHaveLength(0);
  });

  it('fetches headers in one batched call when asked', async () => {
    stubMetadata([
      { title: 'Sales', rows: 3000, cols: 26 },
      { title: 'Costs', rows: 10, cols: 3 },
    ]);
    const out = payload(await runMeta({ spreadsheet_id: 'abc', include_headers: true }));

    expect(out.sheets[0].headers).toEqual(['h1_0', 'h2_0']);
    expect(out.sheets[1].headers).toEqual(['h1_1', 'h2_1']);
    expect(calls.filter((c) => c.includes('values:batchGet'))).toHaveLength(1);
  });

  it('clamps the header range to the response column limit', async () => {
    stubMetadata([{ title: 'Wide', rows: 10, cols: 400 }]);
    await runMeta({ spreadsheet_id: 'abc', include_headers: true });

    const batch = calls.find((c) => c.includes('values:batchGet'))!;
    expect(batch).toContain("'Wide'!A1:IV1");
  });

  it('skips object sheets when batching header ranges', async () => {
    // A chart tab has no cells. Asking values:batchGet for one poisons the
    // whole batch, which would cost every other tab its headers.
    stubMetadata([
      { title: 'Data', rows: 10, cols: 3 },
      { title: 'Q3 Chart' },
      { title: 'More', rows: 10, cols: 2 },
    ]);
    const out = payload(await runMeta({ spreadsheet_id: 'abc', include_headers: true }));

    const batch = calls.find((c) => c.includes('values:batchGet'))!;
    expect(batch).not.toContain('Q3 Chart');
    expect((batch.match(/ranges=/g) || []).length).toBe(2);

    // The chart tab still appears, just without headers or dimensions.
    expect(out.sheets.map((s: any) => s.title)).toEqual(['Data', 'Q3 Chart', 'More']);
    expect(out.sheets[1].headers).toBeUndefined();
    expect(out.sheets[1].rowCount).toBeUndefined();
    // Headers land on the tabs they were requested for, not by position.
    expect(out.sheets[0].headers).toEqual(['h1_0', 'h2_0']);
    expect(out.sheets[2].headers).toEqual(['h1_1', 'h2_1']);
  });

  it('makes no header call when every tab is an object sheet', async () => {
    stubMetadata([{ title: 'Chart A' }, { title: 'Chart B' }]);
    const out = payload(await runMeta({ spreadsheet_id: 'abc', include_headers: true }));

    expect(calls.filter((c) => c.includes('values:batchGet'))).toHaveLength(0);
    expect(out.sheets).toHaveLength(2);
  });

  it('caps a spreadsheet with an absurd number of tabs', async () => {
    const tabs = Array.from({ length: 1_050 }, (_, i) => ({ title: `T${i}`, rows: 10, cols: 3 }));
    stubMetadata(tabs);
    const out = payload(await runMeta({ spreadsheet_id: 'abc' }));

    expect(out.sheets).toHaveLength(1_000);
    expect(out.truncated).toBe(true);
    expect(out.message).toContain('not listed');
  });
});

describe('tool namespace', () => {
  it('exposes every tool from all four modules with no name collisions', () => {
    // The modules are merged by spread, so a name defined twice would silently
    // overwrite and one tool would vanish from the server. TypeScript does not
    // catch it: an intersection with a duplicate key is a valid type.
    const names = [
      ...Object.keys(readTools),
      ...Object.keys(writeTools),
      ...Object.keys(formatTools),
      ...Object.keys(chartTools),
    ];
    expect(new Set(names).size).toBe(names.length);
    expect(Object.keys(tools)).toHaveLength(names.length);
  });
});
