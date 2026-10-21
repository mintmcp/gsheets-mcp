import { describe, it, expect, afterEach, vi } from 'vitest';
import { writeTools } from '../tools/write.js';
import { requestContext } from '../auth.js';

function jsonResponse(body: unknown) {
  return new Response(JSON.stringify(body), {
    status: 200,
    headers: { 'content-type': 'application/json' },
  });
}

interface Sent { url: string; method: string; body?: any }

function stubSheets() {
  const sent: Sent[] = [];
  vi.stubGlobal('fetch', async (input: any, init?: any) => {
    const url = decodeURIComponent(typeof input === 'string' ? input : String(input?.url ?? input));
    const method = init?.method ?? 'GET';
    if (method === 'GET') {
      return jsonResponse({ sheets: [{ properties: { sheetId: 11, title: 'Data' } }] });
    }
    sent.push({ url, method, body: init?.body ? JSON.parse(init.body) : undefined });
    return jsonResponse({
      updatedCells: 2,
      updates: { updatedRows: 1, updatedRange: "'Data'!A5:C5" },
      clearedRanges: ["'Data'!A1:B2", "'Other'!C1"],
      replies: [{}],
    });
  });
  return sent;
}

const call = (tool: string, args: Record<string, unknown>) =>
  requestContext.run({ accessToken: 'test-token' }, () =>
    (writeTools as any)[tool].handler(args),
  );

const error = (res: any) => JSON.parse(res.content[0].text).error as string;

afterEach(() => vi.unstubAllGlobals());

describe('update_range', () => {
  it('sends numbers and booleans as typed values, not strings', async () => {
    const sent = stubSheets();
    await call('update_range', {
      spreadsheet_id: 's', sheet_name: 'Data', range: 'A1:C1', data: [['North', 12, true]],
    });
    expect(sent[0].method).toBe('PUT');
    expect(sent[0].url).toContain("values/'Data'!A1:C1?valueInputOption=USER_ENTERED");
    expect(sent[0].body).toEqual({ values: [['North', 12, true]] });
  });

  it('pads ragged rows with empty strings and keeps null, which Google skips', async () => {
    const sent = stubSheets();
    await call('update_range', {
      spreadsheet_id: 's', sheet_name: 'Data', range: 'A1:C2', data: [[1, null], ['x', 'y', 'z']],
    });
    expect(sent[0].body).toEqual({ values: [[1, null, ''], ['x', 'y', 'z']] });
  });

  it('accepts a sheet-qualified range that matches sheet_name, and R1C1', async () => {
    const sent = stubSheets();
    await call('update_range', {
      spreadsheet_id: 's', sheet_name: 'Data', range: 'Data!R1C1:R2C2', data: [[1, 2], [3, 4]],
    });
    expect(sent[0].url).toContain("values/'Data'!A1:B2?");
  });

  it('rejects a range naming another tab than sheet_name', async () => {
    const sent = stubSheets();
    const res = await call('update_range', {
      spreadsheet_id: 's', sheet_name: 'Data', range: 'Other!A1', data: [['x']],
    });
    expect(res.isError).toBe(true);
    expect(error(res)).toMatch(/Other/);
    expect(sent).toHaveLength(0);
  });

  it('takes the tab from a qualified range when sheet_name is omitted', async () => {
    const sent = stubSheets();
    await call('update_range', { spreadsheet_id: 's', range: "'My Tab'!A1:B1", data: [[1, 2]] });
    expect(sent[0].url).toContain("values/'My Tab'!A1:B1?");
  });

  it('needs a tab from somewhere', async () => {
    const sent = stubSheets();
    const res = await call('update_range', { spreadsheet_id: 's', range: 'A1', data: [['x']] });
    expect(res.isError).toBe(true);
    expect(error(res)).toMatch(/sheet_name/);
    expect(sent).toHaveLength(0);
  });

  it('rejects an object cell before sending anything', async () => {
    const sent = stubSheets();
    const res = await call('update_range', {
      spreadsheet_id: 's', sheet_name: 'Data', range: 'A1', data: [[{ v: 1 }]],
    });
    expect(res.isError).toBe(true);
    expect(sent).toHaveLength(0);
  });
});

describe('update_cell', () => {
  it('takes the tab from a qualified cell when sheet_name is omitted', async () => {
    const sent = stubSheets();
    await call('update_cell', { spreadsheet_id: 's', cell: 'Data!B3', content: [{ text: 'hi' }] });
    expect(sent[0].url).toContain("values/'Data'!B3?valueInputOption=USER_ENTERED");
    expect(sent[0].body).toEqual({ values: [['hi']] });
  });

  it('translates an R1C1 cell for the hyperlink path', async () => {
    const sent = stubSheets();
    await call('update_cell', {
      spreadsheet_id: 's', sheet_name: 'Data', cell: 'R3C2',
      content: [{ text: 'Go', url: 'https://example.com' }],
    });
    expect(sent[0].body.requests[0].updateCells.range).toEqual({
      sheetId: 11, startRowIndex: 2, endRowIndex: 3, startColumnIndex: 1, endColumnIndex: 2,
    });
  });

  it('links only its own segment, with offsets in UTF-16 units as Sheets counts them', async () => {
    const sent = stubSheets();
    await call('update_cell', {
      spreadsheet_id: 's', sheet_name: 'Data', cell: 'A1',
      content: [{ text: '🎉 see ' }, { text: 'docs', url: 'https://example.com' }, { text: ' now' }],
    });
    const cell = sent[0].body.requests[0].updateCells.rows[0].values[0];
    expect(cell.userEnteredValue).toEqual({ stringValue: '🎉 see docs now' });
    expect(cell.textFormatRuns).toEqual([
      { startIndex: 0, format: {} },
      { startIndex: 7, format: { link: { uri: 'https://example.com' } } },
      { startIndex: 11, format: {} },
    ]);
  });

  it('still refuses a range', async () => {
    const sent = stubSheets();
    const res = await call('update_cell', { spreadsheet_id: 's', cell: 'Data!A1:B2', content: [{ text: 'x' }] });
    expect(res.isError).toBe(true);
    expect(error(res)).toMatch(/update_range/);
    expect(sent).toHaveLength(0);
  });

  it('rejects a cell naming another tab than sheet_name', async () => {
    const sent = stubSheets();
    const res = await call('update_cell', {
      spreadsheet_id: 's', sheet_name: 'Data', cell: 'Other!A1', content: [{ text: 'x' }],
    });
    expect(res.isError).toBe(true);
    expect(sent).toHaveLength(0);
  });
});

describe('insert_rows', () => {
  it('appends typed values after the tab by default', async () => {
    const sent = stubSheets();
    const res = await call('insert_rows', {
      spreadsheet_id: 's', sheet_name: 'Data', data: [['West', 9, false]],
    });
    expect(sent[0].url).toContain("values/'Data':append?");
    expect(sent[0].url).toContain('valueInputOption=USER_ENTERED');
    expect(sent[0].url).toContain('insertDataOption=INSERT_ROWS');
    expect(sent[0].body).toEqual({ values: [['West', 9, false]] });
    expect(res.structuredContent).toMatchObject({ updatedRows: 1, updatedRange: "'Data'!A5:C5" });
  });

  it('appends after the logical table that a range anchors, as values.append does', async () => {
    const sent = stubSheets();
    await call('insert_rows', {
      spreadsheet_id: 's', sheet_name: 'Data', range: 'F1', data: [['a']],
    });
    expect(sent[0].url).toContain("values/'Data'!F1:append?");
  });

  it('accepts whole columns, the usual values.append anchor', async () => {
    const sent = stubSheets();
    await call('insert_rows', { spreadsheet_id: 's', range: 'Data!A:D', data: [['a']] });
    expect(sent[0].url).toContain("values/'Data'!A:D:append?");
  });

  it('takes the tab from a qualified range when sheet_name is omitted', async () => {
    const sent = stubSheets();
    await call('insert_rows', { spreadsheet_id: 's', range: 'Data!A1:D1', data: [['a']] });
    expect(sent[0].url).toContain("values/'Data'!A1:D1:append?");
  });

  it('needs a tab from somewhere', async () => {
    const sent = stubSheets();
    const res = await call('insert_rows', { spreadsheet_id: 's', data: [['a']] });
    expect(res.isError).toBe(true);
    expect(error(res)).toMatch(/sheet_name/);
    expect(sent).toHaveLength(0);
  });
});

describe('clear_values', () => {
  it('qualifies bare ranges with sheet_name', async () => {
    const sent = stubSheets();
    await call('clear_values', { spreadsheet_id: 's', sheet_name: 'Data', ranges: ['A1:B2', 'R1C3'] });
    expect(sent[0].url).toContain('values:batchClear');
    expect(sent[0].body).toEqual({ ranges: ["'Data'!A1:B2", "'Data'!C1"] });
  });

  it('clears across tabs when every range carries its own tab', async () => {
    const sent = stubSheets();
    const res = await call('clear_values', {
      spreadsheet_id: 's', ranges: ['Data!A1:B2', "'Other'!C1", 'Data!1:3', 'Other!A:C'],
    });
    expect(sent[0].body).toEqual({
      ranges: ["'Data'!A1:B2", "'Other'!C1", "'Data'!1:3", "'Other'!A:C"],
    });
    expect(res.structuredContent.clearedRanges).toEqual(["'Data'!A1:B2", "'Other'!C1"]);
  });

  it('mixes bare ranges on sheet_name with ranges on other tabs', async () => {
    const sent = stubSheets();
    await call('clear_values', { spreadsheet_id: 's', sheet_name: 'Data', ranges: ['A:A', 'Other!C:C'] });
    expect(sent[0].body).toEqual({ ranges: ["'Data'!A:A", "'Other'!C:C"] });
  });

  it('refuses a bare range when no sheet_name was given', async () => {
    const sent = stubSheets();
    const res = await call('clear_values', { spreadsheet_id: 's', ranges: ['Data!A1', 'B2'] });
    expect(res.isError).toBe(true);
    expect(error(res)).toMatch(/ranges\[1\]/);
    expect(sent).toHaveLength(0);
  });
});
