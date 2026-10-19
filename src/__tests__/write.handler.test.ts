import { describe, it, expect, afterEach, vi } from 'vitest';
import { writeTools } from '../tools/write.js';
import { requestContext } from '../auth.js';

/**
 * Drives the value-writing handlers against a stubbed Sheets API and asserts
 * the query string and body each one sends.
 */

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
      return jsonResponse({ sheets: [{ properties: { sheetId: 11, title: 'Data', index: 0 } }] });
    }
    sent.push({ url, method, body: init?.body ? JSON.parse(init.body) : undefined });
    return jsonResponse({ updatedCells: 2, updates: { updatedRows: 1 }, replies: [{}] });
  });
  return sent;
}

const call = (tool: string, args: Record<string, unknown>) =>
  requestContext.run({ accessToken: 'test-token' }, () =>
    (writeTools as any)[tool].handler(args),
  );

afterEach(() => vi.unstubAllGlobals());

describe('value_input_option', () => {
  it('update_range defaults to USER_ENTERED and can send RAW', async () => {
    const sent = stubSheets();
    await call('update_range', { spreadsheet_id: 's', sheet_name: 'Data', range: 'A1:B1', data: [['=1+1', '01']] });
    expect(sent[0].url).toContain('valueInputOption=USER_ENTERED');

    await call('update_range', {
      spreadsheet_id: 's', sheet_name: 'Data', range: 'A1:B1', data: [['=1+1', '01']], value_input_option: 'RAW',
    });
    expect(sent[1].url).toContain('valueInputOption=RAW');
    expect(sent[1].body).toEqual({ values: [['=1+1', '01']] });
  });

  it('insert_rows passes the option through to values:append', async () => {
    const sent = stubSheets();
    await call('insert_rows', { spreadsheet_id: 's', sheet_name: 'Data', data: [['01']], value_input_option: 'RAW' });
    expect(sent[0].url).toContain("values/'Data':append");
    expect(sent[0].url).toContain('valueInputOption=RAW');
    expect(sent[0].url).toContain('insertDataOption=INSERT_ROWS');
  });

  it('update_cell passes the option through for a plain value', async () => {
    const sent = stubSheets();
    await call('update_cell', {
      spreadsheet_id: 's', sheet_name: 'Data', cell: 'B2', content: [{ text: '01' }], value_input_option: 'RAW',
    });
    expect(sent[0].method).toBe('PUT');
    expect(sent[0].url).toContain('valueInputOption=RAW');
  });

  it('update_cell with hyperlinks writes a literal string, so the option is moot', async () => {
    const sent = stubSheets();
    await call('update_cell', {
      spreadsheet_id: 's', sheet_name: 'Data', cell: 'B2',
      content: [{ text: 'see ' }, { text: 'docs', url: 'https://x' }],
    });
    expect(sent[0].method).toBe('POST');
    expect(sent[0].body.requests[0].updateCells.rows[0].values[0].userEnteredValue).toEqual({ stringValue: 'see docs' });
  });

  it('rejects an unknown option before sending anything', async () => {
    const sent = stubSheets();
    const res = await call('update_range', {
      spreadsheet_id: 's', sheet_name: 'Data', range: 'A1', data: [['x']], value_input_option: 'MAGIC',
    });
    expect(res.isError).toBe(true);
    expect(sent).toHaveLength(0);
  });
});
