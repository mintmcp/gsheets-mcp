import { describe, it, expect, afterEach, vi } from 'vitest';
import { z } from 'zod';
import { structureTools, MAX_BATCH_REQUESTS } from '../tools/structure.js';
import { requestContext } from '../auth.js';

function jsonResponse(body: unknown, status = 200) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { 'content-type': 'application/json' },
  });
}

interface Sent { url: string; method: string; body?: any }

function stubSheets(reply: unknown = { replies: [{}] }, status = 200) {
  const sent: Sent[] = [];
  vi.stubGlobal('fetch', async (input: any, init?: any) => {
    const url = decodeURIComponent(typeof input === 'string' ? input : String(input?.url ?? input));
    const method = init?.method ?? 'GET';
    if (method === 'GET') {
      return jsonResponse({
        sheets: [
          { properties: { sheetId: 11, title: 'Data' } },
          { properties: { sheetId: 22, title: 'Other' } },
        ],
      });
    }
    sent.push({ url, method, body: init?.body ? JSON.parse(init.body) : undefined });
    return jsonResponse(reply, status);
  });
  return sent;
}

const call = (tool: string, args: Record<string, unknown>) =>
  requestContext.run({ accessToken: 'test-token' }, () =>
    (structureTools as any)[tool].handler(args),
  );

const payload = (res: any) => res.structuredContent ?? JSON.parse(res.content[0].text);
const errorBody = (res: any) => JSON.parse(res.content[0].text);

afterEach(() => vi.unstubAllGlobals());

describe('insert_dimensions', () => {
  it('inserts rows at a 1-based row number, inheriting from the row above by default', async () => {
    const sent = stubSheets();
    const out = payload(await call('insert_dimensions', {
      spreadsheet_id: 's', sheet_name: 'Data', dimension: 'ROWS', start_index: 5, count: 2,
    }));
    expect(sent[0].body.requests).toEqual([{
      insertDimension: {
        range: { sheetId: 11, dimension: 'ROWS', startIndex: 4, endIndex: 6 },
        inheritFromBefore: true,
      },
    }]);
    expect(out.message).toMatch(/rows 5-6/);
  });

  it('inserts one column at a letter, inheriting from the right at column A', async () => {
    const sent = stubSheets();
    const out = payload(await call('insert_dimensions', {
      spreadsheet_id: 's', sheet_name: 'Data', dimension: 'COLUMNS', start_index: 'A',
    }));
    expect(sent[0].body.requests[0].insertDimension).toEqual({
      range: { sheetId: 11, dimension: 'COLUMNS', startIndex: 0, endIndex: 1 },
      inheritFromBefore: false,
    });
    expect(out.message).toMatch(/column A/);
  });

  it('honours an explicit inherit_from_before', async () => {
    const sent = stubSheets();
    await call('insert_dimensions', {
      spreadsheet_id: 's', sheet_name: 'Data', dimension: 'COLUMNS', start_index: 'C', inherit_from_before: false,
    });
    expect(sent[0].body.requests[0].insertDimension.inheritFromBefore).toBe(false);
  });

  it('refuses inheriting from before row 1, which Google rejects', async () => {
    const sent = stubSheets();
    const res = await call('insert_dimensions', {
      spreadsheet_id: 's', sheet_name: 'Data', dimension: 'ROWS', start_index: 1, inherit_from_before: true,
    });
    expect(res.isError).toBe(true);
    expect(errorBody(res).error).toMatch(/row 1/);
    expect(sent).toHaveLength(0);
  });

  it('rejects a letter for ROWS and a number for COLUMNS', async () => {
    stubSheets();
    const rows = await call('insert_dimensions', {
      spreadsheet_id: 's', sheet_name: 'Data', dimension: 'ROWS', start_index: 'C',
    });
    expect(errorBody(rows).error).toMatch(/row number/);
    const cols = await call('insert_dimensions', {
      spreadsheet_id: 's', sheet_name: 'Data', dimension: 'COLUMNS', start_index: 3,
    });
    expect(errorBody(cols).error).toMatch(/column letter/);
  });
});

describe('copy_sheet_to_spreadsheet', () => {
  it('names the destination when it is the .xlsx upload Google refuses', async () => {
    vi.stubGlobal('fetch', async (input: any, init?: any) => {
      const url = decodeURIComponent(String(input));
      if ((init?.method ?? 'GET') === 'POST') {
        return jsonResponse({ error: { code: 400, message: 'The document must not be an Office file.', status: 'FAILED_PRECONDITION' } }, 400);
      }
      if (url.includes('/drive/')) {
        expect(url).toContain('/files/dest');
        return jsonResponse({ name: 'Target.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', webViewLink: 'https://drive.google.com/file/d/dest' });
      }
      return jsonResponse({ sheets: [{ properties: { sheetId: 11, title: 'Data' } }] });
    });
    const res = await call('copy_sheet_to_spreadsheet', { spreadsheet_id: 'src', sheet_name: 'Data', destination_spreadsheet_id: 'dest' });
    expect(res.isError).toBe(true);
    expect(errorBody(res).error).toMatch(/Target\.xlsx.*convert_to_google_sheet/);
  });

  it('posts sheets.copyTo and reports the new tab', async () => {
    const sent = stubSheets({ sheetId: 99, title: 'Copy of Data', index: 1 });
    const out = payload(await call('copy_sheet_to_spreadsheet', {
      spreadsheet_id: 's', sheet_name: 'Data', destination_spreadsheet_id: 'dest',
    }));
    expect(sent[0].url).toContain('/s/sheets/11:copyTo');
    expect(sent[0].body).toEqual({ destinationSpreadsheetId: 'dest' });
    expect(out).toMatchObject({ id: 'dest', sheetId: 99, sheetTitle: 'Copy of Data', index: 1 });
  });

  it('names an unknown tab without sending anything', async () => {
    const sent = stubSheets();
    const res = await call('copy_sheet_to_spreadsheet', {
      spreadsheet_id: 's', sheet_name: 'Nope', destination_spreadsheet_id: 'dest',
    });
    expect(res.isError).toBe(true);
    expect(errorBody(res).error).toBe('Sheet tab "Nope" not found');
    expect(sent).toHaveLength(0);
  });
});

describe('batch_update', () => {
  it('forwards raw requests and returns the positional replies', async () => {
    const sent = stubSheets({ spreadsheetId: 's', replies: [{}, { addSheet: { properties: { sheetId: 5 } } }] });
    const requests = [
      { updateSpreadsheetProperties: { properties: { title: 'T' }, fields: 'title' } },
      { addSheet: { properties: { title: 'New' } } },
    ];
    const out = payload(await call('batch_update', { spreadsheet_id: 's', requests }));
    expect(sent[0].url).toContain('/s:batchUpdate');
    expect(sent[0].body).toEqual({ requests });
    expect(out).toEqual({
      id: 's',
      requestCount: 2,
      replies: [{}, { addSheet: { properties: { sheetId: 5 } } }],
      message: '2 request(s) applied',
    });
  });

  it('rejects a request with no key or two keys, naming its index', async () => {
    const sent = stubSheets();
    const two = await call('batch_update', {
      spreadsheet_id: 's', requests: [{ addSheet: {} }, { addSheet: {}, deleteSheet: {} }],
    });
    expect(two.isError).toBe(true);
    expect(errorBody(two).error).toMatch(/requests\[1\].*exactly one/);
    const none = await call('batch_update', { spreadsheet_id: 's', requests: [{}] });
    expect(errorBody(none).error).toMatch(/requests\[0\]/);
    expect(sent).toHaveLength(0);
  });

  it('rejects an empty array and a payload over the byte cap', async () => {
    const sent = stubSheets();
    const empty = await call('batch_update', { spreadsheet_id: 's', requests: [] });
    expect(empty.isError).toBe(true);
    const big = await call('batch_update', {
      spreadsheet_id: 's', requests: [{ pasteData: { data: 'x'.repeat(3_000_000) } }],
    });
    expect(big.isError).toBe(true);
    expect(errorBody(big).error).toMatch(/bytes/);
    expect(sent).toHaveLength(0);
  });

  it('caps the batch at MAX_BATCH_REQUESTS requests in the schema', () => {
    const schema = z.object(structureTools.batch_update.schema);
    const batch = (n: number) => ({ spreadsheet_id: 's', requests: Array.from({ length: n }, () => ({ addSheet: {} })) });
    expect(schema.safeParse(batch(MAX_BATCH_REQUESTS)).success).toBe(true);
    expect(schema.safeParse(batch(MAX_BATCH_REQUESTS + 1)).success).toBe(false);
  });

  it('names the failing request index and kind from a Google 400', async () => {
    stubSheets({ error: { code: 400, message: 'Invalid requests[1].addSheet: No request set.', status: 'INVALID_ARGUMENT' } }, 400);
    const res = await call('batch_update', {
      spreadsheet_id: 's', requests: [{ deleteSheet: { sheetId: 1 } }, { addSheet: { x: 1 } }],
    });
    expect(res.isError).toBe(true);
    expect(errorBody(res)).toMatchObject({
      status: 400,
      code: 'invalid_argument',
      failingRequestIndex: 1,
      failingRequestKind: 'addSheet',
    });
  });

  it('points an uploaded .xlsx to convert_to_google_sheet instead of blaming the request', async () => {
    vi.stubGlobal('fetch', async (input: any, init?: any) => {
      if ((init?.method ?? 'GET') === 'POST') {
        return jsonResponse({ error: { code: 400, message: 'This operation is not supported for this document. The document must not be an Office file.', status: 'FAILED_PRECONDITION' } }, 400);
      }
      return jsonResponse({
        name: 'Budget.xlsx',
        mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        webViewLink: 'https://drive.google.com/file/d/s',
      });
    });
    const res = await call('batch_update', { spreadsheet_id: 's', requests: [{ addSheet: {} }] });
    expect(res.isError).toBe(true);
    expect(errorBody(res).error).toMatch(/convert_to_google_sheet/);
  });

});
