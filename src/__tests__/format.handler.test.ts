import { describe, it, expect, afterEach, vi } from 'vitest';
import { formatTools } from '../tools/format.js';
import { requestContext } from '../auth.js';

function jsonResponse(body: unknown) {
  return new Response(JSON.stringify(body), {
    status: 200,
    headers: { 'content-type': 'application/json' },
  });
}

function stubSheets() {
  const writes: any[] = [];
  vi.stubGlobal('fetch', async (input: any, init?: any) => {
    const url = typeof input === 'string' ? input : String(input?.url ?? input);
    if (init?.method === 'POST') {
      writes.push(JSON.parse(init.body));
      return jsonResponse({ replies: [{}] });
    }
    if (url.includes('sheets.properties')) {
      return jsonResponse({ sheets: [{ properties: { sheetId: 11, title: 'Data', index: 0 } }] });
    }
    throw new Error(`unexpected request ${url}`);
  });
  return writes;
}

const call = (tool: string, args: Record<string, unknown>) =>
  requestContext.run({ accessToken: 'test-token' }, () =>
    (formatTools as any)[tool].handler(args),
  );

const payload = (res: any) => res.structuredContent ?? JSON.parse(res.content[0].text);

afterEach(() => vi.unstubAllGlobals());

const grid = { sheetId: 11, startRowIndex: 0, endRowIndex: 3, startColumnIndex: 0, endColumnIndex: 2 };

describe('set_borders', () => {
  it('applies one style to every side by default', async () => {
    const writes = stubSheets();
    const out = payload(await call('set_borders', {
      spreadsheet_id: 's', sheet_name: 'Data', range: 'A1:B3', style: 'SOLID', color: '#FF0000',
    }));
    const border = { style: 'SOLID', color: { red: 1, green: 0, blue: 0 } };
    expect(writes[0].requests).toEqual([{
      updateBorders: {
        range: grid,
        top: border, bottom: border, left: border, right: border,
        innerHorizontal: border, innerVertical: border,
      },
    }]);
    expect(out.message).toMatch(/A1:B3/);
  });

  it('limits the change to the sides named, with outer and inner shorthands', async () => {
    const writes = stubSheets();
    await call('set_borders', {
      spreadsheet_id: 's', sheet_name: 'Data', range: 'A1:B3', style: 'DASHED', sides: ['outer'],
    });
    expect(Object.keys(writes[0].requests[0].updateBorders).sort()).toEqual(['bottom', 'left', 'range', 'right', 'top']);
    expect(writes[0].requests[0].updateBorders.top).toEqual({ style: 'DASHED' });

    await call('set_borders', {
      spreadsheet_id: 's', sheet_name: 'Data', range: 'A1:B3', style: 'DOTTED', sides: ['inner', 'top'],
    });
    expect(Object.keys(writes[1].requests[0].updateBorders).sort()).toEqual(['innerHorizontal', 'innerVertical', 'range', 'top']);
  });

  it('clears borders with style NONE', async () => {
    const writes = stubSheets();
    await call('set_borders', { spreadsheet_id: 's', sheet_name: 'Data', range: 'A1:B3', style: 'NONE' });
    expect(writes[0].requests[0].updateBorders.innerVertical).toEqual({ style: 'NONE' });
  });

  it('rejects an unknown side before sending anything', async () => {
    const writes = stubSheets();
    const res = await call('set_borders', {
      spreadsheet_id: 's', sheet_name: 'Data', range: 'A1:B3', style: 'SOLID', sides: ['middle'],
    });
    expect(res.isError).toBe(true);
    expect(writes).toHaveLength(0);
  });
});

describe('merge_cells', () => {
  it('merges the whole range by default', async () => {
    const writes = stubSheets();
    await call('merge_cells', { spreadsheet_id: 's', sheet_name: 'Data', range: 'A1:B3' });
    expect(writes[0].requests).toEqual([{ mergeCells: { range: grid, mergeType: 'MERGE_ALL' } }]);
  });

  it('refuses a single cell, which has nothing to merge', async () => {
    const writes = stubSheets();
    const res = await call('merge_cells', { spreadsheet_id: 's', sheet_name: 'Data', range: 'A1' });
    expect(res.isError).toBe(true);
    expect(writes).toHaveLength(0);
  });

  it('merges row by row when asked', async () => {
    const writes = stubSheets();
    await call('merge_cells', { spreadsheet_id: 's', sheet_name: 'Data', range: 'A1:B3', merge_type: 'MERGE_ROWS' });
    expect(writes[0].requests[0].mergeCells.mergeType).toBe('MERGE_ROWS');
  });

  it('unmerges every merge inside the range', async () => {
    const writes = stubSheets();
    await call('unmerge_cells', { spreadsheet_id: 's', sheet_name: 'Data', range: 'A1:B3' });
    expect(writes[0].requests).toEqual([{ unmergeCells: { range: grid } }]);
  });
});
