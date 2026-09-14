import { describe, it, expect, afterEach, vi } from 'vitest';
import { chartTools } from '../tools/charts.js';
import { requestContext } from '../auth.js';

/**
 * Drives the chart handlers against a stubbed Sheets API. The pure-function
 * tests in chart.test.ts cannot reach this layer, which is where a chart sheet
 * passed as a tab name used to reach Google and come back as "No grid with id".
 */

function jsonResponse(body: unknown) {
  return new Response(JSON.stringify(body), {
    status: 200,
    headers: { 'content-type': 'application/json' },
  });
}

/** One grid tab and one chart sheet, the shape every chart mask reads. */
function stubSheets() {
  const writes: any[] = [];
  vi.stubGlobal('fetch', async (input: any, init?: any) => {
    const url = typeof input === 'string' ? input : String(input?.url ?? input);
    if (init?.method === 'POST') {
      writes.push(JSON.parse(init.body));
      return jsonResponse({ replies: [{ addChart: { chart: { chartId: 7 } } }] });
    }
    if (url.includes('values:batchGet')) {
      return jsonResponse({ valueRanges: [{ values: [['Model']] }, { values: [['Q1']] }] });
    }
    return jsonResponse({
      sheets: [
        { properties: { sheetId: 1, title: 'Data', sheetType: 'GRID' }, charts: [] },
        { properties: { sheetId: 2, title: 'Chart1', sheetType: 'OBJECT' }, charts: [] },
      ],
    });
  });
  return writes;
}

const call = (tool: string, args: Record<string, unknown>) =>
  requestContext.run({ accessToken: 'test-token' }, () =>
    (chartTools as any)[tool].handler(args),
  );

afterEach(() => vi.unstubAllGlobals());

describe('a chart sheet is refused before the request reaches Google', () => {
  const base = {
    spreadsheet_id: 's', chart_type: 'COLUMN',
    domain_range: 'A1:A9', series_ranges: ['B1:B9'],
  };

  it('refuses it as the anchor tab, and sends nothing', async () => {
    const writes = stubSheets();
    const result = await call('add_chart', {
      ...base, sheet_name: 'Data', anchor_cell: 'A1', anchor_sheet_name: 'Chart1',
    });
    expect(JSON.stringify(result)).toMatch(/is a chart sheet, which holds no cells/);
    expect(writes).toHaveLength(0);
  });

  it('refuses it as the source tab', async () => {
    stubSheets();
    const result = await call('add_chart', { ...base, sheet_name: 'Chart1', anchor_cell: 'A1' });
    expect(JSON.stringify(result)).toMatch(/is a chart sheet, which holds no cells/);
  });

  it('still accepts a grid tab', async () => {
    const writes = stubSheets();
    const result = await call('add_chart', {
      ...base, sheet_name: 'Data', anchor_cell: 'F2',
    });
    expect(JSON.stringify(result)).toContain('"chartId":7');
    expect(writes).toHaveLength(1);
  });

  it('names an unknown tab rather than reporting it as a chart sheet', async () => {
    stubSheets();
    const result = await call('add_chart', { ...base, sheet_name: 'Nope', anchor_cell: 'A1' });
    expect(JSON.stringify(result)).toMatch(/Sheet tab .*Nope.* not found/);
  });
});
