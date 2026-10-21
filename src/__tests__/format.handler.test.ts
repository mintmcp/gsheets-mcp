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
  const sent: any[] = [];
  vi.stubGlobal('fetch', async (_input: any, init?: any) => {
    if ((init?.method ?? 'GET') === 'GET') {
      return jsonResponse({ sheets: [
        { properties: { sheetId: 11, title: 'Data' } },
        { properties: { sheetId: 22, title: 'Other' } },
      ] });
    }
    sent.push(JSON.parse(init.body));
    return jsonResponse({ replies: [{}] });
  });
  return sent;
}

const call = (tool: string, args: Record<string, unknown>) =>
  requestContext.run({ accessToken: 'test-token' }, () =>
    (formatTools as any)[tool].handler(args),
  );

afterEach(() => vi.unstubAllGlobals());

describe.each(['format_cells', 'clear_formatting'])('%s', (tool) => {
  const extra = tool === 'format_cells' ? { format: { textFormat: { bold: true } } } : {};

  it('takes the tab from a qualified range when sheet_name is omitted', async () => {
    const sent = stubSheets();
    await call(tool, { spreadsheet_id: 's', range: 'Other!A1:B2', ...extra });
    expect(sent[0].requests[0].repeatCell.range).toEqual({
      sheetId: 22, startRowIndex: 0, endRowIndex: 2, startColumnIndex: 0, endColumnIndex: 2,
    });
  });

  it('needs a tab from somewhere', async () => {
    const sent = stubSheets();
    const res = await call(tool, { spreadsheet_id: 's', range: 'A1', ...extra });
    expect(res.isError).toBe(true);
    expect(sent).toHaveLength(0);
  });
});
