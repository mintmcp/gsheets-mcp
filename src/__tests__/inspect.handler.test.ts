import { describe, it, expect, afterEach, vi } from 'vitest';
import { inspectTools, MAX_RAW_BYTES } from '../tools/inspect.js';
import { requestContext } from '../auth.js';

function stub(sheetsBody: unknown) {
  const urls: string[] = [];
  vi.stubGlobal('fetch', async (input: any) => {
    const url = decodeURIComponent(String(input));
    urls.push(url);
    const body = typeof sheetsBody === 'string' ? sheetsBody : JSON.stringify(sheetsBody);
    return new Response(body, { status: 200, headers: { 'content-type': 'application/json' } });
  });
  return urls;
}

const call = (args: Record<string, unknown>) =>
  requestContext.run({ accessToken: 't' }, () => (inspectTools as any).inspect_spreadsheet.handler(args));

afterEach(() => vi.unstubAllGlobals());

describe('inspect_spreadsheet', () => {
  it('sends the field mask and translated ranges and returns the raw document', async () => {
    const raw = { spreadsheetId: 's', sheets: [{ properties: { sheetId: 1, title: 'Data' } }] };
    const urls = stub(raw);
    const res = await call({
      spreadsheet_id: 's',
      fields: ['sheets.properties.sheetId', 'sheets.properties.title'],
      ranges: ['Data!A1:B2', 'Other!R1C1', 'Whole Tab', "'Q''s'!A:C"],
    });
    expect(res.isError).toBeUndefined();
    expect(urls).toHaveLength(1);
    expect(urls[0]).toContain('/s?');
    expect(urls[0]).toContain('fields=sheets.properties.sheetId,sheets.properties.title');
    expect(urls[0]).toContain("ranges='Data'!A1:B2");
    expect(urls[0]).toContain("ranges='Other'!A1");
    expect(urls[0]).toContain("ranges='Whole+Tab'");
    expect(urls[0]).toContain("ranges='Q''s'!A:C");
    expect(urls[0]).not.toContain('includeGridData');
    expect(res.structuredContent.spreadsheet).toEqual(raw);
    expect(res.structuredContent.bytes).toBeGreaterThan(10);
  });

  it('reads a quoted cell-like name as a tab and an unquoted one as cells, as Google does', async () => {
    const urls = stub({});
    await call({ spreadsheet_id: 's', fields: ['sheets.data'], ranges: ["'Q1'", 'Q1'] });
    expect(urls[0]).toContain("ranges='Q1'&ranges=Q1");
  });

  it('reports a malformed qualified range instead of sending it as a tab name', async () => {
    const urls = stub({});
    const res = await call({ spreadsheet_id: 's', fields: ['sheets.data'], ranges: ['Data!C3:A1'] });
    expect(res.isError).toBe(true);
    expect(JSON.parse(res.content[0].text).error).toMatch(/ranges\[0\].*reversed/);
    expect(urls).toHaveLength(0);
  });

  it('always asks for comments, so the mask alone decides whether they come back', async () => {
    const urls = stub({});
    await call({ spreadsheet_id: 's', fields: ['sheets.comment_anchors'] });
    expect(urls[0]).toContain('commentsViewMode=COMMENTS_VIEW_MODE_DEFAULT_FOR_CURRENT_ACCESS');
  });

  it('passes a plain name through for Google to resolve as a tab or named range', async () => {
    const urls = stub({});
    await call({ spreadsheet_id: 's', fields: ['sheets.data'], ranges: ['QuarterTotals', 'Q 4'] });
    expect(urls[0]).toContain("ranges=QuarterTotals&ranges='Q+4'");
  });
  it('accepts the mask as one comma-separated string too', async () => {
    const urls = stub({});
    await call({ spreadsheet_id: 's', fields: 'properties.title,sheets.properties' });
    expect(urls[0]).toContain('fields=properties.title,sheets.properties');
  });

  it('requires a non-empty mask', async () => {
    const urls = stub({});
    const res = await call({ spreadsheet_id: 's', fields: [] });
    expect(res.isError).toBe(true);
    expect(JSON.parse(res.content[0].text).error).toMatch(/fields/);
    expect(urls).toHaveLength(0);
  });

  it('rejects a malformed unqualified range instead of sending it as a tab name', async () => {
    const urls = stub({});
    const res = await call({ spreadsheet_id: 's', fields: ['sheets.data'], ranges: ['C3:A1'] });
    expect(res.isError).toBe(true);
    const relative = await call({ spreadsheet_id: 's', fields: ['sheets.data'], ranges: ['R[1]C[1]'] });
    expect(JSON.parse(relative.content[0].text).error).toMatch(/relative R1C1/);
    expect(res.isError).toBe(true);
    expect(urls).toHaveLength(0);
  });

  it('refuses an answer over the byte cap and says to narrow the mask', async () => {
    stub('{"sheets":"' + 'x'.repeat(MAX_RAW_BYTES + 10) + '"}');
    const res = await call({ spreadsheet_id: 's', fields: ['sheets'] });
    expect(res.isError).toBe(true);
    expect(JSON.parse(res.content[0].text).error).toMatch(/add `ranges`/);
  });

  it('says to shrink the ranges when the request already had them', async () => {
    stub('{"sheets":"' + 'x'.repeat(MAX_RAW_BYTES + 10) + '"}');
    const res = await call({ spreadsheet_id: 's', fields: ['sheets.data'], ranges: ['Data!A1:Z9999'] });
    expect(JSON.parse(res.content[0].text).error).toMatch(/shrink `ranges`/);
  });
});
