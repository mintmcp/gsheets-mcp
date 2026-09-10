import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { tools } from '../tools/index.js';
import { requestContext } from '../auth.js';
import { stubFetch, jsonResponse, LABEL_SCHEMA_BODY } from './labelStubs.js';

const SHEET_META = {
  properties: { title: 'Budget' },
  sheets: [{ properties: { title: 'Sheet1', index: 0, gridProperties: { rowCount: 10, columnCount: 3 } } }],
  spreadsheetUrl: 'https://docs.google.com/spreadsheets/d/s1',
};

const LISTED_LABEL = {
  labels: [{
    id: 'lbl1', revisionId: 'rev7',
    fields: { field1: { valueType: 'selection', selection: ['choiceA'] } },
  }],
};

const call = (tool: string, args: any) =>
  requestContext.run({ accessToken: 'tok' }, () => (tools as any)[tool].handler(args));

describe('label enrichment on read tools', () => {
  const savedProfile = process.env.PROFILE;
  beforeEach(() => {
    // the unrestricted cases need PROFILE genuinely unset, whatever the
    // ambient environment carries
    delete process.env.PROFILE;
  });
  afterEach(() => {
    vi.unstubAllGlobals();
    vi.unstubAllEnvs();
    if (savedProfile === undefined) delete process.env.PROFILE;
    else process.env.PROFILE = savedProfile;
  });

  it('get_metadata returns visible labels and _meta.applied when enrichment is on', async () => {
    stubFetch([
      ['sheets.googleapis.com', () => jsonResponse(SHEET_META)],
      ['listLabels', () => jsonResponse(LISTED_LABEL)],
      ['drivelabels.googleapis.com', () => jsonResponse({ properties: { title: 'Classification' }, ...LABEL_SCHEMA_BODY })],
    ]);
    const res: any = await call('get_metadata', { spreadsheet_id: 's1' });
    expect(res.isError).toBeUndefined();
    expect(res.structuredContent.labels).toEqual([{
      labelId: 'lbl1', revisionId: 'rev7', title: 'Classification', resolved: true,
      values: [{ fieldId: 'field1', valueType: 'selection', choiceId: 'choiceA', displayName: 'Confidential', resolved: true }],
    }]);
    expect(res.structuredContent.labelsError).toBeUndefined();
    expect(res._meta.applied).toEqual(res.structuredContent.labels);
  });

  it('a standard deployment makes no label calls and returns no label surface', async () => {
    vi.stubEnv('PROFILE', 'standard');
    const calls = stubFetch([
      ['sheets.googleapis.com', () => jsonResponse(SHEET_META)],
    ]);
    const res: any = await call('get_metadata', { spreadsheet_id: 's1' });
    expect(res.isError).toBeUndefined();
    expect(res.structuredContent.labels).toBeUndefined();
    expect(res._meta).toBeUndefined();
    expect(calls.some((c) => c.url.includes('listLabels'))).toBe(false);
  });

  it('a labels deployment enriches like unrestricted does', async () => {
    vi.stubEnv('PROFILE', 'labels');
    stubFetch([
      ['sheets.googleapis.com', () => jsonResponse(SHEET_META)],
      ['listLabels', () => jsonResponse(LISTED_LABEL)],
      ['drivelabels.googleapis.com', () => jsonResponse({ properties: { title: 'Classification' }, ...LABEL_SCHEMA_BODY })],
    ]);
    const res: any = await call('get_metadata', { spreadsheet_id: 's1' });
    expect(res._meta.applied).toHaveLength(1);
    expect(res.structuredContent.labels).toHaveLength(1);
  });

  it('a malformed label wire degrades to labelsError instead of failing the read', async () => {
    stubFetch([
      ['sheets.googleapis.com', () => jsonResponse(SHEET_META)],
      ['listLabels', () => jsonResponse({ labels: [null] })],
    ]);
    const res: any = await call('get_metadata', { spreadsheet_id: 's1' });
    expect(res.isError).toBeUndefined();
    expect(res.structuredContent.title).toBe('Budget');
    expect(res.structuredContent.labels).toEqual([]);
    expect(res.structuredContent.labelsError).toBe('label read failed');
    expect(res._meta.labelsError).toBe('label read failed');
  });

  it('get_sheet_data error envelopes still carry _meta, without a visible labels field', async () => {
    stubFetch([
      ['sheets.googleapis.com', () => jsonResponse({ error: { message: 'boom' } }, 500)],
      ['listLabels', () => jsonResponse(LISTED_LABEL)],
      ['drivelabels.googleapis.com', () => jsonResponse({ properties: { title: 'Classification' }, ...LABEL_SCHEMA_BODY })],
    ]);
    const res: any = await call('get_sheet_data', { spreadsheet_id: 's1' });
    expect(res.isError).toBe(true);
    expect(res.structuredContent).toBeUndefined();
    expect(res._meta.applied).toHaveLength(1);
  });
});
