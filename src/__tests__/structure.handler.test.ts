import { describe, it, expect, afterEach, vi } from 'vitest';
import { structureTools } from '../tools/structure.js';
import { requestContext } from '../auth.js';

/**
 * Drives the structure handlers against a stubbed Sheets API and asserts the
 * batchUpdate request each one sends, since the request body is the whole
 * behaviour of these tools.
 */

function jsonResponse(body: unknown) {
  return new Response(JSON.stringify(body), {
    status: 200,
    headers: { 'content-type': 'application/json' },
  });
}

function stubSheets(replies: unknown[] = [{}]) {
  const writes: any[] = [];
  vi.stubGlobal('fetch', async (input: any, init?: any) => {
    const url = typeof input === 'string' ? input : String(input?.url ?? input);
    if (init?.method === 'POST') {
      writes.push(JSON.parse(init.body));
      return jsonResponse({ replies });
    }
    if (url.includes('sheets.properties')) {
      return jsonResponse({
        sheets: [
          { properties: { sheetId: 11, title: 'Data', index: 0 } },
          { properties: { sheetId: 22, title: 'Other', index: 1 } },
        ],
      });
    }
    throw new Error(`unexpected request ${url}`);
  });
  return writes;
}

const call = (tool: string, args: Record<string, unknown>) =>
  requestContext.run({ accessToken: 'test-token' }, () =>
    (structureTools as any)[tool].handler(args),
  );

const payload = (res: any) => res.structuredContent ?? JSON.parse(res.content[0].text);

afterEach(() => vi.unstubAllGlobals());

describe('rename_sheet', () => {
  it('sends updateSheetProperties with only the title field', async () => {
    const writes = stubSheets();
    const out = payload(await call('rename_sheet', {
      spreadsheet_id: 's', sheet_name: 'Data', new_name: 'Sales',
    }));
    expect(writes[0].requests).toEqual([{
      updateSheetProperties: { properties: { sheetId: 11, title: 'Sales' }, fields: 'title' },
    }]);
    expect(out.sheetTitle).toBe('Sales');
  });

  it('names an unknown tab without sending anything', async () => {
    const writes = stubSheets();
    const res = await call('rename_sheet', { spreadsheet_id: 's', sheet_name: 'Nope', new_name: 'X' });
    expect(res.isError).toBe(true);
    expect(JSON.parse(res.content[0].text).error).toBe('Sheet tab "Nope" not found');
    expect(writes).toHaveLength(0);
  });
});

describe('delete_sheet', () => {
  it('sends deleteSheet for the resolved id', async () => {
    const writes = stubSheets();
    const out = payload(await call('delete_sheet', { spreadsheet_id: 's', sheet_name: 'Other' }));
    expect(writes[0].requests).toEqual([{ deleteSheet: { sheetId: 22 } }]);
    expect(out.message).toMatch(/"Other" deleted/);
  });
});

describe('duplicate_sheet', () => {
  it('copies the tab and reports the name Google assigned', async () => {
    const writes = stubSheets([{ duplicateSheet: { properties: { sheetId: 33, title: 'Copy of Data' } } }]);
    const out = payload(await call('duplicate_sheet', { spreadsheet_id: 's', sheet_name: 'Data' }));
    expect(writes[0].requests).toEqual([{ duplicateSheet: { sourceSheetId: 11, insertSheetIndex: 1 } }]);
    expect(out.sheetTitle).toBe('Copy of Data');
  });

  it('passes a requested name through', async () => {
    const writes = stubSheets([{ duplicateSheet: { properties: { sheetId: 33, title: 'Backup' } } }]);
    await call('duplicate_sheet', { spreadsheet_id: 's', sheet_name: 'Data', new_name: 'Backup' });
    expect(writes[0].requests[0].duplicateSheet).toEqual({ sourceSheetId: 11, insertSheetIndex: 1, newSheetName: 'Backup' });
  });
});

describe('insert_rows_or_columns', () => {
  it('inserts rows before a 1-based row number', async () => {
    const writes = stubSheets();
    const out = payload(await call('insert_rows_or_columns', {
      spreadsheet_id: 's', sheet_name: 'Data', dimension: 'ROWS', start_index: 5, count: 2,
    }));
    expect(writes[0].requests).toEqual([{
      insertDimension: {
        range: { sheetId: 11, dimension: 'ROWS', startIndex: 4, endIndex: 6 },
        inheritFromBefore: true,
      },
    }]);
    expect(out.message).toBe('rows 5-6 inserted');
  });

  it('cannot inherit from before the first row, which Google refuses', async () => {
    const writes = stubSheets();
    await call('insert_rows_or_columns', {
      spreadsheet_id: 's', sheet_name: 'Data', dimension: 'ROWS', start_index: 1,
    });
    expect(writes[0].requests[0].insertDimension.inheritFromBefore).toBe(false);
  });

  it('accepts a column letter for COLUMNS', async () => {
    const writes = stubSheets();
    const out = payload(await call('insert_rows_or_columns', {
      spreadsheet_id: 's', sheet_name: 'Data', dimension: 'COLUMNS', start_index: 'C', count: 1,
    }));
    expect(writes[0].requests[0].insertDimension.range).toEqual({
      sheetId: 11, dimension: 'COLUMNS', startIndex: 2, endIndex: 3,
    });
    expect(out.message).toBe('column C inserted');
  });

  it('rejects a column letter for ROWS and a row number for COLUMNS', async () => {
    const writes = stubSheets();
    const rows = await call('insert_rows_or_columns', {
      spreadsheet_id: 's', sheet_name: 'Data', dimension: 'ROWS', start_index: 'C', count: 1,
    });
    expect(rows.isError).toBe(true);
    expect(rows.content[0].text).toMatch(/1-based row number/);
    const cols = await call('insert_rows_or_columns', {
      spreadsheet_id: 's', sheet_name: 'Data', dimension: 'COLUMNS', start_index: 3, count: 1,
    });
    expect(cols.isError).toBe(true);
    expect(cols.content[0].text).toMatch(/column letter/);
    expect(writes).toHaveLength(0);
  });

  it('rejects row 0 and a zero count', async () => {
    const writes = stubSheets();
    const zero = await call('insert_rows_or_columns', {
      spreadsheet_id: 's', sheet_name: 'Data', dimension: 'ROWS', start_index: 0, count: 1,
    });
    expect(zero.isError).toBe(true);
    const none = await call('insert_rows_or_columns', {
      spreadsheet_id: 's', sheet_name: 'Data', dimension: 'ROWS', start_index: 1, count: 0,
    });
    expect(none.isError).toBe(true);
    expect(writes).toHaveLength(0);
  });
});

describe('delete_rows_or_columns', () => {
  it('deletes a run of rows', async () => {
    const writes = stubSheets();
    const out = payload(await call('delete_rows_or_columns', {
      spreadsheet_id: 's', sheet_name: 'Data', dimension: 'ROWS', start_index: 2, count: 3,
    }));
    expect(writes[0].requests).toEqual([{
      deleteDimension: { range: { sheetId: 11, dimension: 'ROWS', startIndex: 1, endIndex: 4 } },
    }]);
    expect(out.message).toMatch(/rows 2-4 deleted/);
  });

  it('deletes a single column by letter', async () => {
    const writes = stubSheets();
    const out = payload(await call('delete_rows_or_columns', {
      spreadsheet_id: 's', sheet_name: 'Data', dimension: 'COLUMNS', start_index: 'AA',
    }));
    expect(writes[0].requests[0].deleteDimension.range).toEqual({
      sheetId: 11, dimension: 'COLUMNS', startIndex: 26, endIndex: 27,
    });
    expect(out.message).toMatch(/column AA deleted/);
  });
});

describe('freeze_panes', () => {
  it('sets both frozen counts in one request', async () => {
    const writes = stubSheets();
    await call('freeze_panes', { spreadsheet_id: 's', sheet_name: 'Data', rows: 1, columns: 2 });
    expect(writes[0].requests).toEqual([{
      updateSheetProperties: {
        properties: { sheetId: 11, gridProperties: { frozenRowCount: 1, frozenColumnCount: 2 } },
        fields: 'gridProperties.frozenRowCount,gridProperties.frozenColumnCount',
      },
    }]);
  });

  it('touches only the count that was passed', async () => {
    const writes = stubSheets();
    await call('freeze_panes', { spreadsheet_id: 's', sheet_name: 'Data', rows: 0 });
    expect(writes[0].requests[0].updateSheetProperties).toEqual({
      properties: { sheetId: 11, gridProperties: { frozenRowCount: 0 } },
      fields: 'gridProperties.frozenRowCount',
    });
  });

  it('refuses a call that freezes nothing', async () => {
    const writes = stubSheets();
    const res = await call('freeze_panes', { spreadsheet_id: 's', sheet_name: 'Data' });
    expect(res.isError).toBe(true);
    expect(writes).toHaveLength(0);
  });
});

describe('sort_range', () => {
  it('sorts by absolute column index within the range', async () => {
    const writes = stubSheets();
    await call('sort_range', {
      spreadsheet_id: 's', sheet_name: 'Data', range: 'B1:D20',
      sort_by: [{ column: 'C', descending: true }, { column: 'B' }],
    });
    expect(writes[0].requests).toEqual([{
      sortRange: {
        range: { sheetId: 11, startRowIndex: 0, endRowIndex: 20, startColumnIndex: 1, endColumnIndex: 4 },
        sortSpecs: [
          { dimensionIndex: 2, sortOrder: 'DESCENDING' },
          { dimensionIndex: 1, sortOrder: 'ASCENDING' },
        ],
      },
    }]);
  });

  it('keeps a header row in place when asked', async () => {
    const writes = stubSheets();
    const out = payload(await call('sort_range', {
      spreadsheet_id: 's', sheet_name: 'Data', range: 'A1:C10', has_header: true,
      sort_by: [{ column: 'A' }],
    }));
    expect(writes[0].requests[0].sortRange.range.startRowIndex).toBe(1);
    expect(out.message).toMatch(/A2:C10/);
  });

  it('rejects a sort column outside the range', async () => {
    const writes = stubSheets();
    const res = await call('sort_range', {
      spreadsheet_id: 's', sheet_name: 'Data', range: 'B1:D20', sort_by: [{ column: 'E' }],
    });
    expect(res.isError).toBe(true);
    expect(res.content[0].text).toMatch(/column E is outside/);
    expect(writes).toHaveLength(0);
  });

  it('rejects a header-only range', async () => {
    const writes = stubSheets();
    const res = await call('sort_range', {
      spreadsheet_id: 's', sheet_name: 'Data', range: 'A1:C1', has_header: true, sort_by: [{ column: 'A' }],
    });
    expect(res.isError).toBe(true);
    expect(writes).toHaveLength(0);
  });
});

describe('find_replace', () => {
  const reply = { findReplace: { occurrencesChanged: 3, valuesChanged: 2, formulasChanged: 1, rowsChanged: 3, sheetsChanged: 1 } };

  it('searches every tab when no sheet is named', async () => {
    const writes = stubSheets([reply]);
    const out = payload(await call('find_replace', { spreadsheet_id: 's', find: 'foo', replacement: 'bar' }));
    expect(writes[0].requests).toEqual([{
      findReplace: {
        find: 'foo', replacement: 'bar', allSheets: true,
        matchCase: false, matchEntireCell: false, searchByRegex: false, includeFormulas: false,
      },
    }]);
    expect(out.occurrencesChanged).toBe(3);
    expect(out.formulasChanged).toBe(1);
  });

  it('scopes to a tab, and to a range within it', async () => {
    const writes = stubSheets([reply]);
    await call('find_replace', { spreadsheet_id: 's', sheet_name: 'Data', find: 'a', replacement: 'b' });
    expect(writes[0].requests[0].findReplace.sheetId).toBe(11);
    expect(writes[0].requests[0].findReplace.allSheets).toBeUndefined();

    await call('find_replace', { spreadsheet_id: 's', sheet_name: 'Data', range: 'A1:B5', find: 'a', replacement: 'b' });
    expect(writes[1].requests[0].findReplace.range).toEqual({
      sheetId: 11, startRowIndex: 0, endRowIndex: 5, startColumnIndex: 0, endColumnIndex: 2,
    });
    expect(writes[1].requests[0].findReplace.sheetId).toBeUndefined();
  });

  it('requires a sheet_name alongside a range', async () => {
    const writes = stubSheets([reply]);
    const res = await call('find_replace', { spreadsheet_id: 's', range: 'A1:B5', find: 'a', replacement: 'b' });
    expect(res.isError).toBe(true);
    expect(res.content[0].text).toMatch(/sheet_name/);
    expect(writes).toHaveLength(0);
  });

  it('forwards the matching flags', async () => {
    const writes = stubSheets([reply]);
    await call('find_replace', {
      spreadsheet_id: 's', find: '^a.*', replacement: 'b',
      match_case: true, match_entire_cell: true, use_regex: true, include_formulas: true,
    });
    expect(writes[0].requests[0].findReplace).toMatchObject({
      matchCase: true, matchEntireCell: true, searchByRegex: true, includeFormulas: true,
    });
  });
});
