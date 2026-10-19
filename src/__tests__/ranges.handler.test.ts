import { describe, it, expect, afterEach, vi } from 'vitest';
import { rangeTools } from '../tools/ranges.js';
import { requestContext } from '../auth.js';

/**
 * Drives the validation, named-range and filter handlers against a stubbed
 * Sheets API and asserts the batchUpdate request each one sends.
 */

function jsonResponse(body: unknown) {
  return new Response(JSON.stringify(body), {
    status: 200,
    headers: { 'content-type': 'application/json' },
  });
}

function stubSheets(replies: unknown[] = [{}], namedRanges: unknown[] = []) {
  const writes: any[] = [];
  vi.stubGlobal('fetch', async (input: any, init?: any) => {
    const url = typeof input === 'string' ? input : String(input?.url ?? input);
    if (init?.method === 'POST') {
      writes.push(JSON.parse(init.body));
      return jsonResponse({ replies });
    }
    if (url.includes('namedRanges')) return jsonResponse({ namedRanges });
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
    (rangeTools as any)[tool].handler(args),
  );

const payload = (res: any) => res.structuredContent ?? JSON.parse(res.content[0].text);
const error = (res: any) => JSON.parse(res.content[0].text).error as string;

afterEach(() => vi.unstubAllGlobals());

const grid = { sheetId: 11, startRowIndex: 0, endRowIndex: 10, startColumnIndex: 1, endColumnIndex: 2 };

describe('set_data_validation', () => {
  it('sends a dropdown list rule with strict rejection and the dropdown shown', async () => {
    const writes = stubSheets();
    const out = payload(await call('set_data_validation', {
      spreadsheet_id: 's', sheet_name: 'Data', range: 'B1:B10',
      condition: { type: 'ONE_OF_LIST', values: ['Yes', 'No'] },
    }));
    expect(writes[0].requests).toEqual([{
      setDataValidation: {
        range: grid,
        rule: {
          condition: { type: 'ONE_OF_LIST', values: [{ userEnteredValue: 'Yes' }, { userEnteredValue: 'No' }] },
          strict: true,
          showCustomUi: true,
        },
      },
    }]);
    expect(out.message).toMatch(/B1:B10/);
  });

  it('prefixes a range source with = so Google reads it as a reference', async () => {
    const writes = stubSheets();
    await call('set_data_validation', {
      spreadsheet_id: 's', sheet_name: 'Data', range: 'B1:B10',
      condition: { type: 'ONE_OF_RANGE', values: ["Other!A1:A5"] },
      strict: false, input_message: 'Pick one', show_dropdown: false,
    });
    expect(writes[0].requests[0].setDataValidation.rule).toEqual({
      condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: '=Other!A1:A5' }] },
      strict: false,
      showCustomUi: false,
      inputMessage: 'Pick one',
    });
  });

  it('builds a checkbox rule from BOOLEAN with no values', async () => {
    const writes = stubSheets();
    await call('set_data_validation', {
      spreadsheet_id: 's', sheet_name: 'Data', range: 'B1:B10', condition: { type: 'BOOLEAN' },
    });
    expect(writes[0].requests[0].setDataValidation.rule.condition).toEqual({ type: 'BOOLEAN' });
  });

  it('rejects an unknown condition type before sending anything', async () => {
    const writes = stubSheets();
    const res = await call('set_data_validation', {
      spreadsheet_id: 's', sheet_name: 'Data', range: 'B1:B10', condition: { type: 'IS_NICE' },
    });
    expect(res.isError).toBe(true);
    expect(error(res)).toMatch(/condition.type/);
    expect(writes).toHaveLength(0);
  });
});

describe('clear_data_validation', () => {
  it('sends setDataValidation with no rule', async () => {
    const writes = stubSheets();
    await call('clear_data_validation', { spreadsheet_id: 's', sheet_name: 'Data', range: 'B1:B10' });
    expect(writes[0].requests).toEqual([{ setDataValidation: { range: grid } }]);
  });
});

describe('named ranges', () => {
  it('adds a named range and returns the id Google assigned', async () => {
    const writes = stubSheets([{ addNamedRange: { namedRange: { namedRangeId: 'nr1', name: 'Prices' } } }]);
    const out = payload(await call('add_named_range', {
      spreadsheet_id: 's', name: 'Prices', sheet_name: 'Data', range: 'B1:B10',
    }));
    expect(writes[0].requests).toEqual([{ addNamedRange: { namedRange: { name: 'Prices', range: grid } } }]);
    expect(out).toMatchObject({ namedRangeId: 'nr1', name: 'Prices', range: "'Data'!B1:B10" });
  });

  it('deletes by name after resolving the id', async () => {
    const writes = stubSheets([{}], [{ namedRangeId: 'nr1', name: 'Prices' }, { namedRangeId: 'nr2', name: 'Costs' }]);
    await call('delete_named_range', { spreadsheet_id: 's', name: 'Costs' });
    expect(writes[0].requests).toEqual([{ deleteNamedRange: { namedRangeId: 'nr2' } }]);
  });

  it('names an unknown named range without sending anything', async () => {
    const writes = stubSheets([{}], [{ namedRangeId: 'nr1', name: 'Prices' }]);
    const res = await call('delete_named_range', { spreadsheet_id: 's', name: 'Nope' });
    expect(res.isError).toBe(true);
    expect(error(res)).toMatch(/"Nope" not found/);
    expect(writes).toHaveLength(0);
  });
});

describe('basic filter', () => {
  it('sets a filter over a range with sort and criteria', async () => {
    const writes = stubSheets();
    await call('set_basic_filter', {
      spreadsheet_id: 's', sheet_name: 'Data', range: 'A1:D50',
      sort_by: [{ column: 'B', descending: true }],
      criteria: [
        { column: 'C', hidden_values: ['Closed'] },
        { column: 'D', condition: { type: 'NUMBER_GREATER', values: ['10'] } },
      ],
    });
    expect(writes[0].requests).toEqual([{
      setBasicFilter: {
        filter: {
          range: { sheetId: 11, startRowIndex: 0, endRowIndex: 50, startColumnIndex: 0, endColumnIndex: 4 },
          sortSpecs: [{ dimensionIndex: 1, sortOrder: 'DESCENDING' }],
          filterSpecs: [
            { columnIndex: 2, filterCriteria: { hiddenValues: ['Closed'] } },
            { columnIndex: 3, filterCriteria: { condition: { type: 'NUMBER_GREATER', values: [{ userEnteredValue: '10' }] } } },
          ],
        },
      },
    }]);
  });

  it('covers the whole tab when no range is given', async () => {
    const writes = stubSheets();
    await call('set_basic_filter', { spreadsheet_id: 's', sheet_name: 'Data' });
    expect(writes[0].requests).toEqual([{ setBasicFilter: { filter: { range: { sheetId: 11 } } } }]);
  });

  it('refuses a criteria column outside the range', async () => {
    const writes = stubSheets();
    const res = await call('set_basic_filter', {
      spreadsheet_id: 's', sheet_name: 'Data', range: 'A1:B5', criteria: [{ column: 'F', hidden_values: ['x'] }],
    });
    expect(res.isError).toBe(true);
    expect(error(res)).toMatch(/column F is outside/);
    expect(writes).toHaveLength(0);
  });

  it('refuses a validation-only condition type on a filter', async () => {
    const writes = stubSheets();
    const res = await call('set_basic_filter', {
      spreadsheet_id: 's', sheet_name: 'Data', criteria: [{ column: 'A', condition: { type: 'BOOLEAN' } }],
    });
    expect(res.isError).toBe(true);
    expect(error(res)).toMatch(/criteria\[0\].condition.type/);
    expect(writes).toHaveLength(0);
  });

  it('refuses a criterion that neither hides values nor sets a condition', async () => {
    const writes = stubSheets();
    const res = await call('set_basic_filter', {
      spreadsheet_id: 's', sheet_name: 'Data', criteria: [{ column: 'A' }],
    });
    expect(res.isError).toBe(true);
    expect(writes).toHaveLength(0);
  });

  it('clears the filter', async () => {
    const writes = stubSheets();
    await call('clear_basic_filter', { spreadsheet_id: 's', sheet_name: 'Data' });
    expect(writes[0].requests).toEqual([{ clearBasicFilter: { sheetId: 11 } }]);
  });
});

describe('filter views', () => {
  it('adds a view and reports its id', async () => {
    const writes = stubSheets([{ addFilterView: { filter: { filterViewId: 77, title: 'Open' } } }]);
    const out = payload(await call('add_filter_view', {
      spreadsheet_id: 's', sheet_name: 'Data', title: 'Open', range: 'A1:D50',
      criteria: [{ column: 'C', hidden_values: ['Closed'] }],
    }));
    expect(writes[0].requests).toEqual([{
      addFilterView: {
        filter: {
          title: 'Open',
          range: { sheetId: 11, startRowIndex: 0, endRowIndex: 50, startColumnIndex: 0, endColumnIndex: 4 },
          filterSpecs: [{ columnIndex: 2, filterCriteria: { hiddenValues: ['Closed'] } }],
        },
      },
    }]);
    expect(out).toMatchObject({ filterViewId: 77, title: 'Open' });
  });

  it('deletes a view by id', async () => {
    const writes = stubSheets();
    await call('delete_filter_view', { spreadsheet_id: 's', filter_view_id: 77 });
    expect(writes[0].requests).toEqual([{ deleteFilterView: { filterId: 77 } }]);
  });
});
