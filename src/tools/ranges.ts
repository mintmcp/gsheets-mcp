/**
 * Tools that attach behaviour to a range rather than content: data
 * validation rules, named ranges, the basic filter and filter views. All are
 * single `batchUpdate` requests, so all are native-only.
 */

import { z } from 'zod';
import { withGoogleAuth as requirePermissionSecure } from '../auth.js';
import { wrapHandler, toolResponse } from '../lib/errors.js';
import { assertBareA1Range, parseA1Range, quoteSheetName } from '../lib/a1.js';
import { getSheetId, makeSheetsRequest, resolveGridRange } from '../lib/google.js';
import { applyBatchUpdate } from '../lib/batch.js';
import { nativeOnly } from '../lib/office.js';
import {
  validationConditionSchema,
  buildValidationCondition,
  sortBySchema,
  buildSortSpecs,
  filterCriteriaSchema,
  buildFilterSpecs,
  type SortKey,
  type FilterCriterion,
} from '../lib/criteria.js';

const spreadsheetId = z.string().describe('Google Sheets spreadsheet ID');
const sheetName = z.string().describe('Name of the sheet tab');
const rangeSchema = z.string().describe('Bounded bare A1 range: "A1" or "A1:C3". Whole-column/whole-row forms are not supported.');

const idAndMessage = {
  id: z.string(),
  message: z.string(),
};

const SPREADSHEETS_SCOPE = 'https://www.googleapis.com/auth/spreadsheets';

const filterInputs = {
  spreadsheet_id: spreadsheetId,
  sheet_name: sheetName,
  range: z.string().optional().describe('Bounded A1 range the filter covers, header row included, e.g. "A1:D50". Omit to cover the whole tab.'),
  sort_by: sortBySchema.optional(),
  criteria: filterCriteriaSchema.optional(),
};

interface FilterInput {
  sheet_name: string;
  range?: string;
  sort_by?: SortKey[];
  criteria?: FilterCriterion[];
}

interface Filter {
  range: { sheetId: number } & Partial<ReturnType<typeof parseA1Range>>;
  sortSpecs?: ReturnType<typeof buildSortSpecs>;
  filterSpecs?: ReturnType<typeof buildFilterSpecs>;
}

/**
 * A filter's shape is the same whether it is the tab's basic filter or a
 * saved view: a range (the whole tab when omitted), sort keys and per-column
 * criteria. Column letters are checked against the range before anything
 * is sent, since Google's answer to an out-of-range column is a bare 400.
 */
function buildFilter(sheetId: number, input: FilterInput): { filter: Filter; where: string } {
  const a1 = input.range !== undefined ? assertBareA1Range(input.range) : undefined;
  const rect = a1 !== undefined ? parseA1Range(a1) : undefined;
  const where = a1 !== undefined ? `${a1} on "${input.sheet_name}"` : `"${input.sheet_name}"`;
  return {
    filter: {
      range: { sheetId, ...rect },
      ...(input.sort_by && { sortSpecs: buildSortSpecs(input.sort_by, rect, where) }),
      ...(input.criteria && { filterSpecs: buildFilterSpecs(input.criteria, rect, where) }),
    },
    where,
  };
}

export const rangeTools = {
      set_data_validation: {
        description: 'Restrict what a range of cells accepts, as Data > Data validation does. Common rules: a dropdown from fixed options ({"type":"ONE_OF_LIST","values":["Yes","No"]}), a dropdown from cells ({"type":"ONE_OF_RANGE","values":["Lists!A1:A9"]}), a checkbox ({"type":"BOOLEAN"}), a number bound ({"type":"NUMBER_BETWEEN","values":["1","10"]}), a valid date ({"type":"DATE_IS_VALID"}), or any formula ({"type":"CUSTOM_FORMULA","values":["=ISEMAIL(A1)"]}). By default invalid input is rejected (strict) and list rules show a dropdown. Existing values are not re-checked. Replaces any rule already on the range; use clear_data_validation to remove rules.',
        outputSchema: idAndMessage,
        schema: {
          spreadsheet_id: spreadsheetId,
          sheet_name: sheetName,
          range: rangeSchema,
          condition: validationConditionSchema,
          strict: z.boolean().optional().describe('true (default) rejects invalid input; false allows it with a warning'),
          show_dropdown: z.boolean().optional().describe('For ONE_OF_LIST and ONE_OF_RANGE, show the in-cell dropdown (default true)'),
          input_message: z.string().optional().describe('Help text shown when the cell is selected'),
        },
        handler: requirePermissionSecure(SPREADSHEETS_SCOPE, wrapHandler(nativeOnly(async ({
          spreadsheet_id, sheet_name, range, condition, strict = true, show_dropdown = true, input_message,
        }: any, context: any) => {
          const { accessToken } = context;
          const rule = {
            condition: buildValidationCondition(condition),
            strict,
            showCustomUi: show_dropdown,
            ...(input_message !== undefined && { inputMessage: input_message }),
          };
          const { a1, grid } = await resolveGridRange(spreadsheet_id, sheet_name, range, accessToken);
          await applyBatchUpdate(spreadsheet_id, [{ setDataValidation: { range: grid, rule } }], accessToken);
          return toolResponse({
            id: spreadsheet_id,
            message: `Data validation (${condition.type}) set on ${a1} of "${sheet_name}"`,
          });
        }))),
      },

      clear_data_validation: {
        description: 'Remove every data validation rule from a range. Cell values are untouched.',
        outputSchema: idAndMessage,
        schema: {
          spreadsheet_id: spreadsheetId,
          sheet_name: sheetName,
          range: rangeSchema,
        },
        handler: requirePermissionSecure(SPREADSHEETS_SCOPE, wrapHandler(nativeOnly(async ({ spreadsheet_id, sheet_name, range }: any, context: any) => {
          const { accessToken } = context;
          const { a1, grid } = await resolveGridRange(spreadsheet_id, sheet_name, range, accessToken);
          await applyBatchUpdate(spreadsheet_id, [{ setDataValidation: { range: grid } }], accessToken);
          return toolResponse({
            id: spreadsheet_id,
            message: `Data validation cleared from ${a1} of "${sheet_name}"`,
          });
        }))),
      },

      add_named_range: {
        description: 'Name a range so formulas and other tools can refer to it by name (e.g. =SUM(Prices)). Names must start with a letter or underscore, use only letters, digits and underscores, not look like a cell reference, and be unique in the spreadsheet; Google rejects anything else. get_metadata lists existing named ranges.',
        outputSchema: {
          id: z.string(),
          namedRangeId: z.string().describe('Google\'s id for the named range'),
          name: z.string(),
          range: z.string().describe('The sheet-qualified A1 range it covers'),
          message: z.string(),
        },
        schema: {
          spreadsheet_id: spreadsheetId,
          name: z.string().min(1).describe('The name, e.g. "Prices" or "q3_totals"'),
          sheet_name: sheetName,
          range: rangeSchema,
        },
        handler: requirePermissionSecure(SPREADSHEETS_SCOPE, wrapHandler(nativeOnly(async ({ spreadsheet_id, name, sheet_name, range }: any, context: any) => {
          const { accessToken } = context;
          const { a1, grid } = await resolveGridRange(spreadsheet_id, sheet_name, range, accessToken);
          const [reply] = await applyBatchUpdate<{ addNamedRange?: { namedRange?: { namedRangeId?: string } } }>(
            spreadsheet_id,
            [{ addNamedRange: { namedRange: { name, range: grid } } }],
            accessToken,
          );
          const namedRangeId = reply?.addNamedRange?.namedRange?.namedRangeId;
          if (!namedRangeId) throw new Error('Google created the named range but returned no id for it');
          const qualified = `${quoteSheetName(sheet_name)}!${a1}`;
          return toolResponse({
            id: spreadsheet_id,
            namedRangeId,
            name,
            range: qualified,
            message: `Named range "${name}" now refers to ${qualified}`,
          });
        }))),
      },

      delete_named_range: {
        description: 'Delete a named range by name. Cells are untouched; formulas that used the name will show #REF!.',
        destructiveHint: true,
        outputSchema: idAndMessage,
        schema: {
          spreadsheet_id: spreadsheetId,
          name: z.string().min(1).describe('Name of the named range to delete'),
        },
        handler: requirePermissionSecure(SPREADSHEETS_SCOPE, wrapHandler(nativeOnly(async ({ spreadsheet_id, name }: any, context: any) => {
          const { accessToken } = context;
          const { namedRanges = [] } = await makeSheetsRequest(
            `/${encodeURIComponent(spreadsheet_id)}?fields=namedRanges(namedRangeId,name)`,
            accessToken,
            { method: 'GET' },
          ) as { namedRanges?: Array<{ namedRangeId: string; name: string }> };
          const target = namedRanges.find((r) => r.name === name);
          if (!target) throw new Error(`Named range "${name}" not found`);
          await applyBatchUpdate(spreadsheet_id, [{ deleteNamedRange: { namedRangeId: target.namedRangeId } }], accessToken);
          return toolResponse({
            id: spreadsheet_id,
            message: `Named range "${name}" deleted`,
          });
        }))),
      },

      set_basic_filter: {
        description: 'Turn on the tab\'s filter (Data > Create a filter), optionally sorted and with rows hidden by value or condition. A tab has one basic filter; setting it replaces the previous one. Pass `range` with the header row included, or omit it to filter the whole tab. `criteria` names columns by letter: hidden_values hides rows showing those values, condition keeps only rows passing a test (e.g. {"type":"NUMBER_GREATER","values":["10"]}). Hidden rows are still returned by get_sheet_data. Use clear_basic_filter to turn the filter off.',
        outputSchema: idAndMessage,
        schema: filterInputs,
        handler: requirePermissionSecure(SPREADSHEETS_SCOPE, wrapHandler(nativeOnly(async (args: any, context: any) => {
          const { accessToken } = context;
          const sheetId = await getSheetId(args.spreadsheet_id, args.sheet_name, accessToken);
          const { filter, where } = buildFilter(sheetId, args);
          await applyBatchUpdate(args.spreadsheet_id, [{ setBasicFilter: { filter } }], accessToken);
          return toolResponse({
            id: args.spreadsheet_id,
            message: `Basic filter set on ${where}`,
          });
        }))),
      },

      clear_basic_filter: {
        description: 'Turn off the tab\'s basic filter, showing every row again. Saved filter views are untouched.',
        outputSchema: idAndMessage,
        schema: {
          spreadsheet_id: spreadsheetId,
          sheet_name: sheetName,
        },
        handler: requirePermissionSecure(SPREADSHEETS_SCOPE, wrapHandler(nativeOnly(async ({ spreadsheet_id, sheet_name }: any, context: any) => {
          const { accessToken } = context;
          const sheetId = await getSheetId(spreadsheet_id, sheet_name, accessToken);
          await applyBatchUpdate(spreadsheet_id, [{ clearBasicFilter: { sheetId } }], accessToken);
          return toolResponse({
            id: spreadsheet_id,
            message: `Basic filter cleared from "${sheet_name}"`,
          });
        }))),
      },

      add_filter_view: {
        description: 'Save a named filter view (Data > Filter views) that users can switch to without changing what others see. Takes the same range, sort_by and criteria as set_basic_filter. get_metadata lists a tab\'s filter views with their ids.',
        outputSchema: {
          id: z.string(),
          filterViewId: z.number().describe('Google\'s id for the view, needed to delete it'),
          title: z.string(),
          message: z.string(),
        },
        schema: {
          ...filterInputs,
          title: z.string().min(1).describe('Name shown in the filter views menu'),
        },
        handler: requirePermissionSecure(SPREADSHEETS_SCOPE, wrapHandler(nativeOnly(async (args: any, context: any) => {
          const { accessToken } = context;
          const sheetId = await getSheetId(args.spreadsheet_id, args.sheet_name, accessToken);
          const { filter, where } = buildFilter(sheetId, args);
          const [reply] = await applyBatchUpdate<{ addFilterView?: { filter?: { filterViewId?: number } } }>(
            args.spreadsheet_id,
            [{ addFilterView: { filter: { title: args.title, ...filter } } }],
            accessToken,
          );
          const filterViewId = reply?.addFilterView?.filter?.filterViewId;
          if (filterViewId === undefined) throw new Error('Google created the filter view but returned no id for it');
          return toolResponse({
            id: args.spreadsheet_id,
            filterViewId,
            title: args.title,
            message: `Filter view "${args.title}" added on ${where}`,
          });
        }))),
      },

      delete_filter_view: {
        description: 'Delete a saved filter view by id (from add_filter_view or get_metadata).',
        destructiveHint: true,
        outputSchema: idAndMessage,
        schema: {
          spreadsheet_id: spreadsheetId,
          filter_view_id: z.coerce.number().int().describe('The filter view\'s id'),
        },
        handler: requirePermissionSecure(SPREADSHEETS_SCOPE, wrapHandler(nativeOnly(async ({ spreadsheet_id, filter_view_id }: any, context: any) => {
          const { accessToken } = context;
          await applyBatchUpdate(spreadsheet_id, [{ deleteFilterView: { filterId: Number(filter_view_id) } }], accessToken);
          return toolResponse({
            id: spreadsheet_id,
            message: `Filter view ${filter_view_id} deleted`,
          });
        }))),
      },
};
