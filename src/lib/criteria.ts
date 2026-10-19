/**
 * The pieces of a Sheets request that describe a test on cell values: the
 * BooleanCondition behind data validation and filter criteria, and the sort
 * and filter specs that sort_range, the basic filter and filter views share.
 * They live together so the three tools that take a `condition` or a
 * `sort_by` accept exactly the same shape.
 */

import { z } from 'zod';
import { columnLetterToIndex } from './a1.js';

/**
 * Google's ConditionType enum, minus the two that belong to one feature
 * each: BOOLEAN (a checkbox) is validation-only, FILTER_EXPRESSION is
 * filter-only. Google validates how many values each type takes and answers
 * a structured 400 for a mismatch, so arity is not duplicated here.
 */
const SHARED_CONDITION_TYPES = [
  'NUMBER_GREATER', 'NUMBER_GREATER_THAN_EQ', 'NUMBER_LESS', 'NUMBER_LESS_THAN_EQ',
  'NUMBER_EQ', 'NUMBER_NOT_EQ', 'NUMBER_BETWEEN', 'NUMBER_NOT_BETWEEN',
  'TEXT_CONTAINS', 'TEXT_NOT_CONTAINS', 'TEXT_STARTS_WITH', 'TEXT_ENDS_WITH',
  'TEXT_EQ', 'TEXT_NOT_EQ', 'TEXT_IS_EMAIL', 'TEXT_IS_URL',
  'DATE_EQ', 'DATE_NOT_EQ', 'DATE_BEFORE', 'DATE_AFTER', 'DATE_ON_OR_BEFORE', 'DATE_ON_OR_AFTER',
  'DATE_BETWEEN', 'DATE_NOT_BETWEEN', 'DATE_IS_VALID',
  'ONE_OF_RANGE', 'ONE_OF_LIST', 'BLANK', 'NOT_BLANK', 'CUSTOM_FORMULA',
] as const;

export const VALIDATION_CONDITION_TYPES = [...SHARED_CONDITION_TYPES, 'BOOLEAN'] as const;
export const FILTER_CONDITION_TYPES = [...SHARED_CONDITION_TYPES, 'FILTER_EXPRESSION'] as const;
type ConditionTypes = typeof VALIDATION_CONDITION_TYPES | typeof FILTER_CONDITION_TYPES;
export type ConditionType = ConditionTypes[number];

const VALUES_DESCRIPTION = 'The values the type compares against, in order: the list for ONE_OF_LIST, one bound for NUMBER_GREATER, two for NUMBER_BETWEEN, a "Tab!A1:A9" reference for ONE_OF_RANGE, a formula for CUSTOM_FORMULA, none for BLANK / TEXT_IS_EMAIL';

export const validationConditionSchema = z.object({
  type: z.enum(VALIDATION_CONDITION_TYPES).describe('Google ConditionType, e.g. ONE_OF_LIST, ONE_OF_RANGE, NUMBER_BETWEEN, TEXT_IS_EMAIL, DATE_IS_VALID, CUSTOM_FORMULA, BOOLEAN (checkbox)'),
  values: z.array(z.string()).optional().describe(`${VALUES_DESCRIPTION} / BOOLEAN`),
}).describe('A test a cell value must pass');

export const filterConditionSchema = z.object({
  type: z.enum(FILTER_CONDITION_TYPES).describe('Google ConditionType, e.g. NUMBER_GREATER, TEXT_CONTAINS, DATE_AFTER, NOT_BLANK, CUSTOM_FORMULA'),
  values: z.array(z.string()).optional().describe(VALUES_DESCRIPTION),
}).describe('A test rows must pass to stay visible');

export type ValidationCondition = z.infer<typeof validationConditionSchema>;
export type FilterCondition = z.infer<typeof filterConditionSchema>;

export interface BooleanCondition {
  type: ConditionType;
  values?: Array<{ userEnteredValue: string }>;
}

/**
 * Handlers receive input the MCP server already validated, but each enum is
 * checked here too so a direct caller cannot send Google a type it answers
 * with a bare 400.
 */
function buildCondition(input: { type: string; values?: string[] }, allowed: ConditionTypes, paramName: string): BooleanCondition {
  if (!(allowed as readonly string[]).includes(input.type)) {
    throw new Error(`${paramName}.type must be one of ${allowed.join(', ')}, got ${JSON.stringify(input.type)}`);
  }
  const type = input.type as ConditionType;
  // A ONE_OF_RANGE source is a formula to Google, so "Tab!A1:A9" is accepted
  // and the "=" supplied here rather than rejected
  const values = (input.values ?? []).map((value) => ({
    userEnteredValue: type === 'ONE_OF_RANGE' && !value.startsWith('=') ? `=${value}` : value,
  }));
  return values.length > 0 ? { type, values } : { type };
}

export function buildValidationCondition(input: ValidationCondition): BooleanCondition {
  return buildCondition(input, VALIDATION_CONDITION_TYPES, 'condition');
}

export function buildFilterCondition(input: FilterCondition, paramName: string): BooleanCondition {
  return buildCondition(input, FILTER_CONDITION_TYPES, paramName);
}

export const sortBySchema = z.array(z.object({
  column: z.string().describe('Column letter, e.g. "B"'),
  descending: z.boolean().optional().describe('Sort this key descending (default ascending)'),
})).min(1).describe('Sort keys, primary first');

export type SortKey = z.infer<typeof sortBySchema>[number];

/** Columns the keys may name: 0-based, end exclusive. Absent means any. */
export interface ColumnBounds {
  startColumnIndex: number;
  endColumnIndex: number;
}

function columnIndexWithin(column: string, bounds: ColumnBounds | undefined, paramName: string, describe: string): number {
  if (typeof column !== 'string' || !/^[A-Za-z]+$/.test(column)) {
    throw new Error(`${paramName} column must be a letter like "B", got ${JSON.stringify(column)}`);
  }
  const index = columnLetterToIndex(column);
  if (bounds && (index < bounds.startColumnIndex || index >= bounds.endColumnIndex)) {
    throw new Error(`${paramName} column ${column.toUpperCase()} is outside the range ${describe}`);
  }
  return index;
}

export function buildSortSpecs(
  keys: SortKey[],
  bounds: ColumnBounds | undefined,
  describe: string,
): Array<{ dimensionIndex: number; sortOrder: 'ASCENDING' | 'DESCENDING' }> {
  return keys.map(({ column, descending }) => ({
    dimensionIndex: columnIndexWithin(column, bounds, 'sort_by', describe),
    sortOrder: descending ? 'DESCENDING' : 'ASCENDING',
  }));
}

export const filterCriteriaSchema = z.array(z.object({
  column: z.string().describe('Column letter the criterion applies to, e.g. "C"'),
  hidden_values: z.array(z.string()).optional().describe('Cell values to hide, as displayed'),
  condition: filterConditionSchema.optional().describe('Show only rows whose cell passes this test'),
})).min(1).describe('One entry per filtered column');

export type FilterCriterion = z.infer<typeof filterCriteriaSchema>[number];

export function buildFilterSpecs(
  criteria: FilterCriterion[],
  bounds: ColumnBounds | undefined,
  describe: string,
): Array<{ columnIndex: number; filterCriteria: { hiddenValues?: string[]; condition?: BooleanCondition } }> {
  return criteria.map(({ column, hidden_values, condition }, i) => {
    if (hidden_values === undefined && condition === undefined) {
      throw new Error(`criteria[${i}] for column ${String(column).toUpperCase()} needs hidden_values and/or a condition`);
    }
    return {
      columnIndex: columnIndexWithin(column, bounds, 'criteria', describe),
      filterCriteria: {
        ...(hidden_values !== undefined && { hiddenValues: hidden_values }),
        ...(condition !== undefined && { condition: buildFilterCondition(condition, `criteria[${i}].condition`) }),
      },
    };
  });
}
