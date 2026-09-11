/**
 * Decodes the native Sheets grid payload into flat cells. Owns only the
 * Google-specific shape; the budget, cell type, and limits are shared with
 * the .xlsx decoder via sheetBudget.ts so the two branches of get_sheet_data
 * cannot drift apart.
 */

import {
  admitCell,
  createBudget,
  gridDimensions,
  makeCell,
  type Budget,
  type BudgetLimits,
  type Cell,
  type CellType,
  type Hyperlink,
} from './sheetBudget.js';

interface RawCell {
  userEnteredValue?: { formulaValue?: string };
  effectiveValue?: {
    stringValue?: string;
    numberValue?: number;
    boolValue?: boolean;
    errorValue?: { type?: string; message?: string };
  };
  formattedValue?: string;
  hyperlink?: string;
  textFormatRuns?: Array<{
    startIndex?: number;
    format?: { link?: { uri?: string } };
  }>;
}

export interface RawRow {
  values?: RawCell[];
}

export interface DecodeResult {
  data: Cell[][];
  rowCount: number;
  columnCount: number;
  truncated: boolean;
  /**
   * True when the final emitted row is incomplete. Only possible when a
   * single row alone blows the character budget, since otherwise partial
   * rows are dropped to keep row-aligned paging gap-free.
   */
  partialRow: boolean;
}

function resultOf(raw: RawCell): { display: string; type: CellType } {
  const ev = raw.effectiveValue;
  if (!ev) return { display: '', type: 'empty' };
  if (ev.errorValue) return { display: raw.formattedValue || '#ERROR!', type: 'error' };
  if (ev.numberValue !== undefined) {
    return { display: raw.formattedValue || String(ev.numberValue), type: 'number' };
  }
  if (ev.boolValue !== undefined) {
    return { display: raw.formattedValue || String(ev.boolValue), type: 'boolean' };
  }
  return { display: raw.formattedValue || ev.stringValue || '', type: 'string' };
}

function linksFor(raw: RawCell, displayLength: number): Hyperlink[] | undefined {
  const runs = raw.textFormatRuns;
  if (runs && runs.length > 0) {
    const links: Hyperlink[] = [];
    for (let i = 0; i < runs.length; i++) {
      const url = runs[i].format?.link?.uri;
      if (!url) continue;
      const start = runs[i].startIndex || 0;
      const end = i + 1 < runs.length ? (runs[i + 1].startIndex || displayLength) : displayLength;
      links.push({ url, start, end });
    }
    return links;
  }
  return raw.hyperlink ? [{ url: raw.hyperlink, start: 0, end: displayLength }] : undefined;
}

function decodeCell(raw: RawCell, budget: Budget): Cell {
  const { display, type } = resultOf(raw);
  return makeCell({
    display,
    type,
    formula: raw.userEnteredValue?.formulaValue,
    links: linksFor(raw, display.length),
  }, budget);
}

export function decodeGrid(rowData: RawRow[], limits: BudgetLimits = {}): DecodeResult {
  const budget = createBudget(limits);
  const data: Cell[][] = [];
  let truncated = false;
  let partialRow = false;

  outer: for (const rawRow of rowData) {
    const row: Cell[] = [];
    for (const rawCell of rawRow.values || []) {
      const cell = decodeCell(rawCell, budget);
      if (!admitCell(budget, cell)) {
        truncated = true;
        // Drop the partial row so `data` always ends on a complete row and a
        // caller resuming at the next row strands nothing. The exception is a
        // single row that alone exceeds the budget: emit it rather than
        // returning nothing and looping forever on the same row.
        if (row.length > 0 && data.length === 0) {
          data.push(row);
          partialRow = true;
        }
        break outer;
      }
      row.push(cell);
    }
    data.push(row);
  }

  return { data, ...gridDimensions(data), truncated, partialRow };
}
