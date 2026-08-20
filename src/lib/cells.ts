/**
 * Decodes the native Sheets grid payload into flat cells. Owns only the
 * Google-specific shape; the budget, cell type, and limits are shared with
 * the .xlsx decoder via sheetBudget.ts so the two branches of get_sheet_data
 * cannot drift apart.
 */

import {
  boundLinks,
  admitCell,
  clipValue,
  createBudget,
  gridDimensions,
  safeLinkUrl,
  type Budget,
  type BudgetLimits,
  type Cell,
  type CellType,
  type Hyperlink,
} from './sheetBudget.js';

interface RawCell {
  userEnteredValue?: {
    stringValue?: string;
    numberValue?: number;
    boolValue?: boolean;
    formulaValue?: string;
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

function classify(raw: RawCell): { value: string; type: CellType } {
  const uev = raw.userEnteredValue;
  if (!uev) return { value: '', type: 'empty' };
  if (uev.formulaValue !== undefined) return { value: uev.formulaValue, type: 'formula' };
  if (uev.numberValue !== undefined) {
    return { value: raw.formattedValue || String(uev.numberValue), type: 'number' };
  }
  if (uev.boolValue !== undefined) {
    return { value: raw.formattedValue || String(uev.boolValue), type: 'boolean' };
  }
  return { value: raw.formattedValue || uev.stringValue || '', type: 'string' };
}

function linksFor(
  raw: RawCell,
  displayLength: number,
  budget: Budget,
): Hyperlink[] | undefined {
  const runs = raw.textFormatRuns;
  if (runs && runs.length > 0) {
    const links: Hyperlink[] = [];
    for (let i = 0; i < runs.length; i++) {
      const url = safeLinkUrl(runs[i].format?.link?.uri, budget);
      if (!url) continue;
      const start = runs[i].startIndex || 0;
      const end = i + 1 < runs.length ? (runs[i + 1].startIndex || displayLength) : displayLength;
      links.push({ url, start, end });
    }
    return boundLinks(links, displayLength);
  }
  const whole = safeLinkUrl(raw.hyperlink, budget);
  if (whole) {
    return boundLinks([{ url: whole, start: 0, end: displayLength }], displayLength);
  }
  return undefined;
}

function decodeCell(raw: RawCell, budget: Budget): Cell {
  const { value, type } = classify(raw);
  const clipped = clipValue(value, budget);
  const cell: Cell = { value: clipped };
  if (type !== 'string') cell.type = type;
  if (clipped.length < value.length) cell.valueShortened = true;

  const display = raw.formattedValue || clipped;
  const links = linksFor(raw, Math.min(display.length, budget.maxCellChars), budget);
  if (links) cell.hyperlinks = links;

  return cell;
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
