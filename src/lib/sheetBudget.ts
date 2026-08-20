/**
 * The single definition of what a bounded sheet read costs and returns.
 *
 * `get_sheet_data` has two decoders — native Google grids (cells.ts) and
 * uploaded .xlsx workbooks (xlsx.ts) — that iterate very different sources
 * but must produce the same cell shape under the same limits. Everything
 * they share lives here so the two branches of one tool cannot drift.
 */

/**
 * Bounds the FETCH: the A1 window is sized to hold at most this many cells,
 * so an oversized tab is never requested in the first place. Callers who need
 * more follow `nextRange`.
 */
export const MAX_CELLS = 5_000;

/**
 * Bounds what one response can serialize to, so a pathological sheet cannot
 * produce a gigantic payload: MAX_CELLS alone permits 5,000 cells of 32,768
 * characters each. Text-heavy tabs truncate here and page through nextRange
 * like any other overflow.
 *
 * Sized for the connector, not for one client. 100,000 characters lets a full
 * 5,000-cell page through at typical content length, which is the read a
 * caller asking for a bounded rectangle actually wants. Clients impose their
 * own, stricter ceilings — Claude Code refuses a tool result over 25,000
 * tokens, roughly 48,000 characters of this JSON — but that is a client-side
 * setting its user can raise, and pinning the connector to the strictest one
 * would shortchange every other caller.
 */
export const MAX_OUTPUT_CHARS = 100_000;
export const MAX_CELL_CHARS = 32_768;

/**
 * Serialized cost of one cell's JSON envelope: `{"value":"..."},` is 14
 * characters around the value itself. Accuracy matters because the character
 * budget is what keeps a page ingestible.
 */
const CELL_ENVELOPE_CHARS = 14;

/**
 * Cost of the `type` field, which is emitted for everything except plain
 * strings. `,"type":"formula"` is the longest at 17. Omitting this charged a
 * sheet of numbers roughly half its real size, so a numeric tab could return
 * about twice MAX_OUTPUT_CHARS.
 */
const TYPE_FIELD_CHARS = 17;

/** Rough serialized cost of one `{"url":"...","start":N,"end":N}` entry. */
const HYPERLINK_ENVELOPE_CHARS = 34;

export type CellType = 'string' | 'number' | 'boolean' | 'formula' | 'empty';

export interface Hyperlink {
  url: string;
  start: number;
  end: number;
}

export interface Cell {
  value: string;
  /**
   * Omitted for plain strings. `string` is the fallback branch of both
   * decoders, so emitting it says nothing while costing 17 bytes on the ~90%
   * of cells that are text. Absent means string.
   */
  type?: CellType;
  hyperlinks?: Hyperlink[];
}

export interface Budget {
  cells: number;
  chars: number;
  maxCells: number;
  maxChars: number;
  maxCellChars: number;
}

export interface BudgetLimits {
  maxCells?: number;
  maxChars?: number;
  maxCellChars?: number;
}

export function createBudget(limits: BudgetLimits = {}): Budget {
  return {
    cells: 0,
    chars: 0,
    maxCells: limits.maxCells ?? MAX_CELLS,
    maxChars: limits.maxChars ?? MAX_OUTPUT_CHARS,
    maxCellChars: limits.maxCellChars ?? MAX_CELL_CHARS,
  };
}

/** Clip a single cell value so one pathological cell cannot blow the budget. */
export function clipValue(value: string, budget: Budget): string {
  return value.length > budget.maxCellChars
    ? value.slice(0, budget.maxCellChars)
    : value;
}

/**
 * Hyperlink targets we are willing to hand to a model. Anything else — most
 * importantly `javascript:` and `data:` — is dropped rather than forwarded.
 * Both decoders must apply this: sheet content is untrusted input regardless
 * of whether it arrived as a native grid or an uploaded workbook.
 */
const SAFE_LINK_SCHEME = /^(https?|mailto):/i;

/**
 * Normalize a hyperlink target, or return undefined to drop it. Clipping
 * bounds what one cell can cost: a URL is charged in full against the
 * character budget, so an unbounded one would consume a page by itself.
 */
export function safeLinkUrl(url: string | undefined, budget: Budget): string | undefined {
  if (!url || !SAFE_LINK_SCHEME.test(url)) return undefined;
  return url.length > budget.maxCellChars ? url.slice(0, budget.maxCellChars) : url;
}

/**
 * Clamp link offsets to the text actually emitted, dropping links that fall
 * entirely past it. Callers pass the display text they will return, so an
 * offset can never point outside `cell.value`.
 */
export function boundLinks(
  links: Hyperlink[],
  displayLength: number,
): Hyperlink[] | undefined {
  const bounded: Hyperlink[] = [];
  for (const link of links) {
    const start = Math.min(link.start, displayLength);
    const end = Math.min(link.end, displayLength);
    if (end > start) bounded.push({ url: link.url, start, end });
  }
  return bounded.length > 0 ? bounded : undefined;
}

function cellCost(cell: Cell): number {
  return cell.value.length
    + CELL_ENVELOPE_CHARS
    + (cell.type ? TYPE_FIELD_CHARS : 0)
    + (cell.hyperlinks?.reduce(
      (sum, link) => sum + link.url.length + HYPERLINK_ENVELOPE_CHARS,
      0,
    ) ?? 0);
}

/**
 * Charge a cell only if it fits, so the budget is a ceiling rather than a
 * threshold crossed on the way out. Testing the budget before decoding let
 * the cell that tripped the limit through, overshooting by as much as one
 * MAX_CELL_CHARS value.
 *
 * The first cell is always admitted: returning an empty page would leave a
 * paging caller looping on the same row forever.
 */
export function admitCell(budget: Budget, cell: Cell): boolean {
  const cost = cellCost(cell);
  if (budget.cells > 0
    && (budget.cells + 1 > budget.maxCells || budget.chars + cost > budget.maxChars)) {
    return false;
  }
  budget.cells++;
  budget.chars += cost;
  return true;
}

/**
 * Charge a row that carries no cells, or refuse it if the budget is spent.
 * Mirrors `admitCell` so there is one way to spend the budget. Only the .xlsx
 * decoder needs it: it
 * walks a whole worksheet, so a workbook with a million leading blank rows
 * would otherwise cost nothing and iterate forever. The native decoder reads
 * an A1 window whose height is already floor(MAX_CELLS / columns), so its row
 * count is bounded before decoding starts and charging blanks there would
 * only shrink sparse single-column reads for no gain.
 */
export function admitEmptyRow(budget: Budget): boolean {
  if (budget.cells + 1 > budget.maxCells
    || budget.chars + CELL_ENVELOPE_CHARS > budget.maxChars) {
    return false;
  }
  budget.cells++;
  budget.chars += CELL_ENVELOPE_CHARS;
  return true;
}

/**
 * The `truncated` / `message` pair every bounded response carries. Built from
 * the notes so a caller cannot be told data is missing without being told
 * why, or told why without the flag being set — four call sites used to
 * maintain that pairing by hand.
 */
export function truncationFields(
  notes: string[],
): { truncated?: true; message?: string } {
  return notes.length > 0 ? { truncated: true, message: notes.join(' ') } : {};
}

export function gridDimensions(data: Cell[][]): { rowCount: number; columnCount: number } {
  let columnCount = 0;
  for (const row of data) {
    if (row.length > columnCount) columnCount = row.length;
  }
  return { rowCount: data.length, columnCount };
}
