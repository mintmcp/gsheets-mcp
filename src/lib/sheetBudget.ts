/**
 * The single definition of what a bounded sheet read costs and returns.
 *
 * `get_sheet_data` has two decoders — native Google grids (cells.ts) and
 * uploaded .xlsx workbooks (xlsx.ts) — that iterate very different sources
 * but must produce the same cell shape under the same limits. Everything
 * they share lives here so the two branches of one tool cannot drift: when
 * these were defined twice, the character envelope silently disagreed (32
 * vs 50) and the advertised character cap meant two different things.
 */

/**
 * Sized for the consumer, not just for the process. A cell serializes to
 * roughly 36 characters, so 5,000 cells is ~180KB — large but readable by a
 * model in one go, where the previous 50,000 was ~1.8MB and blew past most
 * context windows. Callers who need more follow `nextRange`.
 */
export const MAX_CELLS = 5_000;

/**
 * Inbound writes are a separate concern from read page size: the request body
 * is already bounded at 10MB, and shrinking batches to the read cap would
 * force callers into needless round trips.
 */
export const MAX_WRITE_CELLS = 50_000;
/**
 * Scaled with MAX_CELLS at the same ~80 chars per cell. Left at its old
 * 4,000,000 it would have become the binding limit for text-heavy sheets and
 * quietly reintroduced megabyte responses through the other door.
 */
export const MAX_OUTPUT_CHARS = 400_000;
export const MAX_CELL_CHARS = 32_768;

/**
 * Rough serialized cost of one cell's JSON envelope — the braces, keys and
 * quotes around its value. Deliberately generous: overestimating shrinks the
 * response, underestimating overshoots the budget it exists to enforce.
 * Sized for `{"value":"..."}`, since `type` is omitted on the string cells
 * that dominate a typical sheet.
 */
export const CELL_ENVELOPE_CHARS = 30;

/** Rough serialized cost of one `{"url":"...","start":N,"end":N}` entry. */
export const HYPERLINK_ENVELOPE_CHARS = 34;

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

export function exhausted(budget: Budget): boolean {
  return budget.cells >= budget.maxCells || budget.chars >= budget.maxChars;
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
 * Normalize a hyperlink target, or return undefined to drop it. Clipping is
 * part of the budget contract: `exhausted()` is checked before a cell is
 * charged, so an unbounded URL would overshoot the character cap by its own
 * length before anything noticed.
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

/** Charge one decoded cell against the budget. */
export function chargeCell(budget: Budget, cell: Cell): void {
  budget.cells++;
  budget.chars += cell.value.length + CELL_ENVELOPE_CHARS
    + (cell.hyperlinks?.reduce(
      (sum, link) => sum + link.url.length + HYPERLINK_ENVELOPE_CHARS,
      0,
    ) ?? 0);
}

/** Charge a row that carries no cells (a skipped leading row in a sparse sheet). */
export function chargeEmptyRow(budget: Budget): void {
  budget.cells++;
  budget.chars += CELL_ENVELOPE_CHARS;
}

export function gridDimensions(data: Cell[][]): { rowCount: number; columnCount: number } {
  let columnCount = 0;
  for (const row of data) {
    if (row.length > columnCount) columnCount = row.length;
  }
  return { rowCount: data.length, columnCount };
}
