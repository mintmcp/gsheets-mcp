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
 * Bounds what one response serializes to. This guards the CALLER'S SESSION,
 * not this process: 250,000 characters peaks around 1MB against a ~512MB
 * heap, so it is nowhere near a memory limit. The connector's own memory
 * guards are MAX_RESPONSE_BYTES, MAX_XLSX_BYTES and MAX_PADDED_CELLS.
 *
 * It also earns its place by producing the halt point that `nextRange` is
 * derived from. Without a budget there is no truncation, so there is nothing
 * to resume from and paging stops working. A caller cannot size the request
 * itself either: a 45-row by 8-column tab measured 296 cells and 98,799
 * characters, which nothing in the grid dimensions would have predicted.
 *
 * Measured at 4.58 characters per token on real sheet JSON, so this is
 * roughly 55,000 tokens, and 75,000 on a sheet of short strings. Past about
 * this point MAX_CELLS starts binding first on wide tabs, since the window
 * holds only floor(5000/columns) rows. Clients set their own, stricter
 * ceilings, but pinning the connector to the strictest one would shortchange
 * every other caller.
 */
export const MAX_OUTPUT_CHARS = 250_000;

/**
 * Google's own ceiling on one cell, so a native cell is never clipped and
 * `valueShortened` never fires on that path. This is a backstop for sources
 * Google does not police: SheetJS puts no limit on a decoded string, and the
 * .xlsx it came from is a 7MB zip that decompresses much further.
 *
 * MAX_OUTPUT_CHARS cannot cover this. `admitCell` admits the first cell of a
 * page unconditionally, since rejecting it would return an empty page and
 * leave a paging caller stuck on the same row, so cell one is bounded here or
 * nowhere. `safeLinkUrl` bounds hyperlink targets against the same value.
 */
export const MAX_CELL_CHARS = 50_000;

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

/** Cost of `,"formula":""` around the formula text itself. */
const FORMULA_FIELD_CHARS = 13;

/** Rough serialized cost of one `{"url":"...","start":N,"end":N}` entry. */
const HYPERLINK_ENVELOPE_CHARS = 34;

export type CellType = 'string' | 'number' | 'boolean' | 'error' | 'empty';

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
  formula?: string;
  hyperlinks?: Hyperlink[];
  /**
   * Set when the value was clipped at maxCellChars. Without it a clipped
   * value reads as complete, so a model can quote a half sentence as whole.
   * Mirrors `nameShortened` on .xlsx tab names.
   */
  valueShortened?: true;
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
function clipValue(value: string, budget: Budget): string {
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
function safeLinkUrl(url: string | undefined, budget: Budget): string | undefined {
  if (!url || !SAFE_LINK_SCHEME.test(url)) return undefined;
  return url.length > budget.maxCellChars ? url.slice(0, budget.maxCellChars) : url;
}

function boundLinks(
  links: Hyperlink[],
  displayLength: number,
  budget: Budget,
): Hyperlink[] | undefined {
  const bounded: Hyperlink[] = [];
  for (const link of links) {
    const url = safeLinkUrl(link.url, budget);
    const start = Math.min(link.start, displayLength);
    const end = Math.min(link.end, displayLength);
    if (url && end > start) bounded.push({ url, start, end });
  }
  return bounded.length > 0 ? bounded : undefined;
}

export interface CellParts {
  display: string;
  type: CellType;
  formula?: string;
  /** Unfiltered, offsets relative to `display` */
  links?: Hyperlink[];
}

export function makeCell(parts: CellParts, budget: Budget): Cell {
  const cell: Cell = { value: clipValue(parts.display, budget) };
  if (parts.type !== 'string') cell.type = parts.type;
  if (cell.value.length < parts.display.length) cell.valueShortened = true;
  // A truncated formula written back would corrupt the sheet, so omit it rather than clip it
  if (parts.formula !== undefined && parts.formula.length <= budget.maxCellChars) {
    cell.formula = parts.formula;
  }
  if (parts.links) {
    const links = boundLinks(parts.links, cell.value.length, budget);
    if (links) cell.hyperlinks = links;
  }
  return cell;
}

function cellCost(cell: Cell): number {
  return cell.value.length
    + CELL_ENVELOPE_CHARS
    + (cell.type ? TYPE_FIELD_CHARS : 0)
    + (cell.formula !== undefined ? cell.formula.length + FORMULA_FIELD_CHARS : 0)
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
 * why, or told why without the flag being set.
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
