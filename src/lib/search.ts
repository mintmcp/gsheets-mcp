/**
 * Drive `q` parameter builders.
 *
 * Drive's `q` syntax requires backslash-escaping of `\` and `'` inside
 * single-quoted string literals. Control characters either break the
 * parser or surprise the URL encoding pass, so we reject them outright.
 */

const MIME_SPREADSHEET = "application/vnd.google-apps.spreadsheet";
const MIME_XLSX =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
// C0 control range, excluding TAB (0x09).
const CONTROL_CHAR_RE = /[\x00-\x08\x0A-\x1F]/;

/**
 * Build the `q` string used by `search_spreadsheets` to filter Drive files.
 * Covers native Sheets and uploaded .xlsx, which the read tools can also
 * handle; the caller reports which is which via the `kind` field.
 * If `name` is omitted, undefined, or blank after trimming, no
 * `name contains` clause is emitted.
 */
export function buildDriveSearchQuery(name?: unknown): string {
  let q = `(mimeType = '${MIME_SPREADSHEET}' or mimeType = '${MIME_XLSX}')`;
  if (typeof name === 'string') {
    const trimmed = name.trim();
    if (trimmed.length > 0) {
      if (CONTROL_CHAR_RE.test(trimmed)) {
        throw new Error('Search name must not contain control characters');
      }
      // Escape backslashes BEFORE single quotes so the inserted backslash
      // from quote-escaping isn't itself escaped again.
      const safeName = trimmed.replace(/\\/g, '\\\\').replace(/'/g, "\\'");
      q += ` and name contains '${safeName}'`;
    }
  }
  q += ` and trashed = false`;
  return q;
}
