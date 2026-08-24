# gsheets-mcp

An MCP server that wraps the Google Sheets API (with Drive-side metadata for
discovery) for the MintMCP hosted runtime. Speaks streamable HTTP on
`POST /mcp`, listens on port 8000, and reads the caller's Google OAuth access
token from the per-request `Authorization` header.

## Auth contract

MintMCP handles the Google OAuth flow on the frontend and forwards the user's
access token to this server on every request:

```
Authorization: Bearer <google-access-token>
```

There are no server-side env vars for credentials. The token is parsed in
`src/index.ts`, stashed in `AsyncLocalStorage` for the duration of the
request, and read by each tool handler via `withGoogleAuth`.

### OAuth scopes (configured at the connector level)

- `openid`
- `https://www.googleapis.com/auth/userinfo.email`
- `https://www.googleapis.com/auth/userinfo.profile`
- `https://www.googleapis.com/auth/drive.readonly` — for `search_spreadsheets`, and for reading uploaded `.xlsx` files out of Drive
- `https://www.googleapis.com/auth/drive.file` — for `copy_spreadsheet`, `convert_to_google_sheet`, and folder-scoped `create_spreadsheet`
- `https://www.googleapis.com/auth/spreadsheets` — for everything else

Both Drive scopes are required for the `.xlsx` path: `get_metadata` and
`get_sheet_data` fall back to Drive when the id turns out to be an Excel
upload, so a spreadsheets-only token fails there with a 403.

## Tools

Thirteen tools, grouped by purpose:

| Category | Tools |
| --- | --- |
| Discovery | `search_spreadsheets`, `get_metadata` |
| Read | `get_sheet_data` |
| Write cells / ranges | `update_cell`, `update_range`, `insert_rows` |
| Format | `format_cells` |
| Clear | `clear_values`, `clear_formatting` |
| Structure | `create_spreadsheet`, `add_sheet`, `copy_spreadsheet` |
| Excel uploads | `convert_to_google_sheet` |

Notable behaviors:

- **Bounded A1 only.** `update_range`, `format_cells`, `clear_values`, and
  `clear_formatting` reject sheet-qualified strings (`"Sheet1!A1:C3"`) and
  open-ended forms (`"A:C"`, `"1:3"`, `"A1:C"`). Pass the tab via the
  `sheet_name` argument and an `"A1"` or `"A1:C3"` style range.
- **Hex colors.** `format_cells` accepts `"#FF0000"` or `"#F00"` for any
  color field, in addition to the Sheets API's native
  `{red, green, blue}` float-0..1 object.
- **`verticalAlignment`.** `format_cells` exposes `TOP | MIDDLE | BOTTOM`
  alongside `horizontalAlignment`.
- **Append-only inserts.** `insert_rows` uses
  `values:append + INSERT_ROWS`; it cannot insert at an arbitrary index or
  add columns. Use `update_range` for mid-sheet writes.
- **Uploaded `.xlsx` files are readable, not editable.** The Sheets API
  refuses Excel uploads with a 400 "must not be an Office file". Rather than
  surface that, `get_metadata` and `get_sheet_data` catch it and re-read the
  file's bytes from Drive via SheetJS, so an `.xlsx` id reads like a native
  Sheet. `search_spreadsheets` finds them too, and all three report
  `kind: "native" | "xlsx"` so a caller can tell the difference.
  Reads are bounded by the same caps as a native sheet — see Response limits
  below — plus 1000 tabs and 7MB of file, and a clipped response carries
  `truncated: true` with a `message` saying why. Unlike a native sheet an
  `.xlsx` cannot be paged: there is no `range` or `nextRange`, so an oversized
  workbook is truncated with no way to reach the rest. The seven write tools refuse an `.xlsx` with a message pointing at
  `convert_to_google_sheet`, and `copy_spreadsheet` refuses up front, since
  copying one only yields another read-only Excel file.
- **Converting is Drive-side and lossless.** `convert_to_google_sheet` copies
  the upload with `mimeType: application/vnd.google-apps.spreadsheet`, so
  number formats, formulas, hyperlinks and every tab survive — far better
  than re-typing the data into a new sheet. It creates a new file and leaves
  the original `.xlsx` untouched. Legacy `.xls` cannot be read at all; the
  error says to re-save it as a Google Sheet.

## Reading a sheet efficiently

Call `get_metadata` first. It returns each tab's `rowCount` and `columnCount`
at no extra API cost, so you can ask for the range you actually want instead
of a blind read that comes back truncated. Pass `include_headers: true` to
also get row 1 of each tab (one extra batched call, first 50 tabs):

```jsonc
// get_metadata { "spreadsheet_id": "...", "include_headers": true }
{ "sheets": [
    { "title": "Sales", "rowCount": 3002, "columnCount": 28,
      "headers": ["date", "region", "amount", "..."] }
] }

// then read only what you need
// get_sheet_data { "spreadsheet_id": "...", "range": "A1:C200" }
```

**`rowCount` and `columnCount` are the allocated grid, not the extent of the
data.** Google allocates a default grid, so the tab above reports 3002 x 28
while holding 3000 rows of 26 columns, and a tab with 40 rows of data commonly
reports 1000. Use them to size a request, not to tell a user how big the data
is. Where the data actually ends is settled by reading: a response without
`nextRange` has reached it.

Narrowing columns is usually the bigger win, since the cap counts cells rather
than rows. One column of a 3000-row tab is 3000 cells and fits in a single
call; all 26 columns of the same tab takes 17.

Uploaded `.xlsx` tabs report no dimensions — the tab list is read without
parsing the sheets, so the counts are not available there.

## Response limits

`get_sheet_data` is bounded. It reads the tab's grid dimensions first, then
requests only an A1 window sized to fit the cell cap, so an oversized tab is
never fetched in the first place.

| Limit | Value |
| --- | --- |
| Cells per response | 5,000 |
| Characters per response | 100,000 |
| Characters per cell | 32,768 |
| Columns per response | 256 |
| Tabs listed by `get_metadata` | 200 |
| Cells per write | 50,000 |
| Upstream response bytes | 25 MB |
| .xlsx file size | 10 MB |

The two caps do different jobs. The cell cap bounds what is *fetched*: it
sizes the A1 window, so 5,000 ÷ 26 columns is a 192-row request. The
character cap bounds what is *returned*, stopping one pathological tab from
serializing to a huge payload — 5,000 cells at the per-cell limit would
otherwise be megabytes.

Both are connector bounds. Clients impose their own: Claude Code refuses a
tool result over 25,000 tokens, about 48,000 characters of this JSON. If a
client refuses a response, pass a narrower `range` or raise the client's own
limit; the connector does not shrink itself to the strictest client.

A cell count cannot bound size on its own, since 5,000 cells is 85KB of short
codes or 350KB of prose. So rows per page float with content density, and in
practice the character cap is what binds on all but the shortest values:

| Average cell | Rows returned (26 columns) |
| --- | --- |
| 4 characters | 192 (the cell cap binds first) |
| 20 characters | 115 |
| 200 characters | 18 |

Every one of those carries `truncated: true` and a `nextRange`, so a caller
pages through them the same way regardless of which budget tripped.

`returnedRange` reports the range actually present in `data`, which can be
smaller than the window when a budget trips. When more data remains the
response also carries `truncated: true` and a `nextRange`; pass that value
back as `range` to read the next window:

```jsonc
// get_sheet_data { "spreadsheet_id": "..." }  — tab has 500,000 rows
{
  "returnedRange": "A1:Z192",
  "truncated": true,
  "nextRange": "A193:Z500000",
  "message": "Rows 1-192 returned; pass nextRange as `range` to continue."
}
```

`nextRange` is the remaining *scope*, not the next window: pass it straight
back and it is clamped to a window again, so repeating that until
`nextRange` is absent walks the whole tab with no gaps or overlap.

Truncation is reported exactly, not guessed. The window is sized from the
tab's *allocated* grid, which is usually larger than the used range, so each
read asks for one row beyond the window: if that probe row comes back empty
the data genuinely ended inside the window and `truncated` is absent. A tab
with 50,000 allocated rows but 20 rows of data comes back complete.

An explicit `range` narrows the scope but does not lift the caps — an
oversized rectangle is clamped and paged the same way, with `nextRange`
walking through the range you asked for. The rectangle is first intersected
with the tab, so asking for more rows or columns than exist costs nothing and
is not reported as truncation.

Paging bounds each response; it does not make a huge tab cheap to read in
full, since walking every page still moves every cell through the caller.
For a large tab, use `get_metadata` to find the rows and columns you need
and request those directly.

**Known limitation:** truncation is detected by asking for one row beyond the
window, so a tab with a gap larger than one window (say data in rows 1-10 and
again at 4001+) stops at the first block. Read past a gap with an explicit
`range`.

Write tools reject a matrix over 50,000 cells rather than expanding it in
memory, and `update_cell` accepts at most 1,000 content segments.

## Build and run

```bash
npm install
npm run build
PORT=8000 npm start
```

### Docker (production)

Build for `linux/amd64` (MintMCP's hosts are amd64):

```bash
docker buildx create --name mintmcp-builder --driver docker-container --use   # one-time
docker buildx build \
  --builder mintmcp-builder \
  --platform linux/amd64 \
  -t mintmcp/gsheets-mcp:latest \
  --push .
```

### Deploy via hosted-cli

```bash
hosted-cli build-and-deploy \
  --image mintmcp/gsheets-mcp:latest \
  --connector gsheets-mcp
```

## Verifying a running instance

```bash
curl -s http://localhost:8000/healthz

curl -s -X POST http://localhost:8000/mcp \
  -H "Content-Type: application/json" \
  -H "Accept: application/json, text/event-stream" \
  -H "Authorization: Bearer fake-token" \
  -d '{"jsonrpc":"2.0","id":1,"method":"tools/list","params":{}}'
```

A `tools/list` response should enumerate 13 tools.

## Development

```bash
npm run dev          # tsx watch
npm test             # vitest unit tests for pure helpers
npm run smoke        # docker-based end-to-end smoke check
```
