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
- `https://www.googleapis.com/auth/drive.readonly` — for `search_spreadsheets`, `list_comments`, and for reading uploaded `.xlsx` files out of Drive
- `https://www.googleapis.com/auth/drive.file` — for `copy_spreadsheet`, `convert_to_google_sheet`, and folder-scoped `create_spreadsheet`
- `https://www.googleapis.com/auth/spreadsheets` — for everything else, including
  every chart tool
- `https://www.googleapis.com/auth/drive.labels.readonly` — optional, only in
  the `labels` profile; enables label enrichment (see
  [Drive label enrichment](#drive-label-enrichment))

Charts need no additional scope. `spreadsheets.batchUpdate` accepts any one of
`drive`, `drive.file` or `spreadsheets`, and the chart tools pass the
`spreadsheets` scope the write tools already use.

Both Drive scopes are required for the `.xlsx` path: `get_metadata` and
`get_sheet_data` fall back to Drive when the id turns out to be an Excel
upload, so a spreadsheets-only token fails there with a 403.

### Profiles

A **profile** (`PROFILES` in `src/scopes.ts`) is a frozen, named scope set —
a connector's contract with its users:

| Profile    | Scopes                                                                          |
|------------|---------------------------------------------------------------------------------|
| `standard` | `drive.readonly` + `drive.file` + `spreadsheets`                                 |
| `labels`   | `drive.readonly` + `drive.file` + `spreadsheets` + `drive.labels.readonly`       |

Each deployment selects a profile via the `PROFILE` env var; unset registers
every tool (self-hosted default), an unknown or empty value fails at boot.
Both profiles register all thirty-eight tools — the label scope gates the
enrichment below, not a tool.

### Drive label enrichment

On deployments whose grant includes `drive.labels.readonly`, `get_metadata`
and `get_sheet_data` attach the spreadsheet's applied Drive labels as
`_meta.applied` — one entry per label, `{labelId, revisionId, title,
resolved, values[], skippedValueTypes?}`, the same tree the gdrive-mcp
connector emits (the module is vendored from it; keep them identical).
`get_metadata` additionally returns the tree in its visible `labels` field so
agents can answer classification questions; `get_sheet_data` keeps labels in
`_meta` only. `user` fields are withheld, text values cap at 256 chars,
lookups pin the applied label revision, and failures degrade to a
`labelsError` code instead of failing the read. Without the scope no label
call is made and no `_meta` is returned; absence means "surfacing not
enabled", never "no labels". An unset `PROFILE` also enriches, best effort
against whatever the token allows.

## Tools

Thirty-eight tools, grouped by purpose:

| Category | Tools |
| --- | --- |
| Discovery | `search_spreadsheets`, `get_metadata` |
| Read | `get_sheet_data`, `list_comments` |
| Write cells / ranges | `update_cell`, `update_range`, `insert_rows` |
| Format | `format_cells`, `set_borders`, `merge_cells`, `unmerge_cells` |
| Clear | `clear_values`, `clear_formatting` |
| Validation | `set_data_validation`, `clear_data_validation` |
| Named ranges | `add_named_range`, `delete_named_range` (listed by `get_metadata`) |
| Filters | `set_basic_filter`, `clear_basic_filter`, `add_filter_view`, `delete_filter_view` |
| Structure | `create_spreadsheet`, `add_sheet`, `rename_sheet`, `duplicate_sheet`, `delete_sheet`, `copy_spreadsheet` |
| Rows / columns | `insert_rows_or_columns`, `delete_rows_or_columns`, `freeze_panes`, `sort_range` |
| Find / replace | `find_replace` |
| Charts | `add_chart`, `list_charts`, `update_chart`, `move_chart`, `delete_chart` |
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
- **Literal values on request.** `update_cell`, `update_range` and
  `insert_rows` parse values like typed input by default (`USER_ENTERED`:
  `=` starts a formula, `"01"` becomes `1`). Pass
  `value_input_option: "RAW"` to store every string exactly as given.
- **Conditions are Google's.** `set_data_validation` and the filter tools
  take a `condition` of `{type, values}` using the Sheets API
  `ConditionType` names (`ONE_OF_LIST`, `NUMBER_BETWEEN`, `CUSTOM_FORMULA`,
  `BOOLEAN` for a checkbox, ...). A `ONE_OF_RANGE` source may be written
  `"Tab!A1:A9"`; the leading `=` is added. Google validates which types
  each feature accepts and how many values they take.
- **One filter per tab, many views.** `set_basic_filter` replaces the tab's
  filter; `add_filter_view` saves a named view with the same `range`,
  `sort_by` and `criteria` inputs. `get_metadata` lists the basic filter
  range and each tab's filter views, plus the spreadsheet's named ranges.
- **Append-only inserts.** `insert_rows` uses
  `values:append + INSERT_ROWS`; it cannot insert at an arbitrary index or
  add columns. Use `update_range` for mid-sheet writes, and
  `insert_rows_or_columns` to open up empty rows or columns at a position.
- **Rows by number, columns by letter.** `insert_rows_or_columns` and
  `delete_rows_or_columns` take `start_index` as a 1-based row number for
  `ROWS` and a column letter for `COLUMNS`, matching A1 notation; the other
  form is rejected rather than guessed at.
- **Sorting is range-local.** `sort_range` moves only the cells inside
  `range`, so pass every column of the table. `has_header` keeps the first
  row of the range in place.
- **Find/replace has no dry run.** `find_replace` searches every tab unless
  `sheet_name` (optionally with `range`) narrows it, and reports how many
  cells changed. Read the range first when you need to know what will match.
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
  workbook is truncated with no way to reach the rest. The write, structure and
  chart tools all refuse an `.xlsx` with a message pointing at
  `convert_to_google_sheet` — an uploaded workbook has no chart surface here, so
  even `list_charts` refuses one — and `copy_spreadsheet` refuses up front, since
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
call; all 26 columns of the same tab takes 17. Pages are sized by the columns
that hold data, not the allocated grid: a 3-column tab under Google's default
26-column grid pages every 1,666 rows.

Uploaded `.xlsx` tabs report no dimensions — the tab list is read without
parsing the sheets, so the counts are not available there.

## Response limits

`get_sheet_data` is bounded. It reads the tab's grid dimensions first, then
fetches the tab in chunks: the first sized to fit the cell cap at the
allocated width, the rest at the width the data turned out to have, so an
oversized tab is never fetched in the first place and a narrow one is not
paged as if it were wide.

| Limit | Value |
| --- | --- |
| Cells per response | 5,000 |
| Cells per upstream fetch (at the allocated width) | 50,000 |
| Characters per response | 100,000 |
| Characters per cell | 32,768 |
| Columns per response | 256 |
| Tabs listed by `get_metadata` | 200 |
| Charts listed by `list_charts` | 200 |
| Series per chart | 50 |
| Cells per write | 50,000 |
| Upstream response bytes | 25 MB |
| .xlsx file size | 10 MB |

The two caps do different jobs. The cell cap bounds what is *fetched*: it
sizes the page, so 5,000 ÷ 26 used columns is a 192-row page and 5,000 ÷ 3
is 1,666. The
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

Truncation is reported exactly, not guessed. The tab's *allocated* grid is
usually larger than the used range, so each chunk asks for one row beyond
its end: that row coming back is proof the data continues. A chunk that
comes back short is settled by looking ahead in growing spans, so a tab with
50,000 allocated rows but 20 rows of data comes back complete after a few
empty answers rather than a read of everything after it.

An explicit `range` narrows the scope but does not lift the caps — an
oversized rectangle is clamped and paged the same way, with `nextRange`
walking through the range you asked for. The rectangle is first intersected
with the tab, so asking for more rows or columns than exist costs nothing and
is not reported as truncation.

Paging bounds each response; it does not make a huge tab cheap to read in
full, since walking every page still moves every cell through the caller.
For a large tab, use `get_metadata` to find the rows and columns you need
and request those directly.

A run of blank rows longer than a chunk ends the page rather than padding
it: data in rows 1-10 and again at 4001+ comes back as rows 1-10 with
`nextRange` starting at row 4001, so paging skips the gap instead of walking
through it.

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

A `tools/list` response should enumerate 38 tools.

## Development

```bash
npm run dev          # tsx watch
npm test             # vitest unit tests for pure helpers
npm run smoke        # docker-based end-to-end smoke check
```
