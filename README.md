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
- `https://www.googleapis.com/auth/spreadsheets` — for everything else, including
  every chart tool

Charts need no additional scope. `spreadsheets.batchUpdate` accepts any one of
`drive`, `drive.file` or `spreadsheets`, and the chart tools pass the
`spreadsheets` scope the write tools already use.

Both Drive scopes are required for the `.xlsx` path: `get_metadata` and
`get_sheet_data` fall back to Drive when the id turns out to be an Excel
upload, so a spreadsheets-only token fails there with a 403.

## Tools

Eighteen tools, grouped by purpose:

| Category | Tools |
| --- | --- |
| Discovery | `search_spreadsheets`, `get_metadata` |
| Read | `get_sheet_data` |
| Write cells / ranges | `update_cell`, `update_range`, `insert_rows` |
| Format | `format_cells` |
| Clear | `clear_values`, `clear_formatting` |
| Structure | `create_spreadsheet`, `add_sheet`, `copy_spreadsheet` |
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
  workbook is truncated with no way to reach the rest. The seven write tools and
  all five chart tools refuse an `.xlsx` with a message pointing at
  `convert_to_google_sheet` — an uploaded workbook has no chart surface here, so
  even `list_charts` refuses one — and `copy_spreadsheet` refuses up front, since
  copying one only yields another read-only Excel file.
- **Converting is Drive-side and lossless.** `convert_to_google_sheet` copies
  the upload with `mimeType: application/vnd.google-apps.spreadsheet`, so
  number formats, formulas, hyperlinks and every tab survive — far better
  than re-typing the data into a new sheet. It creates a new file and leaves
  the original `.xlsx` untouched. Legacy `.xls` cannot be read at all; the
  error says to re-save it as a Google Sheet.

## Charts

`add_chart` builds COLUMN, BAR, LINE, AREA, SCATTER, STEPPED_AREA and PIE
charts. Source data is given in the same bounded bare A1 notation the rest of
the connector uses:

```jsonc
// add_chart
{ "spreadsheet_id": "...", "sheet_name": "Sales", "chart_type": "COLUMN",
  "domain_range": "A1:A8", "series_ranges": ["B1:B8", "C1:C8"],
  "title": "Q1 Sales", "axis_titles": { "bottom": "Model", "left": "Units" },
  "anchor_cell": "F2" }
```

- **One range per series.** Each of `domain_range` and `series_ranges` must be
  a single column (`"B1:B20"`) or a single row (`"B1:T1"`), all sharing an
  orientation and the same length. A rectangle like `"B2:D9"` is rejected
  rather than reinterpreted — Google accepts it, but what it plots is not what
  the caller meant. Alignment is checked with A1 arithmetic before any request
  is sent, so a mismatch is reported in the caller's own terms.
- **Placement is required.** Pass `anchor_cell` to overlay the chart on a tab,
  or `new_sheet: true` to give it a dedicated chart sheet. There is no default:
  anchoring somewhere arbitrary would cover the data being charted.
- **`header_count` defaults to 1**, so the first cell of each range names its
  series rather than being plotted. Pass `0` for ranges with no header.
- **Chart ids come back from `add_chart`**, and from `list_charts` otherwise.
  `update_chart`, `move_chart` and `delete_chart` all need one.
- **BAR charts plot against the bottom axis.** A bar chart runs horizontally,
  so Google refuses a bar series targeting anything else. The connector sets
  the value axis from the chart type, and converting to or from BAR with
  `update_chart` moves the series across.
- **Charts cannot be rendered back.** The Sheets API has no chart-image
  export, and Drive renders Sheets only to PDF, so nothing can show a caller
  the finished chart. `add_chart` compensates by reading the first cell of
  every source range before it writes: it refuses a set of ranges that are all
  empty, and returns a `resolved` block reporting the stored ranges and the
  label found in each. That catches the failures that actually happen — a
  range pointing at the wrong column, or at a tab holding nothing.

### Updating replaces the whole spec

`updateChartSpec` carries no field mask: whatever is sent replaces the chart's
specification entirely. `update_chart` therefore reads the current spec,
applies the change and writes the result back. Every top-level field except the
chart-type union is carried across untouched, so styling set in the Sheets UI
that these tools cannot express — title formatting, alt text, hidden-dimension
strategy — survives an update, but **a concurrent edit
made between the read and the write is overwritten** — the Sheets API offers
no ETag or `If-Match`, so last write wins. The response says so explicitly.

Three merges are not plain overwrites. Axes merge by `position`, since
BOTTOM/LEFT/RIGHT identify an axis and its index does not. Source ranges
replace wholesale, so `domain_range` and `series_ranges` must be changed
together — a new domain against the old series would be misaligned. And
converting to or from BAR swaps the horizontal and vertical axis titles,
because a bar chart transposes the plot: categories run up the left and
values along the bottom, the reverse of every other basic type. Carrying the
titles across unswapped would leave each one labelling the other axis's data.

A bar chart has no right-hand axis. Google does not reject a bar spec that
carries one — it answers 200 and then simply does not persist the axis — so
passing `axis_titles.right` with BAR is refused here rather than passed through,
which would report success for a title that never exists.

Retyping within the basic family — COLUMN to LINE, LINE to AREA — carries the
whole `basicChart` across, including axis titles, per-series styling and options
this connector does not model (`compareMode` and the rest). Only a switch
between pie and the rest is treated as a change of chart shape, and a COMBO
chart built in the Sheets UI counts as a basic chart here even though these
tools cannot create one.

What the new type cannot carry is then dropped, because Google refuses these
rather than ignoring them:

- **Stacking** applies to COLUMN, BAR, AREA and STEPPED_AREA only. An inherited
  `stackedType` is dropped when retyping to LINE or SCATTER; passing one
  explicitly to those types is refused up front.
- A **right-hand axis title** is dropped when converting to BAR; passing
  `axis_titles.right` explicitly to a BAR chart is refused.
- **3D** applies to PIE and BAR only — inherited elsewhere it is dropped,
  passed explicitly it is refused. COLUMN is the surprise: Google rejects a 3D
  column chart outright, so converting a 3D bar to COLUMN drops the 3D.
- `lineSmoothing` survives only on LINE, and `interpolateNulls` only on LINE and
  AREA, so LINE to AREA drops the first of them.
- Per-series `lineStyle` and `pointStyle` survive only on LINE, AREA and
  SCATTER, and a per-series COMBO `type` never survives a retype.

Two more are dropped by something other than the type change: a
`totalDataLabel` goes when stacking is turned off, and custom per-point labels
go when the source ranges change, since they were pinned to the old data.

Changing `chart_type` between PIE and the others converts the chart in place,
keeping its source data. Use `move_chart` to reposition or resize: it uses
`updateEmbeddedObjectPosition`, which *does* have a field mask, so it needs no
read and leaves the spec untouched.

### Chart sheets outlive their charts

Google does not document what becomes of a chart sheet whose only object goes
away, so `delete_chart` observes rather than asserts: when the chart had its
own sheet it re-reads the tab list afterwards and reports `ownSheetRemoved`.

Tested against the live API, the answer is **no** in both directions:

- Deleting a chart that owns its sheet leaves the sheet behind, empty
  (`ownSheetRemoved: false`).
- Moving a chart off its own sheet with `move_chart` likewise leaves the
  empty chart sheet behind.

Neither tool deletes the leftover sheet — removing a tab is a bigger action
than the one that was asked for, and `get_metadata` now reports `sheetType`,
so an empty `OBJECT` sheet is visible and can be removed deliberately.

### Deleting

The underlying `deleteEmbeddedObject` request removes images and other
embedded objects too, so `delete_chart` first confirms the id belongs to a
chart and refuses otherwise — it cannot delete a non-chart.

`list_charts` omits the raw `ChartSpec` by default, since a styled chart's
spec is large; `include_spec: true` returns it. Charts built outside this
connector — waterfall, treemap, scorecard and the rest — are still listed,
named by their union member, so they remain movable and deletable. They can be
retitled too: `update_chart` with only `title` and/or `subtitle` writes the rest
of the specification back untouched. Any other field needs `chart_type` to
convert the chart to one of the seven types these tools model.

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
| Charts listed by `list_charts` | 200 |
| Series per chart | 50 |
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

A `tools/list` response should enumerate 18 tools.

## Development

```bash
npm run dev          # tsx watch
npm test             # vitest unit tests for pure helpers
npm run smoke        # docker-based end-to-end smoke check
```
