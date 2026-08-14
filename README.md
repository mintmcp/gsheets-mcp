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
  Reads are capped at 50k cells / 4M characters / 1000 tabs and 20MB of file,
  and a clipped response carries `truncated: true` with a `message` saying
  why. The seven write tools refuse an `.xlsx` with a message pointing at
  `convert_to_google_sheet`, and `copy_spreadsheet` refuses up front, since
  copying one only yields another read-only Excel file.
- **Converting is Drive-side and lossless.** `convert_to_google_sheet` copies
  the upload with `mimeType: application/vnd.google-apps.spreadsheet`, so
  number formats, formulas, hyperlinks and every tab survive — far better
  than re-typing the data into a new sheet. It creates a new file and leaves
  the original `.xlsx` untouched. Legacy `.xls` cannot be read at all; the
  error says to re-save it as a Google Sheet.

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
