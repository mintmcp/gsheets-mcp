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
- `https://www.googleapis.com/auth/drive.readonly` — for `search_spreadsheets` and `get_sheet_data`'s Drive-label read (`files.listLabels`)
- `https://www.googleapis.com/auth/drive.labels.readonly` — for `get_sheet_data`'s Drive-label enrichment (resolves selection-choice ids to names). If absent, label reads fail soft: results carry `labelsError` instead of `labels`.
- `https://www.googleapis.com/auth/drive.file` — for `copy_spreadsheet` and folder-scoped `create_spreadsheet`
- `https://www.googleapis.com/auth/spreadsheets` — for everything else

## Tools

Twelve tools, grouped by purpose:

| Category | Tools |
| --- | --- |
| Discovery | `search_spreadsheets`, `get_metadata` |
| Read | `get_sheet_data` |
| Write cells / ranges | `update_cell`, `update_range`, `insert_rows` |
| Format | `format_cells` |
| Clear | `clear_values`, `clear_formatting` |
| Structure | `create_spreadsheet`, `add_sheet`, `copy_spreadsheet` |

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

## Build and run

```bash
npm install
npm run build
PORT=8000 npm start
```

### Build the image locally

Build the image straight from the repo's `Dockerfile` (from source, on the
current branch) instead of pulling a published tag — handy for testing a
branch or verifying a build. Build for `linux/amd64` to match the MintMCP
runtime (required on Apple Silicon):

```bash
docker build --platform linux/amd64 -t gsheets-mcp:local .
docker run --rm -p 8000:8000 gsheets-mcp:local
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

A `tools/list` response should enumerate 12 tools.

## Development

```bash
npm run dev          # tsx watch
npm test             # vitest unit tests for pure helpers
npm run smoke        # docker-based end-to-end smoke check
```
