#!/usr/bin/env bash
# End-to-end smoke test for gsheets-mcp.
#
# Builds the Docker image, runs the container on an alt port, verifies
# /healthz, MCP initialize, tools/list, and that an unauthenticated
# tools/call returns a structured error without crashing the container.
#
# Exit 0 on success, 1 on any failure.

set -euo pipefail

IMAGE="${IMAGE:-gsheets-mcp:smoke}"
PORT="${SMOKE_PORT:-8765}"
CONTAINER="gsheets-mcp-smoke-$$"
BASE_URL="http://127.0.0.1:${PORT}"

CYAN='\033[36m'
GREEN='\033[32m'
RED='\033[31m'
RESET='\033[0m'

log()  { printf "${CYAN}[smoke]${RESET} %s\n" "$*"; }
ok()   { printf "${GREEN}[ok]${RESET} %s\n" "$*"; }
fail() { printf "${RED}[fail]${RESET} %s\n" "$*" >&2; exit 1; }

cleanup() {
  local code=$?
  if docker ps --format '{{.Names}}' | grep -q "^${CONTAINER}$"; then
    log "stopping container ${CONTAINER}"
    docker stop "${CONTAINER}" >/dev/null 2>&1 || true
  fi
  exit "$code"
}
trap cleanup EXIT INT TERM

REPO_ROOT="$(cd "$(dirname "$0")/.." && pwd)"
cd "${REPO_ROOT}"

log "building image ${IMAGE}"
docker build -t "${IMAGE}" . >/dev/null

log "starting container ${CONTAINER} on port ${PORT}"
docker run --rm -d \
  --name "${CONTAINER}" \
  -p "${PORT}:${PORT}" \
  -e "PORT=${PORT}" \
  "${IMAGE}" >/dev/null

log "waiting for /healthz"
attempt=0
until curl -fsS "${BASE_URL}/healthz" >/dev/null 2>&1; do
  attempt=$((attempt + 1))
  if [ "${attempt}" -gt 30 ]; then
    fail "server never became healthy (30 polls)"
  fi
  sleep 0.5
done
ok "/healthz is up"

# ---- /healthz ----
HEALTH_CODE=$(curl -s -o /tmp/smoke-health.$$ -w "%{http_code}" "${BASE_URL}/healthz")
[ "${HEALTH_CODE}" = "200" ] || fail "/healthz returned ${HEALTH_CODE}"
grep -q '"status":"ok"' /tmp/smoke-health.$$ || fail "/healthz body unexpected: $(cat /tmp/smoke-health.$$)"
ok "GET /healthz returns 200 with status:ok"

# Helper: parse the SSE 'data:' line out of a streamable-HTTP response.
parse_sse_json() {
  local file="$1"
  # Some Accepts return plain JSON, others SSE. Handle both.
  if head -c1 "${file}" | grep -q '{'; then
    cat "${file}"
  else
    awk '/^data: /{sub(/^data: /, ""); print; exit}' "${file}"
  fi
}

# ---- /mcp initialize ----
INIT_BODY='{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2025-06-18","capabilities":{},"clientInfo":{"name":"smoke","version":"1"}}}'
curl -s -X POST "${BASE_URL}/mcp" \
  -H "Content-Type: application/json" \
  -H "Accept: application/json, text/event-stream" \
  -H "Authorization: Bearer fake" \
  -d "${INIT_BODY}" \
  > /tmp/smoke-init.$$
INIT_JSON=$(parse_sse_json /tmp/smoke-init.$$)
echo "${INIT_JSON}" | grep -q '"protocolVersion"' || fail "initialize missing protocolVersion: ${INIT_JSON}"
ok "POST /mcp initialize returns protocolVersion"

# ---- /mcp tools/list ----
LIST_BODY='{"jsonrpc":"2.0","id":2,"method":"tools/list","params":{}}'
curl -s -X POST "${BASE_URL}/mcp" \
  -H "Content-Type: application/json" \
  -H "Accept: application/json, text/event-stream" \
  -H "Authorization: Bearer fake" \
  -d "${LIST_BODY}" \
  > /tmp/smoke-list.$$
LIST_JSON=$(parse_sse_json /tmp/smoke-list.$$)
TOOLS_COUNT=$(echo "${LIST_JSON}" | node -e "let d='';process.stdin.on('data',c=>d+=c).on('end',()=>{try{const r=JSON.parse(d);console.log((r.result?.tools??[]).length)}catch(e){console.log(0)}}")
[ "${TOOLS_COUNT}" = "12" ] || fail "tools/list returned ${TOOLS_COUNT} tools, expected 12"
ok "POST /mcp tools/list returns 12 tools"

# ---- /mcp tools/call search_spreadsheets with fake bearer ----
CALL_BODY='{"jsonrpc":"2.0","id":3,"method":"tools/call","params":{"name":"search_spreadsheets","arguments":{"name":"smoke"}}}'
curl -s -X POST "${BASE_URL}/mcp" \
  -H "Content-Type: application/json" \
  -H "Accept: application/json, text/event-stream" \
  -H "Authorization: Bearer fake-token-that-is-definitely-invalid" \
  -d "${CALL_BODY}" \
  > /tmp/smoke-call.$$
CALL_JSON=$(parse_sse_json /tmp/smoke-call.$$)
# Either Google's unauthenticated error came back, OR we got a structured
# isError. Either way: container must NOT have crashed, and the response
# must look like a JSON-RPC envelope.
echo "${CALL_JSON}" | grep -q '"jsonrpc"' || fail "tools/call response is not JSON-RPC: ${CALL_JSON}"
echo "${CALL_JSON}" | grep -qE '"isError":\s*true|"error"' || fail "tools/call did not surface an error structure: ${CALL_JSON}"
# Verify the container is still running and serving healthz.
docker ps --format '{{.Names}}' | grep -q "^${CONTAINER}$" || fail "container died after tools/call"
curl -fsS "${BASE_URL}/healthz" >/dev/null || fail "/healthz unreachable after tools/call"
ok "POST /mcp tools/call returned structured error without crash"

rm -f /tmp/smoke-health.$$ /tmp/smoke-init.$$ /tmp/smoke-list.$$ /tmp/smoke-call.$$
printf "${GREEN}SUCCESS${RESET}\n"
