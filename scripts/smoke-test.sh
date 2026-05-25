#!/usr/bin/env bash
# End-to-end smoke test for gsheets-mcp.
#
# Builds the Docker image, runs the container on an alt port, verifies
# /healthz, MCP initialize, tools/list, that a tools/call with NO
# Authorization header returns the local missing-token structured error,
# and that a tools/call with a fake bearer returns the upstream-401
# structured error — all without crashing the container.
#
# Exit 0 on success, 1 on any failure.

set -euo pipefail

IMAGE="${IMAGE:-gsheets-mcp:smoke}"
PORT="${SMOKE_PORT:-8765}"
CONTAINER="gsheets-mcp-smoke-$$"
BASE_URL="http://127.0.0.1:${PORT}"
TMP_DIR="$(mktemp -d -t gsheets-mcp-smoke.XXXXXX)"

CYAN='\033[36m'
GREEN='\033[32m'
RED='\033[31m'
RESET='\033[0m'

log()  { printf "${CYAN}[smoke]${RESET} %s\n" "$*"; }
ok()   { printf "${GREEN}[ok]${RESET} %s\n" "$*"; }
fail() { printf "${RED}[fail]${RESET} %s\n" "$*" >&2; exit 1; }

cleanup() {
  local code=$?
  if docker ps --format '{{.Names}}' 2>/dev/null | grep -q "^${CONTAINER}$"; then
    log "stopping container ${CONTAINER}"
    docker stop "${CONTAINER}" >/dev/null 2>&1 || true
  fi
  rm -rf "${TMP_DIR}" 2>/dev/null || true
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
HEALTH_CODE=$(curl -s -o "${TMP_DIR}/health" -w "%{http_code}" "${BASE_URL}/healthz")
[ "${HEALTH_CODE}" = "200" ] || fail "/healthz returned ${HEALTH_CODE}"
grep -q '"status":"ok"' "${TMP_DIR}/health" || fail "/healthz body unexpected: $(cat "${TMP_DIR}/health")"
ok "GET /healthz returns 200 with status:ok"

# Helper: parse the SSE 'data:' line out of a streamable-HTTP response.
parse_sse_json() {
  local file="$1"
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
  > "${TMP_DIR}/init"
INIT_JSON=$(parse_sse_json "${TMP_DIR}/init")
echo "${INIT_JSON}" | grep -q '"protocolVersion"' || fail "initialize missing protocolVersion: ${INIT_JSON}"
ok "POST /mcp initialize returns protocolVersion"

# ---- /mcp tools/list ----
LIST_BODY='{"jsonrpc":"2.0","id":2,"method":"tools/list","params":{}}'
curl -s -X POST "${BASE_URL}/mcp" \
  -H "Content-Type: application/json" \
  -H "Accept: application/json, text/event-stream" \
  -H "Authorization: Bearer fake" \
  -d "${LIST_BODY}" \
  > "${TMP_DIR}/list"
LIST_JSON=$(parse_sse_json "${TMP_DIR}/list")
TOOLS_COUNT=$(echo "${LIST_JSON}" | node -e "let d='';process.stdin.on('data',c=>d+=c).on('end',()=>{try{const r=JSON.parse(d);console.log((r.result?.tools??[]).length)}catch(e){console.log(0)}}")
[ "${TOOLS_COUNT}" = "12" ] || fail "tools/list returned ${TOOLS_COUNT} tools, expected 12"
ok "POST /mcp tools/list returns 12 tools"

# ---- /mcp tools/call WITHOUT Authorization header ----
# Exercises the local missing-token branch in src/auth.ts (withGoogleAuth).
CALL_BODY='{"jsonrpc":"2.0","id":3,"method":"tools/call","params":{"name":"search_spreadsheets","arguments":{"name":"smoke"}}}'
curl -s -X POST "${BASE_URL}/mcp" \
  -H "Content-Type: application/json" \
  -H "Accept: application/json, text/event-stream" \
  -d "${CALL_BODY}" \
  > "${TMP_DIR}/call-noauth"
NOAUTH_JSON=$(parse_sse_json "${TMP_DIR}/call-noauth")
echo "${NOAUTH_JSON}" | grep -q '"jsonrpc"' || fail "no-auth tools/call response is not JSON-RPC: ${NOAUTH_JSON}"
echo "${NOAUTH_JSON}" | grep -q 'Missing Google access token' \
  || fail "no-auth tools/call did not surface the local missing-token error: ${NOAUTH_JSON}"
echo "${NOAUTH_JSON}" | grep -qE '"isError":\s*true' \
  || fail "no-auth tools/call missing isError:true flag: ${NOAUTH_JSON}"
ok "POST /mcp tools/call without Authorization returns local missing-token structured error"

# ---- /mcp tools/call WITH a fake bearer ----
# Exercises the upstream-error path (Google returns 401; wrapHandler turns
# the ApiError into a structured isError response).
curl -s -X POST "${BASE_URL}/mcp" \
  -H "Content-Type: application/json" \
  -H "Accept: application/json, text/event-stream" \
  -H "Authorization: Bearer fake-token-that-is-definitely-invalid" \
  -d "${CALL_BODY}" \
  > "${TMP_DIR}/call-fakeauth"
FAKE_JSON=$(parse_sse_json "${TMP_DIR}/call-fakeauth")
echo "${FAKE_JSON}" | grep -q '"jsonrpc"' || fail "fake-auth tools/call response is not JSON-RPC: ${FAKE_JSON}"
echo "${FAKE_JSON}" | grep -qE '"isError":\s*true|"error"' \
  || fail "fake-auth tools/call did not surface an error structure: ${FAKE_JSON}"
ok "POST /mcp tools/call with fake bearer returns structured upstream error"

# Container still alive and serving healthz?
docker ps --format '{{.Names}}' | grep -q "^${CONTAINER}$" || fail "container died during smoke run"
curl -fsS "${BASE_URL}/healthz" >/dev/null || fail "/healthz unreachable at end of smoke run"

printf "${GREEN}SUCCESS${RESET}\n"
