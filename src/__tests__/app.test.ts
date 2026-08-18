import { describe, it, expect, afterAll, vi } from 'vitest';
import http from 'node:http';
import type { AddressInfo } from 'node:net';
import { createApp, MCP_PATH } from '../app.js';

/**
 * Regression test for transport collisions. A shared module-level McpServer
 * only breaks when one handler holds the connection while awaiting Google —
 * a bare `initialize` returns too fast to overlap. So we gate the upstream
 * fetch to hold request A open while request B connects.
 */

const realFetch = globalThis.fetch;

let inFlight = 0;
let releaseUpstream: () => void;
const upstreamGate = new Promise<void>((resolve) => {
  releaseUpstream = resolve;
});

vi.stubGlobal('fetch', async (input: any, init?: any) => {
  const url = typeof input === 'string' ? input : String(input?.url ?? input);
  if (url.startsWith('http://127.0.0.1')) return realFetch(input, init);

  inFlight++;
  await upstreamGate;
  return new Response(
    JSON.stringify({ sheets: [{ properties: { title: 'Sheet1', index: 0 } }], properties: { title: 'T' }, spreadsheetUrl: 'https://x' }),
    { status: 200, headers: { 'content-type': 'application/json' } },
  );
});

const server = http.createServer(createApp());
await new Promise<void>((resolve) => server.listen(0, '127.0.0.1', resolve));
const { port } = server.address() as AddressInfo;
const url = `http://127.0.0.1:${port}${MCP_PATH}`;

afterAll(() => new Promise<void>((resolve) => server.close(() => resolve())));

function rpc(id: number, method: string, params: unknown) {
  return realFetch(url, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Accept: 'application/json, text/event-stream',
      Authorization: 'Bearer test-token',
    },
    body: JSON.stringify({ jsonrpc: '2.0', id, method, params }),
  });
}

const callGetMetadata = (id: number) =>
  rpc(id, 'tools/call', {
    name: 'get_metadata',
    arguments: { spreadsheet_id: 'sheet-abc' },
  });

async function waitUntil(predicate: () => boolean, timeoutMs = 2000) {
  const deadline = Date.now() + timeoutMs;
  while (!predicate()) {
    if (Date.now() > deadline) throw new Error('timed out waiting for condition');
    await new Promise((r) => setTimeout(r, 10));
  }
}

describe('POST /mcp with overlapping tool calls', () => {
  it('does not collide when a second request arrives mid-flight', async () => {
    const a = callGetMetadata(1);
    await waitUntil(() => inFlight >= 1);

    // A is now parked inside the upstream fetch, holding its transport.
    const b = callGetMetadata(2);

    // Both requests must be inside the upstream at once, or the test proves
    // nothing. A server that rejects B at connect() never gets here and fails
    // this wait, which is the regression we are guarding against.
    await waitUntil(() => inFlight >= 2, 2000);
    expect(inFlight).toBe(2);

    releaseUpstream();
    const [resA, resB] = await Promise.all([a, b]);

    for (const res of [resA, resB]) {
      expect(res.status).toBe(200);
      expect(await res.text()).not.toContain('Already connected to a transport');
    }
  });
});
