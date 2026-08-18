import { describe, it, expect } from 'vitest';
import { readJsonWithLimit, readTextCapped } from '../lib/google.js';

function streamed(chunks: string[]): Response {
  const encoder = new TextEncoder();
  return new Response(new ReadableStream({
    start(controller) {
      for (const chunk of chunks) controller.enqueue(encoder.encode(chunk));
      controller.close();
    },
  }));
}

describe('readJsonWithLimit', () => {
  it('parses a response under the limit', async () => {
    await expect(readJsonWithLimit(streamed(['{"a":', '1}']), 1_000)).resolves.toEqual({ a: 1 });
  });

  it('throws once the byte budget is exceeded', async () => {
    const chunks = Array.from({ length: 20 }, () => 'x'.repeat(500));
    await expect(readJsonWithLimit(streamed(chunks), 1_000)).rejects.toThrow(/too large/i);
  });

  it('names the range argument in the oversized error so the model can recover', async () => {
    const chunks = Array.from({ length: 20 }, () => 'x'.repeat(500));
    await expect(readJsonWithLimit(streamed(chunks), 1_000)).rejects.toThrow(/range/);
  });

  it('reports invalid JSON distinctly from an oversized body', async () => {
    await expect(readJsonWithLimit(streamed(['not json']), 1_000)).rejects.toThrow(/could not be parsed/i);
  });

  it('handles an empty body', async () => {
    await expect(readJsonWithLimit(streamed([]), 1_000)).rejects.toThrow(/could not be parsed/i);
  });

  it('decodes multi-byte UTF-8 split across chunk boundaries', async () => {
    const encoder = new TextEncoder();
    const full = encoder.encode(JSON.stringify({ name: 'café — ☕' }));
    const mid = Math.floor(full.length / 2);
    const response = new Response(new ReadableStream({
      start(controller) {
        controller.enqueue(full.slice(0, mid));
        controller.enqueue(full.slice(mid));
        controller.close();
      },
    }));
    await expect(readJsonWithLimit(response, 1_000)).resolves.toEqual({ name: 'café — ☕' });
  });
});

describe('readTextCapped', () => {
  it('returns a short body unchanged', async () => {
    const res = new Response('{"error":{"message":"nope"}}');
    await expect(readTextCapped(res, 1_000)).resolves.toBe('{"error":{"message":"nope"}}');
  });

  it('caps a huge error body instead of buffering it whole', async () => {
    const encoder = new TextEncoder();
    let produced = 0;
    const body = new ReadableStream({
      pull(controller) {
        if (produced >= 5_000_000) return controller.close();
        produced += 100_000;
        controller.enqueue(encoder.encode('e'.repeat(100_000)));
      },
    });
    const text = await readTextCapped(new Response(body), 64 * 1024);
    expect(text.length).toBe(64 * 1024);
  });

  it('returns empty string when there is no body', async () => {
    const res = new Response(null, { status: 204 });
    await expect(readTextCapped(res, 1_000)).resolves.toBe('');
  });
});
