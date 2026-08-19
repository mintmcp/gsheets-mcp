import { describe, it, expect } from 'vitest';
import { collectStream, readJsonWithLimit, readTextCapped } from '../lib/google.js';

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

describe('collectStream', () => {
  it('returns null when the response has no body', async () => {
    expect(await collectStream(new Response(null), 1_000)).toBeNull();
  });

  it('returns the whole body when it fits', async () => {
    const result = await collectStream(streamed(['abc', 'def']), 1_000);
    expect(result?.overflowed).toBe(false);
    expect(new TextDecoder().decode(result!.bytes)).toBe('abcdef');
  });

  it('does not flag a body that lands exactly on the limit', async () => {
    const result = await collectStream(streamed(['abcde']), 5);
    expect(result?.overflowed).toBe(false);
    expect(result!.bytes).toHaveLength(5);
  });

  it('truncates to the limit and reports overflow', async () => {
    // Each caller decides what overflow means: readTextCapped keeps the
    // prefix, readJsonWithLimit and the Drive download both throw.
    const result = await collectStream(streamed(['abc', 'def', 'ghi']), 4);
    expect(result?.overflowed).toBe(true);
    expect(new TextDecoder().decode(result!.bytes)).toBe('abcd');
  });

  it('stops reading rather than draining an oversized body', async () => {
    let enqueued = 0;
    const response = new Response(new ReadableStream({
      pull(controller) {
        enqueued++;
        if (enqueued > 50) return controller.close();
        controller.enqueue(new TextEncoder().encode('x'.repeat(10)));
      },
    }));
    const result = await collectStream(response, 25);
    expect(result?.overflowed).toBe(true);
    expect(enqueued).toBeLessThan(10);
  });
});
