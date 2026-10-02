import { describe, it, expect, vi, afterEach } from 'vitest';
import { runWithDeadline, currentDeadline } from '../lib/deadline.js';
import { makeDriveRequest, makeSheetsRequest } from '../lib/google.js';
import { ApiError, DeadlineExceededError } from '../lib/errors.js';

/** A fetch that never answers on its own, only rejecting once its signal aborts. */
function hangUntilAborted(_input: unknown, init?: RequestInit): Promise<Response> {
  return new Promise((_, reject) => {
    const signal = init?.signal;
    signal?.addEventListener('abort', () => reject(signal.reason), { once: true });
  });
}

afterEach(() => {
  vi.unstubAllGlobals();
});

describe('runWithDeadline', () => {
  it('exposes a live signal only inside the call', async () => {
    expect(currentDeadline()).toBeUndefined();
    await runWithDeadline(async () => {
      expect(currentDeadline()?.aborted).toBe(false);
    });
  });

  it('aborts the signal once the time is up', async () => {
    await runWithDeadline(async () => {
      await new Promise((r) => setTimeout(r, 40));
      expect(currentDeadline()?.aborted).toBe(true);
    }, 10);
  });

  it('lets a nested call shorten the enclosing deadline', async () => {
    await runWithDeadline(() => runWithDeadline(async () => {
      await new Promise((r) => setTimeout(r, 40));
      expect(currentDeadline()?.aborted).toBe(true);
    }, 10), 60_000);
  });

  it('never lets a nested call extend the enclosing deadline', async () => {
    await runWithDeadline(() => runWithDeadline(async () => {
      await new Promise((r) => setTimeout(r, 40));
      expect(currentDeadline()?.aborted).toBe(true);
    }, 60_000), 10);
  });
});

describe('upstream requests under a deadline', () => {
  it('passes the deadline signal to fetch', async () => {
    let seen: AbortSignal | null | undefined;
    vi.stubGlobal('fetch', async (_input: unknown, init?: RequestInit) => {
      seen = init?.signal;
      return new Response('{}', { headers: { 'content-type': 'application/json' } });
    });
    await runWithDeadline(async () => {
      await makeSheetsRequest('/abc', 't');
      expect(seen).toBe(currentDeadline());
    });
  });

  it('reports a request Google never answers as DeadlineExceededError', async () => {
    vi.stubGlobal('fetch', hangUntilAborted);
    const call = runWithDeadline(() => makeSheetsRequest('/abc', 't'), 20);
    await expect(call).rejects.toBeInstanceOf(DeadlineExceededError);
    await expect(call).rejects.toMatchObject({ api: 'sheets' });
  });

  it('reports a body that stalls mid-stream the same way', async () => {
    vi.stubGlobal('fetch', async (_input: unknown, init?: RequestInit) => {
      const body = new ReadableStream({
        start(controller) {
          controller.enqueue(new TextEncoder().encode('{"a":'));
          init?.signal?.addEventListener('abort', () => controller.error(init.signal!.reason));
        },
      });
      return new Response(body, { headers: { 'content-type': 'application/json' } });
    });
    await expect(runWithDeadline(() => makeSheetsRequest('/abc', 't'), 20))
      .rejects.toBeInstanceOf(DeadlineExceededError);
  });

  it('never cuts off a request that changes something', async () => {
    // Drive still makes the copy if we stop waiting, and the caller loses the
    // new file's id, so a retry makes a second copy.
    let seen: AbortSignal | null | undefined = null;
    vi.stubGlobal('fetch', async (_input: unknown, init?: RequestInit) => {
      seen = init?.signal;
      await new Promise((r) => setTimeout(r, 60));
      return new Response('{"id":"new1"}', { headers: { 'content-type': 'application/json' } });
    });
    const call = runWithDeadline(
      () => makeDriveRequest('/files/abc/copy', 't', { method: 'POST', body: '{}' }),
      20,
    );
    await expect(call).resolves.toEqual({ id: 'new1' });
    expect(seen).toBeUndefined();
  });

  it('still cuts off a read made with an explicit GET', async () => {
    vi.stubGlobal('fetch', hangUntilAborted);
    await expect(runWithDeadline(() => makeSheetsRequest('/abc', 't', { method: 'GET' }), 20))
      .rejects.toBeInstanceOf(DeadlineExceededError);
  });

  it('leaves an answered error as the ApiError it was', async () => {
    vi.stubGlobal('fetch', async () => new Response('{}', { status: 500 }));
    await expect(runWithDeadline(() => makeSheetsRequest('/abc', 't')))
      .rejects.toBeInstanceOf(ApiError);
  });
});
