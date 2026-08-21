import { describe, it, expect } from 'vitest';
import {
  ApiError,
  parseRetryAfter,
  toolError,
  toolResponse,
  wrapHandler,
} from '../lib/errors.js';

describe('parseRetryAfter', () => {
  it('returns undefined for null/empty', () => {
    expect(parseRetryAfter(null)).toBeUndefined();
    expect(parseRetryAfter('')).toBeUndefined();
  });
  it('parses numeric seconds', () => {
    expect(parseRetryAfter('30')).toBe(30);
    expect(parseRetryAfter('0')).toBe(0);
  });
  it('ignores negative seconds (clamped or undefined)', () => {
    // -5 is rejected as a non-negative seconds count; the date-parse path
    // then either fails or clamps to 0 — either is acceptable, what
    // matters is we never return a negative wait.
    const v = parseRetryAfter('-5');
    expect(v === undefined || v === 0).toBe(true);
  });
  it('parses an HTTP-date in the future', () => {
    const future = new Date(Date.now() + 60_000).toUTCString();
    const v = parseRetryAfter(future);
    expect(v).toBeGreaterThan(0);
    expect(v).toBeLessThanOrEqual(60);
  });
  it('clamps a past HTTP-date to 0', () => {
    const past = new Date(Date.now() - 60_000).toUTCString();
    expect(parseRetryAfter(past)).toBe(0);
  });
  it('returns undefined for garbage', () => {
    expect(parseRetryAfter('not-a-date')).toBeUndefined();
  });
});

describe('toolResponse', () => {
  it('prefixes a notice above the payload, leaving structuredContent clean', () => {
    // The .xlsx read-only warning rides in the text channel only: putting it
    // in structuredContent would fail the tool's outputSchema.
    const res = toolResponse({ a: 1 }, 'HEADS UP');
    expect(res.content[0].text).toBe('HEADS UP\n{"a":1}');
    expect(res.structuredContent).toEqual({ a: 1 });
  });

  it('emits no prefix and no stray newline without a notice', () => {
    expect(toolResponse({ a: 1 }).content[0].text).toBe('{"a":1}');
  });

  it('wraps content and exposes structuredContent', () => {
    const result = toolResponse({ a: 1, b: 'two' });
    expect(result.structuredContent).toEqual({ a: 1, b: 'two' });
    expect(result.content).toHaveLength(1);
    expect(result.content[0].type).toBe('text');
    expect(JSON.parse(result.content[0].text)).toEqual({ a: 1, b: 'two' });
  });

  it('emits compact JSON with no indentation', () => {
    const res = toolResponse({ data: [[{ value: 'a', type: 'string' }]] });
    expect(res.content[0].text).not.toMatch(/\n/);
    expect(res.content[0].text).toBe('{"data":[[{"value":"a","type":"string"}]]}');
  });

  it('still round-trips to the structured payload', () => {
    const payload = { id: 'abc', rowCount: 2 };
    const res = toolResponse(payload);
    expect(JSON.parse(res.content[0].text)).toEqual(payload);
    expect(res.structuredContent).toEqual(payload);
  });
});

describe('toolError', () => {
  it('produces an isError envelope with the error message', () => {
    const r = toolError('boom');
    expect(r.isError).toBe(true);
    expect(r.content).toHaveLength(1);
    expect(JSON.parse(r.content[0].text)).toEqual({ error: 'boom' });
  });
  it('merges extra fields into the payload', () => {
    const r = toolError('boom', { code: 'x', status: 418 });
    expect(JSON.parse(r.content[0].text)).toEqual({
      error: 'boom',
      code: 'x',
      status: 418,
    });
  });
});

describe('wrapHandler', () => {
  it('passes successful results through unchanged', async () => {
    const handler = wrapHandler(async () => ({ ok: true }));
    await expect(handler()).resolves.toEqual({ ok: true });
  });

  it('converts a 401 ApiError into an unauthenticated error', async () => {
    const handler = wrapHandler(async () => {
      throw new ApiError('bad token', 401, 'sheets');
    });
    const r = await handler();
    expect(r.isError).toBe(true);
    const body = JSON.parse(r.content[0].text);
    expect(body.error).toBe('bad token');
    expect(body.code).toBe('unauthenticated');
    expect(body.status).toBe(401);
    expect(body.api).toBe('sheets');
  });

  it('converts a 403 ApiError into a permission_denied error', async () => {
    const handler = wrapHandler(async () => {
      throw new ApiError('nope', 403, 'drive');
    });
    const body = JSON.parse((await handler()).content[0].text);
    expect(body.code).toBe('permission_denied');
  });

  it('converts a 404 ApiError into not_found', async () => {
    const handler = wrapHandler(async () => {
      throw new ApiError('missing', 404, 'sheets');
    });
    const body = JSON.parse((await handler()).content[0].text);
    expect(body.code).toBe('not_found');
  });

  it('converts a 400 ApiError into invalid_argument with hint', async () => {
    const handler = wrapHandler(async () => {
      throw new ApiError('bad', 400, 'sheets');
    });
    const body = JSON.parse((await handler()).content[0].text);
    expect(body.code).toBe('invalid_argument');
    expect(body.hint).toMatch(/range/i);
  });

  it('converts a 429 with Retry-After into rate_limited with seconds', async () => {
    const handler = wrapHandler(async () => {
      throw new ApiError('slow', 429, 'sheets', 12);
    });
    const body = JSON.parse((await handler()).content[0].text);
    expect(body.code).toBe('rate_limited');
    expect(body.retryAfterSeconds).toBe(12);
    expect(body.hint).toMatch(/back off/i);
  });

  it('converts a 429 without Retry-After (no retryAfterSeconds key)', async () => {
    const handler = wrapHandler(async () => {
      throw new ApiError('slow', 429, 'sheets');
    });
    const body = JSON.parse((await handler()).content[0].text);
    expect(body.code).toBe('rate_limited');
    expect('retryAfterSeconds' in body).toBe(false);
  });

  it('converts a 5xx into server_error with retry hint', async () => {
    const handler = wrapHandler(async () => {
      throw new ApiError('boom', 503, 'sheets');
    });
    const body = JSON.parse((await handler()).content[0].text);
    expect(body.code).toBe('server_error');
    expect(body.hint).toMatch(/retry/i);
  });

  it('falls back to plain message for non-ApiError throws', async () => {
    const handler = wrapHandler(async () => {
      throw new Error('something else');
    });
    const r = await handler();
    expect(r.isError).toBe(true);
    const body = JSON.parse(r.content[0].text);
    expect(body.error).toBe('something else');
    expect('code' in body).toBe(false);
  });

  it('stringifies non-Error throws', async () => {
    const handler = wrapHandler(async () => {
      throw 'plain string';
    });
    const body = JSON.parse((await handler()).content[0].text);
    expect(body.error).toBe('plain string');
  });
});
