import { describe, it, expect, afterEach, vi } from 'vitest';
import { readTools } from '../tools/read.js';
import { requestContext } from '../auth.js';

function jsonResponse(body: unknown) {
  return new Response(JSON.stringify(body), {
    status: 200,
    headers: { 'content-type': 'application/json' },
  });
}

const call = (args: Record<string, unknown>) =>
  requestContext.run({ accessToken: 'test-token' }, () =>
    (readTools as any).list_comments.handler(args),
  );

const payload = (res: any) => res.structuredContent ?? JSON.parse(res.content[0].text);

afterEach(() => vi.unstubAllGlobals());

describe('list_comments', () => {
  it('reads Drive comments with their replies and the cell anchor', async () => {
    const urls: string[] = [];
    vi.stubGlobal('fetch', async (input: any) => {
      urls.push(decodeURIComponent(String(input)));
      return jsonResponse({
        nextPageToken: 'tok',
        comments: [{
          id: 'c1', content: 'Check this', createdTime: '2026-01-01T00:00:00Z', modifiedTime: '2026-01-02T00:00:00Z',
          author: { displayName: 'Ana', emailAddress: 'ana@example.com' }, resolved: false,
          anchor: '{"type":"workbook-range","uid":0,"range":"123"}',
          quotedFileContent: { value: '42' },
          replies: [{ id: 'r1', content: 'Done', author: { displayName: 'Bo' }, createdTime: '2026-01-03T00:00:00Z' }],
        }],
      });
    });
    const out = payload(await call({ spreadsheet_id: 'abc' }));

    expect(urls[0]).toContain('/drive/v3/files/abc/comments?');
    expect(urls[0]).toContain('pageSize=100');
    expect(urls[0]).toContain('fields=');
    expect(out.nextPageToken).toBe('tok');
    expect(out.comments).toEqual([{
      id: 'c1', content: 'Check this', author: 'Ana', createdTime: '2026-01-01T00:00:00Z',
      modifiedTime: '2026-01-02T00:00:00Z', resolved: false, quoted: '42',
      anchor: '{"type":"workbook-range","uid":0,"range":"123"}',
      replies: [{ id: 'r1', content: 'Done', author: 'Bo', createdTime: '2026-01-03T00:00:00Z' }],
    }]);
  });

  it('passes a page token through and reports null when there is no more', async () => {
    const urls: string[] = [];
    vi.stubGlobal('fetch', async (input: any) => {
      urls.push(String(input));
      return jsonResponse({ comments: [] });
    });
    const out = payload(await call({ spreadsheet_id: 'abc', page_token: 'tok' }));
    expect(urls[0]).toContain('pageToken=tok');
    expect(out).toEqual({ id: 'abc', comments: [], nextPageToken: null });
  });

  it('is gated on the Drive read scope', () => {
    expect((readTools as any).list_comments.handler.scope).toBe('https://www.googleapis.com/auth/drive.readonly');
  });
});
