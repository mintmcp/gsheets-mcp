import { describe, it, expect } from 'vitest';
import { zipSync } from 'fflate';
import { zipUncompressedBytes } from '../lib/zip.js';
import { fixture } from './__fixtures__/index.js';

const archive = (sizes: Record<string, number>) =>
  zipSync(
    Object.fromEntries(Object.entries(sizes).map(([name, n]) => [name, new Uint8Array(n)])),
    { level: 9 },
  );

/** Offset of the end-of-central-directory record (no archive comment). */
const eocdOf = (bytes: Uint8Array) => bytes.length - 22;

describe('zipUncompressedBytes', () => {
  it('sums the uncompressed size of every entry without inflating them', () => {
    const bytes = archive({ 'a.xml': 100_000, 'b/c.xml': 2_500, 'd.bin': 0 });
    expect(bytes.length).toBeLessThan(10_000);
    expect(zipUncompressedBytes(bytes)).toBe(102_500);
  });

  it('reads a real workbook as larger than its compressed bytes', () => {
    const bytes = fixture('basic.xlsx');
    expect(zipUncompressedBytes(bytes)).toBeGreaterThan(bytes.length);
  });

  it('finds the directory past an archive comment', () => {
    const plain = archive({ 'a.xml': 1_000 });
    const comment = new TextEncoder().encode('made by a tool that likes comments');
    const bytes = new Uint8Array(plain.length + comment.length);
    bytes.set(plain);
    bytes.set(comment, plain.length);
    new DataView(bytes.buffer).setUint16(eocdOf(plain) + 20, comment.length, true);
    expect(zipUncompressedBytes(bytes)).toBe(1_000);
  });

  it('reports a ZIP64 archive as unboundedly large', () => {
    const bytes = archive({ 'a.xml': 10 });
    new DataView(bytes.buffer).setUint16(eocdOf(bytes) + 10, 0xffff, true);
    expect(zipUncompressedBytes(bytes)).toBe(Infinity);
  });

  it('leaves the verdict to the parser when there is no directory', () => {
    expect(zipUncompressedBytes(new Uint8Array(64))).toBeUndefined();
    expect(zipUncompressedBytes(new Uint8Array(0))).toBeUndefined();
  });

  it('leaves the verdict to the parser when the directory is corrupt', () => {
    const bytes = archive({ 'a.xml': 10 });
    const view = new DataView(bytes.buffer);
    view.setUint32(view.getUint32(eocdOf(bytes) + 16, true), 0, true);
    expect(zipUncompressedBytes(bytes)).toBeUndefined();
  });
});
