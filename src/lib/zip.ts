/**
 * Reads a zip's central directory without inflating anything, so the cost of
 * opening an .xlsx can be judged before SheetJS pays it.
 */

const EOCD_SIGNATURE = 0x06054b50;
const CENTRAL_ENTRY_SIGNATURE = 0x02014b50;
const EOCD_MIN_BYTES = 22;
const CENTRAL_ENTRY_MIN_BYTES = 46;
const MAX_COMMENT_BYTES = 0xffff;
const ZIP64_MARKER_16 = 0xffff;
const ZIP64_MARKER_32 = 0xffffffff;

/**
 * Sum of every entry's uncompressed size, as the central directory declares
 * it. SheetJS inflates every entry of the archive, not only the tab asked
 * for, so the whole sum is what a read costs.
 *
 * Returns Infinity for a ZIP64 archive (sizes past 4GB, never a readable
 * workbook here) and undefined when no well-formed directory is found, which
 * leaves the verdict to the parser.
 */
export function zipUncompressedBytes(bytes: Uint8Array): number | undefined {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const lowest = Math.max(0, bytes.length - EOCD_MIN_BYTES - MAX_COMMENT_BYTES);

  for (let eocd = bytes.length - EOCD_MIN_BYTES; eocd >= lowest; eocd--) {
    if (view.getUint32(eocd, true) !== EOCD_SIGNATURE) continue;

    const entries = view.getUint16(eocd + 10, true);
    const directoryOffset = view.getUint32(eocd + 16, true);
    if (entries === ZIP64_MARKER_16 || directoryOffset === ZIP64_MARKER_32) return Infinity;

    let total = 0;
    let entry = directoryOffset;
    for (let i = 0; i < entries; i++) {
      if (entry + CENTRAL_ENTRY_MIN_BYTES > bytes.length
        || view.getUint32(entry, true) !== CENTRAL_ENTRY_SIGNATURE) {
        return undefined;
      }
      const size = view.getUint32(entry + 24, true);
      if (size === ZIP64_MARKER_32) return Infinity;
      total += size;
      entry += CENTRAL_ENTRY_MIN_BYTES
        + view.getUint16(entry + 28, true)
        + view.getUint16(entry + 30, true)
        + view.getUint16(entry + 32, true);
    }
    return total;
  }
  return undefined;
}
