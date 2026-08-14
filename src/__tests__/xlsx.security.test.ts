/**
 * Adversarial tests: the parser runs in a Worker on bytes any Drive collaborator
 * can control, so every bound here is a live DoS boundary, not a nicety.
 *
 * SheetJS owns zip and XML safety now. What remains ours is the OUTPUT budget —
 * a small workbook can still expand into hundreds of MB of JSON, because one
 * shared string referenced by N cells costs O(1) of input per reference.
 */

import { describe, it, expect } from 'vitest';
import { zipSync, strToU8 } from 'fflate';
import {
  parseXlsx, toCell,
  XlsxInvalidError, XlsxEncryptedError,
  MAX_CELLS, MAX_CELL_CHARS, MAX_SHEETS, MAX_SHEET_NAME_CHARS,
} from '../lib/xlsx.js';

/** Builds a minimal but valid .xlsx around the given sheet XML (one per tab, or shared). */
function workbook(
  sheetXml: string | string[],
  sheets = [{ name: 'S1', file: 'sheet1.xml' }]
): Uint8Array {
  const parts: Record<string, Uint8Array> = {
    '[Content_Types].xml': strToU8(
      `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
      sheets.map((s) => `<Override PartName="/xl/worksheets/${s.file}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('') +
      `</Types>`
    ),
    '_rels/.rels': strToU8(
      `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`
    ),
    'xl/workbook.xml': strToU8(
      `<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ` +
      `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>` +
      sheets.map((s, i) => `<sheet name="${s.name}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('') +
      `</sheets></workbook>`
    ),
    'xl/_rels/workbook.xml.rels': strToU8(
      `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      sheets.map((s, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/${s.file}"/>`).join('') +
      `</Relationships>`
    ),
  };
  sheets.forEach((s, i) => {
    parts[`xl/worksheets/${s.file}`] = strToU8(Array.isArray(sheetXml) ? sheetXml[i] : sheetXml);
  });
  return zipSync(parts, { level: 9 });
}

const sheetOf = (body: string) =>
  `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${body}</sheetData></worksheet>`;

describe('output volume budget', () => {
  it('stops a shared-string amplification attack at the character budget', () => {
    const long = 'x'.repeat(20_000);
    const rows = Array.from({ length: 200 }, (_, i) =>
      `<row r="${i + 1}"><c r="A${i + 1}" t="inlineStr"><is><t>${long}</t></is></c></row>`
    ).join('');
    const wb = parseXlsx(workbook(sheetOf(rows)), { maxChars: 100_000 });
    expect(wb.truncated).toBe(true);
    const emitted = wb.sheets[0].data.flat().reduce((n, c) => n + c.value.length, 0);
    expect(emitted).toBeLessThan(200 * long.length);
  });

  it('charges the JSON envelope, so empty padding cells cannot slip the budget', () => {
    const rows = Array.from({ length: 50 }, (_, i) =>
      `<row r="${i + 1}"><c r="A${i + 1}" t="inlineStr"><is><t>x</t></is></c></row>`
    ).join('');
    const wb = parseXlsx(workbook(sheetOf(rows)));
    expect(wb.chars).toBeGreaterThan(50 * 10);
  });

  it('caps a single cell value', () => {
    const huge = 'y'.repeat(MAX_CELL_CHARS + 5_000);
    const wb = parseXlsx(workbook(sheetOf(`<row r="1"><c r="A1" t="inlineStr"><is><t>${huge}</t></is></c></row>`)));
    expect(wb.sheets[0].data[0][0].value.length).toBe(MAX_CELL_CHARS);
  });

  it('stops at the cell budget', () => {
    const rows = Array.from({ length: 100 }, (_, i) =>
      `<row r="${i + 1}"><c r="A${i + 1}" t="inlineStr"><is><t>v</t></is></c></row>`
    ).join('');
    const wb = parseXlsx(workbook(sheetOf(rows)), { maxCells: 10 });
    expect(wb.truncated).toBe(true);
    expect(wb.cells).toBeLessThanOrEqual(10);
  });
});

describe('leading-row padding', () => {
  // A 1.4KB file whose unbudgeted row padding used to fill the isolate.
  const farDown = (row: number) =>
    workbook(sheetOf(`<row r="${row}"><c r="A${row}" t="inlineStr"><is><t>x</t></is></c></row>`));

  it('charges padded rows against the budget instead of allocating millions', () => {
    const wb = parseXlsx(farDown(900_000_000));
    expect(wb.truncated).toBe(true);
    expect(wb.sheets[0].rowCount).toBeLessThanOrEqual(MAX_CELLS);
  });

  it('bounds a legal but empty far-down cell at the very last Excel row', () => {
    const wb = parseXlsx(farDown(1_048_576), { maxCells: 100 });
    expect(wb.truncated).toBe(true);
    expect(wb.sheets[0].rowCount).toBeLessThanOrEqual(100);
  });

  it('shares one budget across tabs, so many hostile tabs cost no more than one', () => {
    const many = Array.from({ length: 40 }, (_, i) => ({ name: `S${i}`, file: `sheet${i + 1}.xml` }));
    const hostile = sheetOf('<row r="1048576"><c r="A1048576" t="inlineStr"><is><t>x</t></is></c></row>');
    const wb = parseXlsx(workbook(hostile, many), { maxCells: 500 });
    expect(wb.truncated).toBe(true);
    expect(wb.sheets.reduce((n, s) => n + s.rowCount, 0)).toBeLessThanOrEqual(500);
  });

  it('still aligns row numbers with the spreadsheet when the padding fits', () => {
    const wb = parseXlsx(farDown(5));
    expect(wb.truncated).toBe(false);
    expect(wb.sheets[0].data[4][0].value).toBe('x');
  });
});

describe('selective parsing', () => {
  const twoTabs = () => workbook(
    [
      sheetOf(Array.from({ length: 60 }, (_, i) =>
        `<row r="${i + 1}"><c r="A${i + 1}" t="inlineStr"><is><t>one</t></is></c></row>`).join('')),
      sheetOf('<row r="1"><c r="A1" t="inlineStr"><is><t>two</t></is></c></row>'),
    ],
    [{ name: 'S1', file: 'sheet1.xml' }, { name: 'S2', file: 'sheet2.xml' }]
  );

  it('spends the whole budget on the requested tab, not on the ones before it', () => {
    const wb = parseXlsx(twoTabs(), { maxCells: 20, sheet: 'S2' });
    expect(wb.sheets[1].data[0][0].value).toBe('two');
    expect(wb.sheets[1].truncated).toBe(false);
  });

  it('marks a skipped tab as not requested rather than unreadable', () => {
    const wb = parseXlsx(twoTabs(), { sheet: 'S2' });
    expect(wb.sheets[0].notRequested).toBe(true);
    expect(wb.sheets[0].unreadable).toBeUndefined();
    expect(wb.sheets.map((s) => s.name)).toEqual(['S1', 'S2']);
  });

  it('accepts a tab index and still names every tab', () => {
    const wb = parseXlsx(twoTabs(), { sheet: 0 });
    expect(wb.sheets[0].data[0][0].value).toBe('one');
    expect(wb.sheets[1].notRequested).toBe(true);
  });

  it('does not mistake a name that matches no tab for a broken workbook', () => {
    const wb = parseXlsx(twoTabs(), { sheet: 'Nope' });
    expect(wb.sheets.every((s) => s.notRequested)).toBe(true);
  });
});

describe('workbook-level bounds', () => {
  it('caps the number of sheets and says how many it left out', () => {
    const many = Array.from({ length: MAX_SHEETS + 50 }, (_, i) => ({
      name: `S${i}`, file: `sheet${i + 1}.xml`,
    }));
    const wb = parseXlsx(workbook(sheetOf('<row r="1"/>'), many), { namesOnly: true });
    expect(wb.sheets.length).toBe(MAX_SHEETS);
    // Silently dropping them makes a later read fail as a bare "Sheet not found".
    expect(wb.sheetsOmitted).toBe(50);
  });

  it('caps a tab name so metadata cannot be bulked out, and flags the shortening', () => {
    const wb = parseXlsx(
      workbook(sheetOf('<row r="1"/>'), [{ name: 'S'.repeat(5_000), file: 'sheet1.xml' }]),
      { namesOnly: true }
    );
    expect(wb.sheets[0].name.length).toBe(MAX_SHEET_NAME_CHARS);
    expect(wb.sheets[0].nameShortened).toBe(true);
  });

  it('leaves both signals off for an ordinary workbook', () => {
    const wb = parseXlsx(workbook(sheetOf('<row r="1"/>')), { namesOnly: true });
    expect(wb.sheetsOmitted).toBe(0);
    expect(wb.sheets[0].nameShortened).toBeUndefined();
  });
});

describe('hostile input rejection', () => {
  it('rejects a file that is not a zip', () => {
    expect(() => parseXlsx(new Uint8Array([1, 2, 3, 4, 5]))).toThrow(XlsxInvalidError);
  });

  it('rejects an OLE container rather than parsing it as legacy .xls', () => {
    const ole = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 0, 0]);
    expect(() => parseXlsx(ole)).toThrow(XlsxEncryptedError);
  });

  it('rejects a zip that is not a workbook, as a classified error', () => {
    const notAWorkbook = zipSync({ 'hello.txt': strToU8('hi') });
    expect(() => parseXlsx(notAWorkbook)).toThrow(XlsxInvalidError);
  });
});

describe('hyperlink safety', () => {
  it('drops schemes that are not browsable', () => {
    for (const url of ['javascript:alert(1)', 'data:text/html,<script>', 'file:///etc/passwd']) {
      const cell = toCell({ t: 's', v: 'click', w: 'click', l: { Target: url } });
      expect(cell.hyperlinks).toBeUndefined();
    }
  });

  it('keeps http, https and mailto', () => {
    for (const url of ['https://example.com', 'http://example.com', 'mailto:a@b.com']) {
      const cell = toCell({ t: 's', v: 'click', w: 'click', l: { Target: url } });
      expect(cell.hyperlinks?.[0].url).toBe(url);
    }
  });

  it('caps an over-long hyperlink target', () => {
    const url = `https://example.com/${'u'.repeat(MAX_CELL_CHARS + 1_000)}`;
    const cell = toCell({ t: 's', v: 'click', w: 'click', l: { Target: url } });
    expect(cell.hyperlinks?.[0].url.length).toBe(MAX_CELL_CHARS);
  });
});

describe('an overlong tab name is reachable by either form', () => {
  // Parses for real: hand-building the workbook is how the mismatch went unnoticed.
  const long = 'T'.repeat(400);
  const file = () => workbook(
    sheetOf('<row r="1"><c r="A1" t="inlineStr"><is><t>hit</t></is></c></row>'),
    [{ name: long, file: 'sheet1.xml' }]
  );

  it('resolves the file\'s own full name', () => {
    const wb = parseXlsx(file(), { sheet: long });
    expect(wb.sheets[0].notRequested).toBeUndefined();
    expect(wb.sheets[0].data[0][0].value).toBe('hit');
  });

  it('resolves the shortened name that get_metadata actually emits', () => {
    const wb = parseXlsx(file(), { sheet: long.slice(0, MAX_SHEET_NAME_CHARS) });
    expect(wb.sheets[0].notRequested).toBeUndefined();
    expect(wb.sheets[0].data[0][0].value).toBe('hit');
  });
});
