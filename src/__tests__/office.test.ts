import { describe, it, expect } from 'vitest';
import { ApiError } from '../lib/errors.js';
import {
  driveFileKind,
  isOfficeFileError,
  officeFileMessage,
  availableTabs,
  xlsxSheetOutput,
  xlsxMetadataOutput,
  XLSX_MIME,
  XLS_MIME,
  NATIVE_SHEET_MIME,
} from '../lib/office.js';
import { parseXlsx } from '../lib/xlsx.js';
import { fixture } from './__fixtures__/index.js';

const sheetsError = (
  message: string,
  status = 400,
  reason?: string,
) => new ApiError(message, status, 'sheets', undefined, reason);

describe('driveFileKind', () => {
  it('reports xlsx uploads as xlsx', () => {
    expect(driveFileKind(XLSX_MIME)).toBe('xlsx');
  });

  it('reports native sheets and anything else as native', () => {
    expect(driveFileKind(NATIVE_SHEET_MIME)).toBe('native');
    expect(driveFileKind('')).toBe('native');
  });
});

describe('isOfficeFileError', () => {
  it('matches the Sheets 400 that means "this id is an Office upload"', () => {
    expect(isOfficeFileError(sheetsError('The input must not be an Office file.'))).toBe(true);
  });

  it('matches on the FAILED_PRECONDITION reason even if the wording changes', () => {
    expect(isOfficeFileError(sheetsError('some other text', 400, 'FAILED_PRECONDITION'))).toBe(true);
  });

  it('does NOT match an unrelated 400, so a bad range still surfaces as itself', () => {
    expect(isOfficeFileError(sheetsError('Invalid range: A1:ZZZZ9'))).toBe(false);
  });

  it('does not match non-400 statuses or non-ApiError values', () => {
    expect(isOfficeFileError(sheetsError('must not be an Office file', 403))).toBe(false);
    expect(isOfficeFileError(new Error('must not be an Office file'))).toBe(false);
    expect(isOfficeFileError(undefined)).toBe(false);
  });
});

describe('officeFileMessage', () => {
  it('points .xlsx callers at convert_to_google_sheet', () => {
    const msg = officeFileMessage('Budget.xlsx', XLSX_MIME, 'https://x/1');
    expect(msg).toContain('convert_to_google_sheet');
    expect(msg).toContain('https://x/1');
  });

  it('tells .xls callers to re-save, since this connector cannot read it', () => {
    const msg = officeFileMessage('Old.xls', XLS_MIME, 'https://x/2');
    expect(msg).toContain('legacy Excel');
    expect(msg).not.toContain('convert_to_google_sheet');
  });
});

describe('xlsx tool output shaping', () => {
  const wb = parseXlsx(fixture('basic.xlsx'));

  it('reports kind xlsx so the caller knows it is read-only', () => {
    const out = xlsxSheetOutput('id1', wb, undefined);
    expect(out.kind).toBe('xlsx');
    expect(out.id).toBe('id1');
  });

  it('defaults to the first tab when no name is given', () => {
    expect(xlsxSheetOutput('id1', wb, undefined).sheetName).toBe(wb.sheets[0].name);
  });

  it('resolves a tab by name', () => {
    const name = wb.sheets[0].name;
    expect(xlsxSheetOutput('id1', wb, name).sheetName).toBe(name);
  });

  it('throws a message listing real tabs when the name is unknown', () => {
    expect(() => xlsxSheetOutput('id1', wb, 'NoSuchTab')).toThrow(/not found/);
    expect(() => xlsxSheetOutput('id1', wb, 'NoSuchTab')).toThrow(/Available tabs/);
  });

  it('refuses a tab that was skipped because another was requested', () => {
    const scoped = parseXlsx(fixture('basic.xlsx'), { sheet: 0 });
    const skipped = scoped.sheets.find((s) => s.notRequested);
    if (!skipped) return; // single-tab fixture: nothing to assert
    expect(() => xlsxSheetOutput('id1', scoped, skipped.name)).toThrow(/Ask for this tab by name/);
  });

  it('reports metadata with kind xlsx and one entry per tab', () => {
    const names = parseXlsx(fixture('basic.xlsx'), { namesOnly: true });
    const out = xlsxMetadataOutput('id1', 'Budget.xlsx', 'https://x/1', names);
    expect(out.kind).toBe('xlsx');
    expect(out.title).toBe('Budget.xlsx');
    expect(out.sheets).toHaveLength(names.sheets.length);
    expect(out.sheets.map((s) => s.index)).toEqual(names.sheets.map((_, i) => i));
  });

  it('labels tab listings as file content, not instructions', () => {
    expect(availableTabs(wb)).toContain('not instructions');
  });
});
