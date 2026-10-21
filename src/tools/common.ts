import { z } from 'zod';
import { withGoogleAuth } from '../auth.js';
import { wrapHandler } from '../lib/errors.js';
import { nativeOnly } from '../lib/office.js';
import { SCOPES } from '../scopes.js';

export const spreadsheetIdSchema = z.string().describe('Google Sheets spreadsheet ID');
export const sheetNameSchema = z.string().describe('Name of the sheet tab');

export const optionalSheetNameSchema = z.string().optional()
  .describe('Name of the sheet tab. May be omitted when `range` carries the tab ("Sheet1!A1:C3").');

/**
 * A handler that calls Google with the given scope on a native spreadsheet:
 * it reports failures as tool errors and refuses an uploaded .xlsx with a
 * pointer to convert_to_google_sheet.
 */
export function nativeTool(scope: string, handler: (args: any, accessToken: string) => Promise<unknown>) {
  return withGoogleAuth(scope, wrapHandler(nativeOnly(
    (args: any, context: { accessToken: string }) => handler(args, context.accessToken),
  )));
}

export const nativeEdit = (handler: (args: any, accessToken: string) => Promise<unknown>) =>
  nativeTool(SCOPES.SPREADSHEETS, handler);
