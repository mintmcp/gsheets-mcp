import { readFileSync } from 'node:fs';

/** Loads a checked-in .xlsx fixture as the bytes parseXlsx receives. */
export const fixture = (name: string) =>
  new Uint8Array(readFileSync(new URL(`./${name}`, import.meta.url)));
