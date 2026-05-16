/**
 * Color parsing for Google Sheets API.
 */

export interface RgbFloatColor {
  red?: number;
  green?: number;
  blue?: number;
}

/**
 * Parse a color input into Google Sheets' RGB float form ({red,green,blue}, 0..1).
 * Accepts:
 *   - A hex string: "#FF0000", "FF0000", "#F00", "F00" (alpha not supported)
 *   - An object with red/green/blue floats in 0..1 (passthrough)
 * Returns undefined if the input is undefined or null.
 */
export function parseColor(input: unknown): RgbFloatColor | undefined {
  if (input === undefined || input === null) return undefined;
  if (typeof input === 'string') {
    const hex = input.trim().replace(/^#/, '');
    let r: number;
    let g: number;
    let b: number;
    if (/^[0-9a-fA-F]{3}$/.test(hex)) {
      r = parseInt(hex[0] + hex[0], 16);
      g = parseInt(hex[1] + hex[1], 16);
      b = parseInt(hex[2] + hex[2], 16);
    } else if (/^[0-9a-fA-F]{6}$/.test(hex)) {
      r = parseInt(hex.slice(0, 2), 16);
      g = parseInt(hex.slice(2, 4), 16);
      b = parseInt(hex.slice(4, 6), 16);
    } else {
      throw new Error(
        `Invalid hex color: "${input}". Use "#RRGGBB", "#RGB", or an {red,green,blue} object with floats 0..1.`,
      );
    }
    return { red: r / 255, green: g / 255, blue: b / 255 };
  }
  if (typeof input === 'object') {
    return input as RgbFloatColor;
  }
  throw new Error('Color must be a hex string (e.g. "#FF0000") or an {red,green,blue} object with floats 0..1.');
}
