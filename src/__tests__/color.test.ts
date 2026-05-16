import { describe, it, expect } from 'vitest';
import { parseColor } from '../lib/color.js';

describe('parseColor', () => {
  it('returns undefined for null/undefined', () => {
    expect(parseColor(undefined)).toBeUndefined();
    expect(parseColor(null)).toBeUndefined();
  });

  it('parses 6-digit hex with leading #', () => {
    const c = parseColor('#FF0000');
    expect(c?.red).toBeCloseTo(1, 5);
    expect(c?.green).toBeCloseTo(0, 5);
    expect(c?.blue).toBeCloseTo(0, 5);
  });

  it('parses 6-digit hex without leading #', () => {
    const c = parseColor('00FF00');
    expect(c?.red).toBeCloseTo(0, 5);
    expect(c?.green).toBeCloseTo(1, 5);
    expect(c?.blue).toBeCloseTo(0, 5);
  });

  it('parses 3-digit hex with leading #', () => {
    const c = parseColor('#F00');
    expect(c?.red).toBeCloseTo(1, 5);
    expect(c?.green).toBeCloseTo(0, 5);
    expect(c?.blue).toBeCloseTo(0, 5);
  });

  it('parses 3-digit hex without leading #', () => {
    const c = parseColor('0F0');
    expect(c?.red).toBeCloseTo(0, 5);
    expect(c?.green).toBeCloseTo(1, 5);
    expect(c?.blue).toBeCloseTo(0, 5);
  });

  it('trims whitespace around hex', () => {
    const c = parseColor('  #0000FF  ');
    expect(c?.red).toBeCloseTo(0, 5);
    expect(c?.green).toBeCloseTo(0, 5);
    expect(c?.blue).toBeCloseTo(1, 5);
  });

  it('is case-insensitive', () => {
    const upper = parseColor('#ABCDEF');
    const lower = parseColor('#abcdef');
    expect(upper).toEqual(lower);
  });

  it('passes through RGB float objects', () => {
    const input = { red: 0.5, green: 0.25, blue: 0.75 };
    expect(parseColor(input)).toEqual(input);
  });

  it('throws on too-short hex', () => {
    expect(() => parseColor('#FF')).toThrow(/Invalid hex color/);
  });

  it('throws on too-long hex', () => {
    expect(() => parseColor('#FF00FF00')).toThrow(/Invalid hex color/);
  });

  it('throws on non-hex characters', () => {
    expect(() => parseColor('#GGGGGG')).toThrow(/Invalid hex color/);
    expect(() => parseColor('zzz')).toThrow(/Invalid hex color/);
  });

  it('throws on unsupported scalar types', () => {
    expect(() => parseColor(42 as unknown)).toThrow();
    expect(() => parseColor(true as unknown)).toThrow();
  });
});
