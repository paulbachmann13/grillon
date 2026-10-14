import { describe, expect, it } from 'vitest';
import { formatGeneralFr } from '../../src/univer/french-display';

describe('format Standard à la française', () => {
  it('se comporte comme Excel en français', () => {
    expect(formatGeneralFr(1.5)).toBe('1,5');
    expect(formatGeneralFr(0.1 + 0.2)).toBe('0,3');
    expect(formatGeneralFr(10 / 3)).toBe('3,333333333');
    expect(formatGeneralFr(1 / 3)).toBe('0,333333333');
    expect(formatGeneralFr(-100 / 3)).toBe('-33,3333333');
    expect(formatGeneralFr(42)).toBe('42');
    expect(formatGeneralFr(12345678901)).toBe('12345678901');
    expect(formatGeneralFr(123456789012)).toBe('1,23457E+11');
    expect(formatGeneralFr(0.0000000001)).toBe('1E-10');
    expect(formatGeneralFr(0)).toBe('0');
  });
});
