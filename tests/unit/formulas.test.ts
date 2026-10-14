import { describe, expect, it } from 'vitest';
import { extractRefs, refContains } from '../../src/data/formula-refs';
import { goalSeek } from '../../src/data/goal-seek';
import { isoCeiling, trimRange } from '../../src/univer/custom-functions';

describe('références d’une formule', () => {
  it('trouve cellules, plages, feuilles, et ignore les chaînes et noms de fonctions', () => {
    const refs = extractRefs(`=SUM(A1:B3)+'Données 2026'!$C$4*Feuil2!D5+LOG10(2)&"A1"+VLOOKUP(E1,F1:G9,2,FALSE)`);
    expect(refs.map((r) => [r.sheet, r.text])).toEqual([
      [undefined, 'A1:B3'],
      ['Données 2026', "'Données 2026'!$C$4"],
      ['Feuil2', 'Feuil2!D5'],
      [undefined, 'E1'],
      [undefined, 'F1:G9'],
    ]);
    expect(refs[0]).toMatchObject({ startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 });
    expect(refContains(refs[0], 2, 1)).toBe(true);
    expect(refContains(refs[0], 3, 1)).toBe(false);
  });
});

describe('Valeur cible', () => {
  it('résout une équation linéaire et non linéaire', async () => {
    const r1 = await goalSeek((x) => 3 * x + 5, 20, 0);
    expect(r1.found).toBe(true);
    expect(r1.x).toBeCloseTo(5, 6);
    const r2 = await goalSeek((x) => x * x, 2, 1);
    expect(r2.found).toBe(true);
    expect(r2.x).toBeCloseTo(Math.SQRT2, 6);
  });
  it('mensualité d’emprunt : trouve le taux', async () => {
    const pmt = (rate: number) => (10000 * rate) / (1 - Math.pow(1 + rate, -120));
    const r = await goalSeek(pmt, 106.07, 0.01);
    expect(r.found).toBe(true);
    expect(r.x * 12).toBeCloseTo(0.05, 3);
  });
  it('signale l’absence de solution', async () => {
    expect((await goalSeek((x) => x * x + 1, 0, 1)).found).toBe(false);
  });
});

describe('fonctions ajoutées', () => {
  it('ISO.CEILING et TRIMRANGE', () => {
    expect(isoCeiling(-4.3, 2)).toBe(-4);
    expect(isoCeiling(4.3, -2)).toBe(6);
    expect(isoCeiling(4.3)).toBe(5);
    expect(trimRange([[null, null, null], [null, 1, ''], [null, 3, null], [null, null, null]])).toEqual([[1], [3]]);
  });
});
