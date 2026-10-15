import { describe, expect, it } from 'vitest';
import { autoSumPlan, columnName } from '../../src/data/autosum';

// Grille de test : colonne A = 1,2,3 (lignes 0 à 2) ; ligne 5 = 10, 20 (colonnes 0 à 1).
const numbers = new Set(['0,0', '1,0', '2,0', '5,0', '5,1']);
const isNumber = (r: number, c: number) => numbers.has(`${r},${c}`);

describe('somme automatique', () => {
  it('nomme les colonnes comme Excel', () => {
    expect([0, 25, 26, 51, 701, 702, 16383].map(columnName)).toEqual(['A', 'Z', 'AA', 'AZ', 'ZZ', 'AAA', 'XFD']);
  });

  it('une cellule : somme les nombres au-dessus', () => {
    expect(autoSumPlan(isNumber, { r1: 3, c1: 0, r2: 3, c2: 0 })).toEqual([{ row: 3, col: 0, formula: '=SUM(A1:A3)' }]);
    expect(autoSumPlan(isNumber, { r1: 3, c1: 0, r2: 3, c2: 0 }, 'AVERAGE')[0]!.formula).toBe('=AVERAGE(A1:A3)');
  });

  it('une cellule : à défaut, somme les nombres à gauche', () => {
    expect(autoSumPlan(isNumber, { r1: 5, c1: 2, r2: 5, c2: 2 })).toEqual([{ row: 5, col: 2, formula: '=SUM(A6:B6)' }]);
  });

  it('une cellule : un seul nombre au-dessus', () => {
    expect(autoSumPlan(isNumber, { r1: 6, c1: 0, r2: 6, c2: 0 })[0]!.formula).toBe('=SUM(A6)');
  });

  it('une cellule sans nombres autour : rien', () => {
    expect(autoSumPlan(isNumber, { r1: 9, c1: 5, r2: 9, c2: 5 })).toEqual([]);
  });

  it('plusieurs lignes : un total sous chaque colonne', () => {
    expect(autoSumPlan(isNumber, { r1: 0, c1: 0, r2: 2, c2: 1 }, 'MAX')).toEqual([
      { row: 3, col: 0, formula: '=MAX(A1:A3)' },
      { row: 3, col: 1, formula: '=MAX(B1:B3)' },
    ]);
  });

  it('une ligne : le total est à droite', () => {
    expect(autoSumPlan(isNumber, { r1: 5, c1: 0, r2: 5, c2: 1 })).toEqual([{ row: 5, col: 2, formula: '=SUM(A6:B6)' }]);
  });

  it("totaux à droite de chaque ligne (Analyse rapide)", () => {
    expect(autoSumPlan(isNumber, { r1: 0, c1: 0, r2: 2, c2: 0 }, "SUM", "right")).toEqual([
      { row: 0, col: 1, formula: "=SUM(A1)" },
      { row: 1, col: 1, formula: "=SUM(A2)" },
      { row: 2, col: 1, formula: "=SUM(A3)" },
    ]);
  });
});
