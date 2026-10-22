import { describe, expect, it } from 'vitest';
import { sortFormulaFixes, type FormulaCell } from '../../src/data/sort-formulas';

const range = { startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 };
// Colonne A : valeurs ; colonne B : =A{n}*10 (lignes 1 à 4).
const formulaAt = (row: number, col: number): FormulaCell | null => (col === 1 ? { formula: `=A${row + 1}*10`, f: `=A${row + 1}*10`, si: null } : null);

describe('tri avec formules (comme Excel)', () => {
  it('chaque formule suit sa ligne : ses références relatives sont décalées', () => {
    // Tri décroissant de 5, 3, 8, 1 : la ligne 3 (8) passe en 1, la ligne 1 (5) en 2, la ligne 2 (3) en 3, la ligne 4 reste.
    const order = { 0: 2, 1: 0, 2: 1, 3: 3 };
    const { redo } = sortFormulaFixes(order, range, formulaAt);
    expect([0, 1, 2, 3].map((r) => redo[r][1].f)).toEqual(['=A1*10', '=A2*10', '=A3*10', '=A4*10']);
    expect(redo[0][1].si).toBeNull();
    expect(redo[0][0]).toBeUndefined();
  });

  it('les références absolues ne bougent pas, les autres feuilles non plus', () => {
    const at = (row: number, col: number): FormulaCell | null => (col === 1 ? { formula: `=$A$${row + 1}+Feuil2!A${row + 1}+$D$1`, f: null, si: 'g' } : null);
    const { redo } = sortFormulaFixes({ 0: 2, 2: 0 }, range, at);
    expect(redo[0][1].f).toBe('=$A$3+Feuil2!A1+$D$1');
    expect(redo[2][1].f).toBe('=$A$1+Feuil2!A3+$D$1');
  });

  it('l’annulation rétablit les champs bruts (formule partagée comprise)', () => {
    const at = (row: number, col: number): FormulaCell | null => (col === 1 ? { formula: `=A${row + 1}*10`, f: row === 0 ? '=A1*10' : null, si: 'g' } : null);
    const { undo } = sortFormulaFixes({ 0: 1, 1: 0 }, range, at);
    expect(undo[0][1]).toEqual({ f: null, si: 'g' });
    expect(undo[1][1]).toEqual({ f: '=A1*10', si: 'g' });
  });

  it('sans formule, rien à faire', () => {
    expect(sortFormulaFixes({ 0: 1, 1: 0 }, range, () => null)).toEqual({ redo: {}, undo: {} });
  });
});
