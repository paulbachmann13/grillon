/**
 * Tri d'une plage qui contient des formules : comme Excel, chaque formule suit sa ligne et ses références relatives
 * sont décalées du déplacement de la ligne (une formule « =A1*10 » qui passe de la ligne 1 à la ligne 3 devient « =A3*10 »).
 * Univer, lui, déplace les cellules sans toucher aux formules : les résultats deviennent faux.
 * Logique pure : on calcule les formules à écrire après le tri et celles à rétablir pour l'annulation.
 */
import { shiftFormula } from './formula-shift';

export interface SortRange {
  startRow: number;
  endRow: number;
  startColumn: number;
  endColumn: number;
}

export interface FormulaCell {
  /** Formule réelle de la cellule (celle de la première cellule pour une formule partagée, décalée). */
  formula: string;
  /** Champs bruts enregistrés (pour les rétablir à l'annulation). */
  f: string | null;
  si: string | null;
}

export type CellPatch = Record<number, Record<number, { f: string | null; si: string | null }>>;

/**
 * `order` : ligne de destination → ligne d'origine (celui d'Univer). `formulaAt(row, col)` : la formule de la cellule
 * avant le tri, ou null. Renvoie les modifications à appliquer après le tri (`redo`) et avant de l'annuler (`undo`).
 */
export function sortFormulaFixes(order: Record<number, number>, range: SortRange, formulaAt: (row: number, col: number) => FormulaCell | null): { redo: CellPatch; undo: CellPatch } {
  const redo: CellPatch = {};
  const undo: CellPatch = {};
  for (let target = range.startRow; target <= range.endRow; target++) {
    const source = Object.prototype.hasOwnProperty.call(order, target) ? order[target] : target;
    for (let col = range.startColumn; col <= range.endColumn; col++) {
      const cell = formulaAt(source, col);
      if (!cell) continue;
      // Chaque formule devient explicite (une formule partagée perd son lien : sa première cellule a pu bouger).
      (redo[target] ??= {})[col] = { f: shiftFormula(cell.formula, target - source, 0), si: null };
      // À l'annulation, la cellule retrouve d'abord ce qu'elle contenait juste après le déplacement, puis Univer la remet en place.
      (undo[target] ??= {})[col] = { f: cell.f, si: cell.si };
    }
  }
  return { redo, undo };
}
