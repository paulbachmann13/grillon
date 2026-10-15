/**
 * Somme automatique (Σ, Alt+=) : logique pure, comme dans Excel.
 * - Une seule cellule : somme des nombres juste au-dessus ; sinon juste à gauche.
 * - Une ligne de plusieurs cellules : le total est écrit à droite de la sélection.
 * - Plusieurs lignes : un total par colonne, sous la sélection.
 */
export type AutoSumFunction = 'SUM' | 'AVERAGE' | 'COUNT' | 'MAX' | 'MIN';

export interface Rect {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

export interface AutoSumWrite {
  row: number;
  col: number;
  formula: string;
}

export function columnName(index: number): string {
  let name = '';
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) name = String.fromCharCode(65 + ((n - 1) % 26)) + name;
  return name;
}

const a1 = (row: number, col: number): string => `${columnName(col)}${row + 1}`;
const rangeA1 = (r1: number, c1: number, r2: number, c2: number): string => (r1 === r2 && c1 === c2 ? a1(r1, c1) : `${a1(r1, c1)}:${a1(r2, c2)}`);

export function autoSumPlan(isNumber: (row: number, col: number) => boolean, selection: Rect, fn: AutoSumFunction = 'SUM', direction: 'down' | 'right' | 'auto' = 'auto'): AutoSumWrite[] {
  const { r1, c1, r2, c2 } = selection;
  const make = (row: number, col: number, from: Rect): AutoSumWrite => ({ row, col, formula: `=${fn}(${rangeA1(from.r1, from.c1, from.r2, from.c2)})` });

  if (r1 === r2 && c1 === c2) {
    let top = r1;
    while (top > 0 && isNumber(top - 1, c1)) top--;
    if (top < r1) return [make(r1, c1, { r1: top, c1, r2: r1 - 1, c2: c1 })];
    let left = c1;
    while (left > 0 && isNumber(r1, left - 1)) left--;
    if (left < c1) return [make(r1, c1, { r1, c1: left, r2: r1, c2: c1 - 1 })];
    return [];
  }
  if (direction === 'right') {
    const rightTotals: AutoSumWrite[] = [];
    for (let r = r1; r <= r2; r++) rightTotals.push(make(r, c2 + 1, { r1: r, c1, r2: r, c2 }));
    return rightTotals;
  }
  if (r1 === r2) return [make(r1, c2 + 1, selection)];
  const out: AutoSumWrite[] = [];
  for (let c = c1; c <= c2; c++) out.push(make(r2 + 1, c, { r1, c1: c, r2, c2: c }));
  return out;
}
