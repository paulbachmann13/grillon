/**
 * « Zone de données en cours » (Ctrl+A / CurrentRegion d'Excel) : le plus petit rectangle
 * autour de la sélection, bordé de lignes et de colonnes entièrement vides.
 *
 * (getDataRegion() d'Univer n'étend la plage que d'une ligne/colonne à la fois.)
 */
import type { FWorksheet } from '@univerjs/sheets/facade';
import { isEmpty, type Grid } from '../data/cells';

export interface Rect {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

const ROW_CHUNK = 2000;
const COL_CHUNK = 50;

function read(ws: FWorksheet, r1: number, c1: number, r2: number, c2: number): Grid {
  if (r2 < r1 || c2 < c1) return [];
  return ws.getRange(r1, c1, r2 - r1 + 1, c2 - c1 + 1).getCellDataGrid() as Grid;
}

/** Étend un rectangle jusqu'aux bords vides (lecture par blocs pour rester rapide). */
export function currentRegionRect(ws: FWorksheet, start: Rect): Rect {
  const maxRow = ws.getMaxRows() - 1;
  const maxCol = ws.getMaxColumns() - 1;
  const rect = { ...start };
  const rowEmpty = (g: Grid, i: number) => (g[i] ?? []).every(isEmpty);
  const colEmpty = (g: Grid, j: number) => g.every((row) => isEmpty(row[j]));
  let changed = true;
  while (changed) {
    changed = false;
    const cl = Math.max(0, rect.c1 - 1);
    const cr = Math.min(maxCol, rect.c2 + 1);
    // Vers le bas.
    while (rect.r2 < maxRow) {
      const to = Math.min(maxRow, rect.r2 + ROW_CHUNK);
      const g = read(ws, rect.r2 + 1, cl, to, cr);
      const firstEmpty = g.findIndex((_, i) => rowEmpty(g, i));
      if (firstEmpty === 0) break;
      rect.r2 = firstEmpty < 0 ? to : rect.r2 + firstEmpty;
      changed = true;
      if (firstEmpty >= 0) break;
    }
    // Vers le haut.
    while (rect.r1 > 0) {
      const from = Math.max(0, rect.r1 - ROW_CHUNK);
      const g = read(ws, from, cl, rect.r1 - 1, cr);
      let k = g.length - 1;
      while (k >= 0 && !rowEmpty(g, k)) k--;
      if (k === g.length - 1) break;
      rect.r1 = from + k + 1;
      changed = true;
      if (k >= 0) break;
    }
    const rt = Math.max(0, rect.r1 - 1);
    const rb = Math.min(maxRow, rect.r2 + 1);
    // Vers la droite.
    while (rect.c2 < maxCol) {
      const to = Math.min(maxCol, rect.c2 + COL_CHUNK);
      const g = read(ws, rt, rect.c2 + 1, rb, to);
      const width = to - rect.c2;
      let j = 0;
      while (j < width && !colEmpty(g, j)) j++;
      if (j === 0) break;
      rect.c2 += j;
      changed = true;
      if (j < width) break;
    }
    // Vers la gauche.
    while (rect.c1 > 0) {
      const from = Math.max(0, rect.c1 - COL_CHUNK);
      const g = read(ws, rt, from, rb, rect.c1 - 1);
      const width = rect.c1 - from;
      let j = width - 1;
      while (j >= 0 && !colEmpty(g, j)) j--;
      if (j === width - 1) break;
      rect.c1 = from + j + 1;
      changed = true;
      if (j >= 0) break;
    }
  }
  return rect;
}
