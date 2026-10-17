/**
 * Cellules déverrouillées (Format de cellule → Protection, comme Excel) : par défaut toutes les
 * cellules sont verrouillées ; une fois la feuille protégée, seules les cellules déverrouillées
 * restent modifiables. On les garde sous forme de rectangles disjoints.
 */

export interface Rect {
  startRow: number;
  endRow: number;
  startColumn: number;
  endColumn: number;
}

const intersects = (a: Rect, b: Rect) => a.startRow <= b.endRow && b.startRow <= a.endRow && a.startColumn <= b.endColumn && b.startColumn <= a.endColumn;

/** a privé de b : au plus quatre rectangles (bandes du haut et du bas, puis côtés). */
export function subtract(a: Rect, b: Rect): Rect[] {
  if (!intersects(a, b)) return [a];
  const out: Rect[] = [];
  if (a.startRow < b.startRow) out.push({ ...a, endRow: b.startRow - 1 });
  if (a.endRow > b.endRow) out.push({ ...a, startRow: b.endRow + 1 });
  const top = Math.max(a.startRow, b.startRow);
  const bottom = Math.min(a.endRow, b.endRow);
  if (a.startColumn < b.startColumn) out.push({ startRow: top, endRow: bottom, startColumn: a.startColumn, endColumn: b.startColumn - 1 });
  if (a.endColumn > b.endColumn) out.push({ startRow: top, endRow: bottom, startColumn: b.endColumn + 1, endColumn: a.endColumn });
  return out;
}

/** Verrouille (unlocked = false) ou déverrouille une plage. */
export function setUnlocked(rects: Rect[], range: Rect, unlocked: boolean): Rect[] {
  const rest = rects.flatMap((r) => subtract(r, range));
  return unlocked ? [...rest, { ...range }] : rest;
}

/** Vrai si toutes les cellules de la plage sont déverrouillées. */
export function isRangeUnlocked(rects: Rect[], range: Rect): boolean {
  let left: Rect[] = [range];
  for (const r of rects) {
    left = left.flatMap((x) => subtract(x, r));
    if (!left.length) return true;
  }
  return left.length === 0;
}

/** État de la case « Verrouillée » pour une sélection : true, false ou null (mixte). */
export function lockedState(rects: Rect[], range: Rect): boolean | null {
  if (isRangeUnlocked(rects, range)) return false;
  return rects.some((r) => intersects(r, range)) ? null : true;
}

/** Regroupe des cellules [ligne, colonne] en rectangles (séries par ligne, puis fusion verticale). */
export function cellsToRects(cells: [number, number][]): Rect[] {
  const byRow = new Map<number, number[]>();
  for (const [r, c] of cells) (byRow.get(r) ?? byRow.set(r, []).get(r)!).push(c);
  const open = new Map<string, Rect>();
  const done: Rect[] = [];
  for (const r of [...byRow.keys()].sort((a, b) => a - b)) {
    const cols = [...new Set(byRow.get(r))].sort((a, b) => a - b);
    const runs: [number, number][] = [];
    for (const c of cols) {
      const last = runs[runs.length - 1];
      if (last && last[1] === c - 1) last[1] = c;
      else runs.push([c, c]);
    }
    const next = new Map<string, Rect>();
    for (const [s, e] of runs) {
      const key = `${s}:${e}`;
      const prev = open.get(key);
      if (prev && prev.endRow === r - 1) {
        prev.endRow = r;
        next.set(key, prev);
        open.delete(key);
      } else next.set(key, { startRow: r, endRow: r, startColumn: s, endColumn: e });
    }
    done.push(...open.values());
    open.clear();
    for (const [k, v] of next) open.set(k, v);
  }
  return [...done, ...open.values()];
}
