/**
 * Plan (Grouper / Dissocier) : niveaux de regroupement des lignes ou colonnes, comme Excel
 * (8 niveaux au plus : 0 à 7). Les niveaux sont stockés dans les métadonnées des lignes et
 * colonnes (custom.grillon.outline) et exportés dans le .xlsx (outlineLevel).
 */

export const MAX_OUTLINE_LEVEL = 7;

/** Niveaux après « Grouper » sur [from, to]. */
export function group(levels: number[], from: number, to: number): number[] {
  const out = [...levels];
  for (let i = from; i <= to; i++) out[i] = Math.min(MAX_OUTLINE_LEVEL, (out[i] ?? 0) + 1);
  return out;
}

/** Niveaux après « Dissocier » sur [from, to]. */
export function ungroup(levels: number[], from: number, to: number): number[] {
  const out = [...levels];
  for (let i = from; i <= to; i++) out[i] = Math.max(0, (out[i] ?? 0) - 1);
  return out;
}

/**
 * Bloc de détail qui contient `index` : la plage contiguë de lignes de niveau ≥ niveau(index).
 * Renvoie null si la ligne n'est dans aucun groupe.
 */
export function detailBlock(levels: number[], index: number): { from: number; to: number; level: number } | null {
  const level = levels[index] ?? 0;
  if (level <= 0) return null;
  let from = index;
  while (from > 0 && (levels[from - 1] ?? 0) >= level) from--;
  let to = index;
  while (to + 1 < levels.length && (levels[to + 1] ?? 0) >= level) to++;
  return { from, to, level };
}

/** Lignes à masquer pour n'afficher que les niveaux ≤ `show` (boutons 1, 2, 3… d'Excel). */
export function hiddenForLevel(levels: number[], show: number): boolean[] {
  return levels.map((l) => (l ?? 0) >= show);
}

/** Tous les groupes du plan (un par niveau et par bloc contigu), comme les crochets d'Excel. */
export function outlineBlocks(levels: number[]): { from: number; to: number; level: number }[] {
  const out: { from: number; to: number; level: number }[] = [];
  const max = levels.reduce((m, l) => Math.max(m, l ?? 0), 0);
  for (let level = 1; level <= max; level++) {
    let start = -1;
    for (let i = 0; i <= levels.length; i++) {
      const inside = i < levels.length && (levels[i] ?? 0) >= level;
      if (inside && start < 0) start = i;
      if (!inside && start >= 0) {
        out.push({ from: start, to: i - 1, level });
        start = -1;
      }
    }
  }
  return out;
}

/**
 * Lignes à afficher quand on développe un groupe (bouton +) : celles du niveau du groupe ;
 * les sous-groupes restent réduits, comme dans Excel.
 */
export function rowsToExpand(levels: number[], block: { from: number; to: number; level: number }): number[] {
  const out: number[] = [];
  for (let i = block.from; i <= block.to; i++) if ((levels[i] ?? 0) === block.level) out.push(i);
  return out;
}
