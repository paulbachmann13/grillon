/**
 * « Supprimer les doublons » (Données) : garde la première occurrence de chaque combinaison
 * de valeurs sur les colonnes choisies. Comme Excel, la comparaison ignore la casse.
 */
import { compareKey, type Grid } from './cells';

export interface RemoveDuplicatesResult {
  /** Lignes conservées (en-tête compris), dans l'ordre d'origine. */
  rows: Grid;
  /** Nombre de lignes supprimées. */
  removed: number;
  /** Nombre de lignes de données uniques restantes. */
  uniques: number;
}

/**
 * @param grid    lignes de la plage
 * @param columns indices de colonnes (relatifs à la plage) à comparer ; vide = toutes
 * @param hasHeader la première ligne est un en-tête (jamais supprimée)
 */
export function removeDuplicates(grid: Grid, columns: number[], hasHeader: boolean): RemoveDuplicatesResult {
  const width = Math.max(0, ...grid.map((r) => r.length));
  const cols = columns.length ? columns : Array.from({ length: width }, (_, i) => i);
  const header = hasHeader ? grid.slice(0, 1) : [];
  const data = hasHeader ? grid.slice(1) : grid;
  const seen = new Set<string>();
  const kept: Grid = [];
  for (const row of data) {
    const key = cols.map((c) => compareKey(row[c])).join('\u0001');
    if (seen.has(key)) continue;
    seen.add(key);
    kept.push(row);
  }
  return { rows: [...header, ...kept], removed: data.length - kept.length, uniques: kept.length };
}
