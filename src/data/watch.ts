/**
 * Fenêtre Espion (Formules) : liste de cellules surveillées, comme dans Excel.
 * Logique pure : l'interface lit les valeurs dans la feuille.
 */
export interface WatchItem {
  sheetId: string;
  sheetName: string;
  row: number;
  col: number;
}

const same = (a: WatchItem, b: WatchItem) => a.sheetId === b.sheetId && a.row === b.row && a.col === b.col;

/** Ajoute une cellule (sans doublon). */
export function addWatch(list: readonly WatchItem[], item: WatchItem): WatchItem[] {
  return list.some((w) => same(w, item)) ? [...list] : [...list, item];
}

export function removeWatch(list: readonly WatchItem[], item: WatchItem): WatchItem[] {
  return list.filter((w) => !same(w, item));
}

/** Met à jour le nom d'une feuille renommée. */
export function renameSheet(list: readonly WatchItem[], sheetId: string, sheetName: string): WatchItem[] {
  return list.map((w) => (w.sheetId === sheetId ? { ...w, sheetName } : w));
}

/** Retire les cellules d'une feuille supprimée. */
export function dropSheet(list: readonly WatchItem[], sheetId: string): WatchItem[] {
  return list.filter((w) => w.sheetId !== sheetId);
}

/** Tri d'une colonne de la liste (texte en ordre naturel français, nombres en ordre numérique). */
export function sortByColumn<T>(rows: readonly T[], key: (row: T) => string | number, descending = false): T[] {
  const sign = descending ? -1 : 1;
  return [...rows].sort((a, b) => {
    const x = key(a);
    const y = key(b);
    if (typeof x === 'number' && typeof y === 'number') return sign * (x - y);
    return sign * String(x).localeCompare(String(y), 'fr', { numeric: true });
  });
}
