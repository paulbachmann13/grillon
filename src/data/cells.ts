/**
 * Utilitaires communs aux outils de données (logique pure, sans Univer à l'exécution).
 */
import { CellValueType, type ICellData } from '../model/univer-types';

export type Cell = ICellData | null | undefined;
export type Grid = Cell[][];

/** Valeur « brute » d'une cellule : nombre, booléen, texte, ou null si vide. */
export function cellValue(cell: Cell): string | number | boolean | null {
  if (!cell) return null;
  const v = cell.v;
  if (v === null || v === undefined || v === '') {
    const text = cell.p?.body?.dataStream?.replace(/\r?\n$/, '').replace(/\r$/, '');
    return text ? text : null;
  }
  if (cell.t === CellValueType.BOOLEAN) return v === true || v === 1 || v === '1' || String(v).toUpperCase() === 'TRUE';
  if (typeof v === 'number') return v;
  if (cell.t === CellValueType.NUMBER && v !== '' && Number.isFinite(Number(v))) return Number(v);
  return String(v);
}

/** Texte d'une cellule tel qu'un utilisateur le lit (sans format de nombre). */
export function cellText(cell: Cell): string {
  const v = cellValue(cell);
  if (v === null) return '';
  if (typeof v === 'boolean') return v ? 'VRAI' : 'FAUX';
  return String(v);
}

/** Clé de comparaison « à la Excel » : texte insensible à la casse et aux espaces de bord. */
export function compareKey(cell: Cell): string {
  const v = cellValue(cell);
  if (v === null) return '∅';
  if (typeof v === 'number') return `n:${v}`;
  if (typeof v === 'boolean') return `b:${v}`;
  const n = Number(v.replace(',', '.'));
  if (v.trim() !== '' && Number.isFinite(n) && /^\s*-?\d/.test(v)) return `n:${n}`;
  return `s:${v.trim().toLocaleLowerCase('fr')}`;
}

export function isEmpty(cell: Cell): boolean {
  return cellValue(cell) === null && !cell?.f;
}
