/**
 * Conversions d'unités et de références entre Excel et Univer.
 *
 * - Largeur de colonne Excel : en « caractères » (largeur du chiffre 0 en
 *   Calibri 11 = 7 px, plus 5 px de marge). Univer : en pixels.
 * - Hauteur de ligne Excel : en points. Univer : en pixels (96 dpi).
 * - Dates Excel : numéro de série depuis le 30/12/1899 (système 1900).
 */

const MAX_DIGIT_WIDTH = 7;
const COLUMN_PADDING = 5;

export function excelWidthToPx(chars: number): number {
  return Math.round(chars * MAX_DIGIT_WIDTH + COLUMN_PADDING);
}

export function pxToExcelWidth(px: number): number {
  return Math.max(0, Math.round(((px - COLUMN_PADDING) / MAX_DIGIT_WIDTH) * 100) / 100);
}

export function ptToPx(pt: number): number {
  return Math.round((pt * 96) / 72);
}

export function pxToPt(px: number): number {
  return Math.round(((px * 72) / 96) * 100) / 100;
}

/** Largeur Excel par défaut (8,43 caractères = 64 px) et hauteur par défaut (15 pt = 20 px). */
export const EXCEL_DEFAULT_COL_WIDTH = 8.43;
export const EXCEL_DEFAULT_ROW_HEIGHT_PT = 15;

const MS_PER_DAY = 86_400_000;
const EXCEL_EPOCH_UTC = Date.UTC(1899, 11, 30);

/** Date JS (interprétée en UTC, comme ExcelJS la renvoie) → numéro de série Excel. */
export function dateToSerial(date: Date): number {
  const serial = (date.getTime() - EXCEL_EPOCH_UTC) / MS_PER_DAY;
  // Arrondi à la milliseconde pour éviter les 0,99999999 dus aux flottants.
  return Math.round(serial * MS_PER_DAY) / MS_PER_DAY;
}

export function serialToDate(serial: number): Date {
  return new Date(EXCEL_EPOCH_UTC + Math.round(serial * MS_PER_DAY));
}

/** 0 → A, 25 → Z, 26 → AA. */
export function columnToLetters(col: number): string {
  let n = col + 1;
  let s = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** A → 0, AA → 26. */
export function lettersToColumn(letters: string): number {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

/** (0, 0) → A1. */
export function toA1(row: number, col: number): string {
  return `${columnToLetters(col)}${row + 1}`;
}

/** « $B$3 » ou « B3 » → { row: 2, col: 1 }. */
export function fromA1(ref: string): { row: number; col: number } {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(ref.trim());
  if (!m) throw new Error(`Référence invalide : ${ref}`);
  return { row: Number(m[2]) - 1, col: lettersToColumn(m[1]) };
}

/** « A1:C3 » → plage (indices 0). */
export function rangeFromA1(ref: string): { startRow: number; endRow: number; startColumn: number; endColumn: number } {
  const [a, b = a] = ref.split(':');
  const p = fromA1(a);
  const q = fromA1(b);
  return {
    startRow: Math.min(p.row, q.row),
    endRow: Math.max(p.row, q.row),
    startColumn: Math.min(p.col, q.col),
    endColumn: Math.max(p.col, q.col),
  };
}

export function rangeToA1(r: { startRow: number; endRow: number; startColumn: number; endColumn: number }): string {
  const a = toA1(r.startRow, r.startColumn);
  const b = toA1(r.endRow, r.endColumn);
  return a === b ? a : `${a}:${b}`;
}
