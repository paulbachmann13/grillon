/**
 * Références d'une formule (audit : antécédents / dépendants).
 */
import { fromA1 } from '../io/xlsx/units';

export interface FormulaRef {
  /** Nom de feuille (undefined = même feuille). */
  sheet?: string;
  startRow: number;
  endRow: number;
  startColumn: number;
  endColumn: number;
  /** Texte de la référence telle qu'écrite. */
  text: string;
}

const SHEET = `(?:'((?:[^']|'')+)'|([\\p{L}_][\\p{L}\\p{N}_.]*))!`;
const CELL = `\\$?[A-Za-z]{1,3}\\$?\\d+`;
const REF_RE = new RegExp(`(?<![\\p{L}\\p{N}_.\\]])(?:${SHEET})?(${CELL}(?::${CELL})?)(?![\\p{L}\\p{N}_(\\[])`, 'gu');

/** Retire les chaînes "…" d'une formule (on ne cherche pas de références dedans). */
function stripStrings(formula: string): string {
  return formula.replace(/"(?:[^"]|"")*"/g, (m) => ' '.repeat(m.length));
}

export function extractRefs(formula: string): FormulaRef[] {
  const src = stripStrings(formula);
  const out: FormulaRef[] = [];
  for (const m of src.matchAll(REF_RE)) {
    const sheet = m[1] !== undefined ? m[1].replace(/''/g, "'") : m[2];
    const [a, b = a] = m[3].split(':');
    let p: { row: number; col: number };
    let q: { row: number; col: number };
    try {
      p = fromA1(a);
      q = fromA1(b);
    } catch {
      continue;
    }
    out.push({
      sheet,
      startRow: Math.min(p.row, q.row),
      endRow: Math.max(p.row, q.row),
      startColumn: Math.min(p.col, q.col),
      endColumn: Math.max(p.col, q.col),
      text: m[0],
    });
  }
  return out;
}

export function refContains(ref: FormulaRef, row: number, col: number): boolean {
  return row >= ref.startRow && row <= ref.endRow && col >= ref.startColumn && col <= ref.endColumn;
}
