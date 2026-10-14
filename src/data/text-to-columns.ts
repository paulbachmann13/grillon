/**
 * « Convertir » (texte en colonnes) avec séparateurs, comme l'assistant d'Excel.
 */
import { parseCsvValue } from '../io/csv';
import type { ICellData } from '../model/univer-types';

export interface SplitOptions {
  tab?: boolean;
  semicolon?: boolean;
  comma?: boolean;
  space?: boolean;
  /** Autre séparateur (un caractère). */
  other?: string;
  /** Traiter les séparateurs consécutifs comme un seul. */
  consecutiveAsOne?: boolean;
  /** Qualificateur de texte (les séparateurs entre guillemets sont ignorés). */
  textQualifier?: '"' | "'" | null;
  /** Séparateur décimal des nombres à reconnaître. */
  decimal?: '.' | ',';
}

export function splitText(text: string, opts: SplitOptions): string[] {
  const seps = new Set<string>();
  if (opts.tab) seps.add('\t');
  if (opts.semicolon) seps.add(';');
  if (opts.comma) seps.add(',');
  if (opts.space) seps.add(' ');
  if (opts.other) seps.add(opts.other[0]);
  if (!seps.size) return [text];
  const q = opts.textQualifier === undefined ? '"' : opts.textQualifier;
  const out: string[] = [];
  let cur = '';
  let quoted = false;
  let lastWasSep = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (q && ch === q) {
      if (quoted && text[i + 1] === q) {
        cur += q;
        i++;
      } else quoted = !quoted;
      lastWasSep = false;
      continue;
    }
    if (!quoted && seps.has(ch)) {
      if (opts.consecutiveAsOne && lastWasSep) continue;
      out.push(cur);
      cur = '';
      lastWasSep = true;
      continue;
    }
    cur += ch;
    lastWasSep = false;
  }
  out.push(cur);
  return out;
}

/** Découpe une colonne de textes et type chaque morceau (nombres, dates…). */
export function textToColumns(values: string[], opts: SplitOptions): (ICellData | null)[][] {
  const decimal = opts.decimal ?? ',';
  return values.map((v) => (v === '' ? [] : splitText(v, opts).map((part) => parseCsvValue(part, decimal))));
}
