/**
 * Filtre avancé (zone de critères à la manière d'Excel).
 *
 * La zone de critères a une ligne d'en-têtes (noms de colonnes de la liste) puis des lignes
 * de conditions : les conditions d'une même ligne sont combinées par ET, les lignes par OU.
 * Une condition peut être « >10 », « <=2026 », « <>Paris », « =Lyon » (égalité exacte),
 * « Ly* » (jokers * et ?), ou un texte simple (« commence par », comme Excel).
 */
import { cellValue, compareKey, type Cell, type Grid } from './cells';

type Op = '=' | '<>' | '>' | '>=' | '<' | '<=' | 'begins';

interface Condition {
  col: number;
  op: Op;
  value: string;
}

function parseCondition(raw: string | number | boolean): { op: Op; value: string } | null {
  if (typeof raw === 'number') return { op: '=', value: String(raw) };
  if (typeof raw === 'boolean') return { op: '=', value: raw ? 'VRAI' : 'FAUX' };
  const s = raw.trim();
  if (s === '') return null;
  const m = /^(<>|>=|<=|=|>|<)(.*)$/.exec(s);
  if (m) return { op: m[1] as Op, value: m[2].trim() };
  return { op: 'begins', value: s };
}

function toNumber(s: string): number | null {
  const t = s.replace(/\s/g, '').replace(',', '.');
  if (t === '' || !/^[-+]?(\d+\.?\d*|\.\d+)(e[-+]?\d+)?$/i.test(t)) return null;
  return Number(t);
}

/** Motif Excel (* ? ~) → expression régulière. */
function wildcardRegex(pattern: string, anchoredEnd: boolean): RegExp {
  let re = '';
  for (let i = 0; i < pattern.length; i++) {
    const ch = pattern[i];
    if (ch === '~' && i + 1 < pattern.length) re += pattern[++i].replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    else if (ch === '*') re += '.*';
    else if (ch === '?') re += '.';
    else re += ch.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }
  return new RegExp(`^${re}${anchoredEnd ? '$' : ''}`, 'iu');
}

function matches(cell: Cell, cond: Condition): boolean {
  const v = cellValue(cell);
  const text = v === null ? '' : typeof v === 'boolean' ? (v ? 'VRAI' : 'FAUX') : String(v);
  const num = typeof v === 'number' ? v : null;
  const target = toNumber(cond.value);
  switch (cond.op) {
    case '=':
      if (cond.value === '') return v === null;
      if (num !== null && target !== null) return num === target;
      return wildcardRegex(cond.value, true).test(text);
    case '<>':
      if (cond.value === '') return v !== null;
      if (num !== null && target !== null) return num !== target;
      return !wildcardRegex(cond.value, true).test(text);
    case 'begins':
      if (num !== null && target !== null) return num === target;
      return wildcardRegex(cond.value, false).test(text);
    default: {
      if (num !== null && target !== null) {
        if (cond.op === '>') return num > target;
        if (cond.op === '>=') return num >= target;
        if (cond.op === '<') return num < target;
        return num <= target;
      }
      if (v === null || num !== null || target !== null) return false;
      const cmp = text.localeCompare(cond.value, 'fr', { sensitivity: 'base' });
      if (cond.op === '>') return cmp > 0;
      if (cond.op === '>=') return cmp >= 0;
      if (cond.op === '<') return cmp < 0;
      return cmp <= 0;
    }
  }
}

export interface AdvancedFilterResult {
  /** Indices (dans les données, sans l'en-tête) des lignes retenues. */
  rows: number[];
}

/**
 * @param list     plage de la liste, en-tête compris
 * @param criteria zone de critères, en-tête compris
 * @param unique   extraction sans doublon
 */
export function advancedFilter(list: Grid, criteria: Grid, unique = false): AdvancedFilterResult {
  const [header = [], ...data] = list;
  const [critHeader = [], ...critRows] = criteria;
  const colIndex = new Map(header.map((c, i) => [compareKey(c), i]));
  const orGroups: Condition[][] = critRows
    .map((row) =>
      row.flatMap((cell, i) => {
        const raw = cellValue(cell);
        if (raw === null) return [];
        const col = colIndex.get(compareKey(critHeader[i]));
        const parsed = parseCondition(raw);
        return col === undefined || !parsed ? [] : [{ col, ...parsed }];
      }),
    )
    .filter((g, i) => g.length || critRows[i].some((c) => cellValue(c) !== null));
  const result: number[] = [];
  const seen = new Set<string>();
  data.forEach((row, i) => {
    const ok = !orGroups.length || orGroups.some((g) => g.every((c) => matches(row[c.col], c)));
    if (!ok) return;
    if (unique) {
      const key = header.map((_, c) => compareKey(row[c])).join('\u0001');
      if (seen.has(key)) return;
      seen.add(key);
    }
    result.push(i);
  });
  return { rows: result };
}
