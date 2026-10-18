/**
 * Tableau croisé dynamique (version simple) : champs en lignes, en colonnes, valeurs
 * (somme, moyenne, nombre, min, max) et filtres ; totaux et total général.
 * Le résultat est une grille de cellules écrite dans une feuille (disposition tabulaire).
 */
import { CellValueType, type ICellData } from '../model/univer-types';
import { cellText, cellValue, compareKey, type Grid } from './cells';

export type PivotFunction = 'sum' | 'average' | 'count' | 'min' | 'max';

export interface PivotValueField {
  col: number;
  fn: PivotFunction;
}

export interface PivotFilter {
  col: number;
  /** Valeurs (texte affiché) retenues. */
  include: string[];
}

export interface PivotConfig {
  rows: number[];
  columns: number[];
  values: PivotValueField[];
  filters: PivotFilter[];
}

/** Tableau croisé enregistré dans une feuille (métadonnées custom.grillon.pivot). */
export interface PivotDefinition {
  sourceSheet: string;
  sourceRange: string;
  config: PivotConfig;
  /** Taille de la dernière sortie (pour l'effacer avant d'actualiser). */
  outRows: number;
  outCols: number;
  /** Coin haut gauche de la sortie (par défaut A3, la ligne 1 décrivant la source). */
  outRow?: number;
  outCol?: number;
  /** Nom du tableau croisé (dans le .xlsx). */
  name?: string;
  /** Colonnes source dont le filtre est posé par un segment ou une chronologie. */
  slicerCols?: number[];
}

export interface PivotLabels {
  rowLabels: string;
  grandTotal: string;
  total: (label: string) => string;
  empty: string;
  fn: Record<PivotFunction, string>;
}

export const FR_PIVOT_LABELS: PivotLabels = {
  rowLabels: 'Étiquettes de lignes',
  grandTotal: 'Total général',
  total: (l) => `Total ${l}`,
  empty: '(vide)',
  fn: { sum: 'Somme de', average: 'Moyenne de', count: 'Nombre de', min: 'Min de', max: 'Max de' },
};

interface Acc {
  sum: number;
  count: number;
  numbers: number;
  min: number;
  max: number;
}

const newAcc = (): Acc => ({ sum: 0, count: 0, numbers: 0, min: Infinity, max: -Infinity });

function add(acc: Acc, v: ReturnType<typeof cellValue>): void {
  if (v === null) return;
  acc.count++;
  const n = typeof v === 'number' ? v : typeof v === 'boolean' ? null : Number.isFinite(Number(String(v).replace(',', '.'))) && String(v).trim() !== '' ? Number(String(v).replace(',', '.')) : null;
  if (n === null) return;
  acc.numbers++;
  acc.sum += n;
  acc.min = Math.min(acc.min, n);
  acc.max = Math.max(acc.max, n);
}

function result(acc: Acc | undefined, fn: PivotFunction): number | null {
  if (!acc) return null;
  switch (fn) {
    case 'count':
      return acc.count || null;
    case 'sum':
      return acc.numbers ? acc.sum : acc.count ? 0 : null;
    case 'average':
      return acc.numbers ? acc.sum / acc.numbers : null;
    case 'min':
      return acc.numbers ? acc.min : null;
    case 'max':
      return acc.numbers ? acc.max : null;
  }
}

/** Libellé et clé de tri d'une valeur de champ. */
function keyOf(cell: ICellData | null | undefined, labels: PivotLabels): { key: string; label: string; sort: string | number } {
  const v = cellValue(cell);
  if (v === null) return { key: '∅', label: labels.empty, sort: '￿' };
  return { key: compareKey(cell), label: cellText(cell), sort: typeof v === 'number' ? v : String(v).toLocaleLowerCase('fr') };
}

function compareSort(a: string | number, b: string | number): number {
  if (typeof a === 'number' && typeof b === 'number') return a - b;
  if (typeof a === 'number') return -1;
  if (typeof b === 'number') return 1;
  return a.localeCompare(b, 'fr');
}

const num = (v: number | null): ICellData | null => (v === null ? null : { v, t: CellValueType.NUMBER });
const txt = (v: string, bold = false): ICellData => ({ v, t: CellValueType.STRING, ...(bold ? { s: { bl: 1 } } : {}) });

/** Valeurs distinctes d'une colonne (pour la liste des filtres). */
export function distinctValues(grid: Grid, col: number, labels: PivotLabels = FR_PIVOT_LABELS): string[] {
  const seen = new Map<string, { label: string; sort: string | number }>();
  for (const row of grid.slice(1)) {
    const k = keyOf(row[col], labels);
    if (!seen.has(k.key)) seen.set(k.key, k);
  }
  return [...seen.values()].sort((a, b) => compareSort(a.sort, b.sort)).map((k) => k.label);
}

export function computePivot(grid: Grid, config: PivotConfig, labels: PivotLabels = FR_PIVOT_LABELS): Grid {
  const [header = [], ...data] = grid;
  const name = (c: number) => cellText(header[c]) || `Colonne ${c + 1}`;
  const values = config.values.length ? config.values : [];
  // Filtres.
  const rows = data.filter((row) => config.filters.every((f) => f.include.includes(keyOf(row[f.col], labels).label)));

  // Clés de lignes et de colonnes (combinaisons des champs), triées.
  type Key = { id: string; parts: { label: string; sort: string | number }[] };
  const collect = (fields: number[]) => {
    const map = new Map<string, Key>();
    for (const row of rows) {
      const parts = fields.map((c) => keyOf(row[c], labels));
      const id = parts.map((p) => p.key).join('\u0001');
      if (!map.has(id)) map.set(id, { id, parts });
    }
    return [...map.values()].sort((a, b) => {
      for (let i = 0; i < a.parts.length; i++) {
        const d = compareSort(a.parts[i].sort, b.parts[i].sort);
        if (d) return d;
      }
      return 0;
    });
  };
  const rowKeys = config.rows.length ? collect(config.rows) : [{ id: '', parts: [] }];
  const colKeys = config.columns.length ? collect(config.columns) : [{ id: '', parts: [] }];

  // Agrégats : cellule (ligne, colonne), totaux de ligne, totaux de colonne, total général.
  const cells = new Map<string, Acc[]>();
  const get = (k: string) => {
    let a = cells.get(k);
    if (!a) cells.set(k, (a = values.map(newAcc)));
    return a;
  };
  for (const row of rows) {
    const rk = config.rows.map((c) => keyOf(row[c], labels).key).join('\u0001');
    const ck = config.columns.map((c) => keyOf(row[c], labels).key).join('\u0001');
    for (const k of [`${rk}|${ck}`, `${rk}|*`, `*|${ck}`, '*|*']) {
      const accs = get(k);
      values.forEach((v, i) => add(accs[i], cellValue(row[v.col])));
    }
  }

  const out: Grid = [];
  const multiValue = values.length > 1;
  const valueName = (v: PivotValueField) => `${labels.fn[v.fn]} ${name(v.col)}`;
  const nRowFields = Math.max(1, config.rows.length);
  // En-têtes : ligne des champs de colonnes (si présents), puis ligne des libellés.
  if (config.columns.length) {
    const h1: (ICellData | null)[] = [txt(multiValue ? '' : valueName(values[0] ?? { col: 0, fn: 'count' }), true)];
    for (let i = 1; i < nRowFields; i++) h1.push(null);
    for (const ck of colKeys) for (let v = 0; v < Math.max(1, values.length); v++) h1.push(v === 0 ? txt(ck.parts.map((p) => p.label).join(' - '), true) : null);
    out.push(h1);
  }
  const h2: (ICellData | null)[] = [];
  if (config.rows.length) config.rows.forEach((c, i) => h2.push(txt(i === 0 && config.rows.length === 1 ? labels.rowLabels : name(c), true)));
  else h2.push(txt('', true));
  for (let k = 0; k < colKeys.length; k++) {
    for (const v of values) h2.push(txt(config.columns.length ? (multiValue ? valueName(v) : '') : valueName(v), true));
  }
  if (config.columns.length) for (const v of values) h2.push(txt(multiValue ? `${labels.grandTotal} - ${valueName(v)}` : labels.grandTotal, true));
  out.push(h2);

  const line = (rk: string, label: (ICellData | null)[], bold = false) => {
    const r: (ICellData | null)[] = [...label];
    for (const ck of colKeys) {
      const accs = cells.get(`${rk}|${config.columns.length ? ck.id : '*'}`);
      values.forEach((v, i) => {
        const c = num(result(accs?.[i], v.fn));
        r.push(c && bold ? { ...c, s: { bl: 1 } } : c);
      });
    }
    if (config.columns.length) {
      const accs = cells.get(`${rk}|*`);
      values.forEach((v, i) => {
        const c = num(result(accs?.[i], v.fn));
        r.push(c ? { ...c, s: { bl: 1 } } : c);
      });
    }
    out.push(r);
  };

  if (config.rows.length) {
    for (const rk of rowKeys) line(rk.id, rk.parts.map((p) => txt(p.label)));
  }
  const totalLabel: (ICellData | null)[] = [txt(labels.grandTotal, true)];
  for (let i = 1; i < nRowFields; i++) totalLabel.push(null);
  line('*', totalLabel, true);
  return out;
}
