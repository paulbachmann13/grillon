/**
 * « Sous-total » (Données) : insère une ligne de total après chaque groupe de valeurs
 * identiques dans une colonne, avec la fonction SUBTOTAL, plus un total général,
 * et renvoie les niveaux de plan (comme Excel : détail = 2, sous-totaux = 1).
 */
import { CellValueType, type ICellData } from '../model/univer-types';
import { columnToLetters } from '../io/xlsx/units';
import { cellText, compareKey, type Grid } from './cells';

/** Numéros de fonction de SUBTOTAL. */
export const SUBTOTAL_FUNCTIONS = {
  sum: 9,
  count: 2,
  counta: 3,
  average: 1,
  max: 4,
  min: 5,
  product: 6,
  stdev: 7,
  var: 10,
} as const;

export type SubtotalFunction = keyof typeof SUBTOTAL_FUNCTIONS;

export interface SubtotalOptions {
  /** Colonne (relative à la plage) dont les changements délimitent les groupes. */
  groupBy: number;
  fn: SubtotalFunction;
  /** Colonnes (relatives) qui reçoivent un sous-total. */
  columns: number[];
  /** Ligne (absolue, base 0) de l'en-tête de la plage. */
  startRow: number;
  /** Colonne (absolue, base 0) de la première colonne de la plage. */
  startColumn: number;
  labels?: { total: (group: string) => string; grandTotal: string };
}

export interface SubtotalResult {
  rows: Grid;
  /** Niveau de plan de chaque ligne produite (même ordre que `rows`). */
  outline: number[];
}

const DEFAULT_LABELS = { total: (g: string) => `Total ${g}`, grandTotal: 'Total général' };

/** `grid` contient l'en-tête puis les données (déjà triées par la colonne de regroupement). */
export function computeSubtotals(grid: Grid, opts: SubtotalOptions): SubtotalResult {
  const labels = opts.labels ?? DEFAULT_LABELS;
  const fnNum = SUBTOTAL_FUNCTIONS[opts.fn];
  const width = Math.max(...grid.map((r) => r.length), opts.groupBy + 1, ...opts.columns.map((c) => c + 1));
  const [header, ...data] = grid;
  const rows: Grid = [header];
  const outline: number[] = [0];
  const absRow = (i: number) => opts.startRow + i + 1; // numéro de ligne Excel (base 1)
  const colLetter = (c: number) => columnToLetters(opts.startColumn + c);

  const emitSubtotal = (label: string, fromIdx: number, toIdx: number, level: number) => {
    const row: (ICellData | null)[] = Array.from({ length: width }, () => null);
    row[opts.groupBy] = { v: label, t: CellValueType.STRING, s: { bl: 1 } };
    for (const c of opts.columns) {
      if (c === opts.groupBy) continue;
      row[c] = { f: `=SUBTOTAL(${fnNum},${colLetter(c)}${absRow(fromIdx)}:${colLetter(c)}${absRow(toIdx)})`, s: { bl: 1 } };
    }
    rows.push(row);
    outline.push(level);
  };

  let i = 0;
  while (i < data.length) {
    const key = compareKey(data[i][opts.groupBy]);
    const label = cellText(data[i][opts.groupBy]);
    const from = rows.length;
    while (i < data.length && compareKey(data[i][opts.groupBy]) === key) {
      rows.push(data[i]);
      outline.push(2);
      i++;
    }
    const to = rows.length - 1;
    emitSubtotal(labels.total(label), from, to, 1);
  }
  // Total général : SUBTOTAL ignore les autres SUBTOTAL, on peut donc prendre toute la colonne.
  if (data.length) emitSubtotal(labels.grandTotal, 1, rows.length - 1, 0);
  return { rows, outline };
}
