/**
 * Segments et chronologies (Insertion → Filtres), comme Excel : des boutons posés sur la feuille
 * qui filtrent un tableau ou un tableau croisé dynamique. Logique pure, sans DOM ni Univer.
 */

export type TimelineLevel = 'years' | 'quarters' | 'months' | 'days';

/** Ce qu'un segment filtre : une colonne d'un tableau (plage avec en-têtes) ou le champ source d'un tableau croisé. */
export type SlicerTarget =
  | { kind: 'table'; sheet: string; /** Plage avec la ligne d'en-têtes, « A1:D50 ». */ ref: string; /** Colonne dans la plage (à partir de 0). */ col: number }
  | { kind: 'pivot'; sheet: string; /** Colonne dans la plage source du tableau croisé (à partir de 0). */ col: number };

export interface SlicerSpec {
  id: string;
  type: 'slicer' | 'timeline';
  /** Titre affiché (par défaut le nom de la colonne). */
  title: string;
  target: SlicerTarget;
  /** Segment : libellés retenus (null : tout). */
  selected: string[] | null;
  /** Chronologie : niveau et bornes (clés de période comprises) ; null : toutes les dates. */
  level?: TimelineLevel;
  from?: string | null;
  to?: string | null;
}

/** Cellule d'une colonne : texte affiché et valeur brute. */
export interface ColumnCell {
  label: string;
  value: string | number | boolean | null;
}

export type SlicerFilter =
  | { kind: 'slicer'; col: number; selected: string[] | null }
  | { kind: 'timeline'; col: number; level: TimelineLevel; from: string | null; to: string | null };

export const EMPTY_LABEL = '(vide)';

export function filterOf(spec: SlicerSpec): SlicerFilter {
  return spec.type === 'timeline'
    ? { kind: 'timeline', col: spec.target.col, level: spec.level ?? 'months', from: spec.from ?? null, to: spec.to ?? null }
    : { kind: 'slicer', col: spec.target.col, selected: spec.selected };
}

/** Le filtre retient-il tout ? (rien à masquer) */
export const isOpen = (f: SlicerFilter) => (f.kind === 'slicer' ? f.selected === null : f.from === null && f.to === null);

// --- Dates ---

const DAY_MS = 86_400_000;
const EPOCH = Date.UTC(1899, 11, 30);

/** Numéro de série Excel → année, mois (1-12), jour. */
export function serialToYmd(serial: number): { y: number; m: number; d: number } {
  const date = new Date(EPOCH + Math.floor(serial) * DAY_MS);
  return { y: date.getUTCFullYear(), m: date.getUTCMonth() + 1, d: date.getUTCDate() };
}

export function ymdToSerial(y: number, m: number, d: number): number {
  return Math.round((Date.UTC(y, m - 1, d) - EPOCH) / DAY_MS);
}

const p2 = (n: number) => String(n).padStart(2, '0');

/** Clé de période triable : « 2024 », « 2024-T1 », « 2024-03 », « 2024-03-15 ». */
export function periodKey(serial: number, level: TimelineLevel): string {
  const { y, m, d } = serialToYmd(serial);
  if (level === 'years') return String(y);
  if (level === 'quarters') return `${y}-T${Math.ceil(m / 3)}`;
  if (level === 'months') return `${y}-${p2(m)}`;
  return `${y}-${p2(m)}-${p2(d)}`;
}

const MONTHS = ['janv.', 'févr.', 'mars', 'avr.', 'mai', 'juin', 'juil.', 'août', 'sept.', 'oct.', 'nov.', 'déc.'];

export function periodLabel(key: string): string {
  const q = /^(\d{4})-T(\d)$/.exec(key);
  if (q) return `T${q[2]} ${q[1]}`;
  const parts = key.split('-').map(Number);
  if (parts.length === 1) return String(parts[0]);
  if (parts.length === 2) return `${MONTHS[parts[1] - 1]} ${parts[0]}`;
  return `${parts[2]} ${MONTHS[parts[1] - 1]} ${parts[0]}`;
}

/** Période qui suit (même niveau). */
function nextKey(key: string, level: TimelineLevel): string {
  if (level === 'years') return String(Number(key) + 1);
  if (level === 'quarters') {
    const [y, q] = key.split('-T').map(Number);
    return q === 4 ? `${y + 1}-T1` : `${y}-T${q + 1}`;
  }
  if (level === 'months') {
    const [y, m] = key.split('-').map(Number);
    return m === 12 ? `${y + 1}-01` : `${y}-${p2(m + 1)}`;
  }
  const [y, m, d] = key.split('-').map(Number);
  return periodKey(ymdToSerial(y, m, d) + 1, 'days');
}

/** Toutes les périodes entre la plus ancienne et la plus récente date (une pour chaque période, même sans donnée). */
export function timelinePeriods(values: readonly (string | number | boolean | null)[], level: TimelineLevel, limit = 2000): string[] {
  const serials = values.filter((v): v is number => typeof v === 'number' && Number.isFinite(v) && v >= 1);
  if (!serials.length) return [];
  const first = periodKey(Math.min(...serials), level);
  const last = periodKey(Math.max(...serials), level);
  const out: string[] = [];
  for (let k = first; out.length < limit; k = nextKey(k, level)) {
    out.push(k);
    if (k >= last) break;
  }
  return out;
}

// --- Filtrage ---

export function cellMatches(cell: ColumnCell | undefined, f: SlicerFilter): boolean {
  if (isOpen(f)) return true;
  if (f.kind === 'slicer') return f.selected!.includes(cell?.label || EMPTY_LABEL);
  const v = cell?.value;
  if (typeof v !== 'number') return false;
  const key = periodKey(v, f.level);
  return (f.from === null || key >= f.from) && (f.to === null || key <= f.to);
}

/** Lignes de données retenues par tous les filtres (`skip` : filtre ignoré, pour compter les éléments disponibles). */
export function visibleRows(rows: readonly (readonly ColumnCell[])[], filters: readonly SlicerFilter[], skip = -1): boolean[] {
  return rows.map((row) => filters.every((f, i) => i === skip || cellMatches(row[f.col], f)));
}

/** Éléments d'un segment, triés, avec le nombre de lignes retenues par les AUTRES filtres (0 : grisé, comme Excel). */
export function slicerItems(rows: readonly (readonly ColumnCell[])[], filters: readonly SlicerFilter[], index: number): { label: string; count: number }[] {
  const col = filters[index].col;
  const mask = visibleRows(rows, filters, index);
  const items = new Map<string, { label: string; count: number; sort: string | number }>();
  rows.forEach((row, i) => {
    const cell = row[col];
    const label = cell?.label || EMPTY_LABEL;
    const entry = items.get(label) ?? { label, count: 0, sort: cell?.label ? (typeof cell.value === 'number' ? cell.value : cell.label.toLocaleLowerCase('fr')) : '￿' };
    if (mask[i]) entry.count++;
    items.set(label, entry);
  });
  return [...items.values()]
    .sort((a, b) => (typeof a.sort === 'number' && typeof b.sort === 'number' ? a.sort - b.sort : typeof a.sort === 'number' ? -1 : typeof b.sort === 'number' ? 1 : String(a.sort).localeCompare(String(b.sort), 'fr')))
    .map(({ label, count }) => ({ label, count }));
}

/** Clic sur un élément : seul retenu ; avec Ctrl/⌘/Maj, bascule ; recliquer sur le seul élément retenu efface le filtre. */
export function clickItem(selected: string[] | null, label: string, all: readonly string[], multi: boolean): string[] | null {
  const current = selected ?? [...all];
  if (multi) {
    const next = current.includes(label) ? current.filter((l) => l !== label) : [...current, label];
    // Plus rien de retenu ou tout retenu : pas de filtre, comme Excel.
    return next.length === 0 || next.length === all.length ? null : next;
  }
  if (selected && selected.length === 1 && selected[0] === label) return null;
  return all.length === 1 ? null : [label];
}

/** Clic sur une période de la chronologie : une seule ; avec Maj, étend jusqu'à la période cliquée. */
export function clickPeriod(range: { from: string | null; to: string | null }, key: string, extend: boolean): { from: string | null; to: string | null } {
  if (extend && range.from !== null) {
    const [a, b] = [range.from, key].sort();
    return { from: a, to: range.to !== null && range.to > b ? range.to : b };
  }
  if (range.from === key && range.to === key) return { from: null, to: null };
  return { from: key, to: key };
}

/** Libellés (texte brut du champ source) à garder dans un tableau croisé pour ce filtre. */
export function pivotInclude(column: readonly ColumnCell[], f: SlicerFilter): string[] | null {
  if (isOpen(f)) return null;
  return [...new Set(column.filter((c) => cellMatches(c, f)).map((c) => c.label || EMPTY_LABEL))];
}
