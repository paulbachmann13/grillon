/**
 * Tri comme dans Excel (Accueil → Trier et filtrer, menu contextuel) :
 * - une seule cellule, ou une partie des colonnes d'un tableau : tout le tableau est trié (les lignes restent entières) ;
 * - la ligne d'en-tête est reconnue et reste en haut ;
 * - les textes sont rangés à la française (É avec E, casse, tirets et apostrophes ignorés).
 * Univer, lui, triait la seule sélection, sans en-tête, et rangeait les lettres accentuées après Z.
 */
import type { FUniver } from '@univerjs/presets';
import type { FWorkbook, FWorksheet } from '@univerjs/sheets/facade';
import { SheetsSortService } from '@univerjs/sheets-sort';
import { compareFrench, hasHeaderRow, sortRange, type SortCell } from '../data/sort-detect';
import { currentRegionRect } from './region';

interface RawCell {
  v?: unknown;
  t?: number;
  s?: string | { bl?: number; bg?: { rgb?: string } } | null;
  p?: { body?: { dataStream?: string } } | null;
}

interface Location {
  unitId: string;
  subUnitId: string;
  range: { startRow: number; endRow: number; startColumn: number; endColumn: number };
  colIndex: number;
}

/** Commandes de tri direct d'Univer (ruban, menu contextuel) : remplacées par le tri ci-dessous. */
const DIRECT_SORT = /^sheet\.command\.sort-range-(asc|desc)(-ext)?(-ctx)?$/;

type SortService = {
  registerCompareFn(fn: (type: string, a: RawCell | null | undefined, b: RawCell | null | undefined) => number | null): void;
  applySort(option: { orderRules: { type: string; colIndex: number }[]; range: Location['range']; hasTitle?: boolean }, unitId?: string, subUnitId?: string): void;
};

/** Texte d'une cellule pour le tri (null si ce n'est pas du texte : nombres, vides, logiques sont laissés à Univer). */
function textOf(cell: RawCell | null | undefined): string | null {
  if (!cell) return null;
  const rich = cell.p?.body?.dataStream?.replace(/\r?\n$/, '').replace(/\r$/, '');
  if (rich) return rich;
  if ((cell.t === 1 || cell.t === undefined) && typeof cell.v === 'string' && cell.v !== '') return cell.v;
  return null;
}

const installed = new WeakSet<object>();

export function installSortUi(api: FUniver): void {
  if (installed.has(api)) return;
  installed.add(api);
  const injector = (api as unknown as { _injector: { get<T>(id: unknown): T } })._injector;
  const sorter = injector.get<SortService>(SheetsSortService);

  // Comparaison à la française, essayée avant celle d'Univer (qui prend la main pour le reste : nombres, vides…).
  sorter.registerCompareFn((type, a, b) => {
    const x = textOf(a);
    const y = textOf(b);
    if (x === null || y === null) return null;
    const res = compareFrench(x, y);
    return type === 'asc' ? res : -res;
  });

  const excelSort = (asc: boolean): boolean => {
    const wb = api.getActiveWorkbook() as unknown as FWorkbook | null;
    const ws = wb?.getActiveSheet() as unknown as FWorksheet | undefined;
    const sel = wb?.getActiveRange();
    if (!wb || !ws || !sel) return false;
    const active = wb.getActiveCell?.() ?? sel;
    const selection = { r1: sel.getRow(), c1: sel.getColumn(), r2: sel.getLastRow(), c2: sel.getLastColumn() };
    const region = currentRegionRect(ws, { r1: active.getRow(), c1: active.getColumn(), r2: active.getRow(), c2: active.getColumn() });
    const rect = sortRange(selection, region);
    if (rect.r2 <= rect.r1) return false;
    const styles = (wb as unknown as { getWorkbook(): { getStyles(): { get(id: string): { bl?: number; bg?: { rgb?: string } } | undefined } } }).getWorkbook().getStyles();
    const core = ws.getSheet() as unknown as { getCellRaw(r: number, c: number): RawCell | null | undefined };
    const sample: SortCell[][] = [];
    for (let r = rect.r1; r <= Math.min(rect.r2, rect.r1 + 20); r++) {
      const row: SortCell[] = [];
      for (let c = rect.c1; c <= rect.c2; c++) {
        const raw = core.getCellRaw(r, c);
        const style = typeof raw?.s === 'string' ? styles.get(raw.s) : raw?.s;
        const text = textOf(raw);
        row.push({ v: text ?? raw?.v ?? null, t: text !== null ? 1 : raw?.t, bold: !!style?.bl, fill: style?.bg?.rgb });
      }
      sample.push(row);
    }
    const loc: Location = {
      unitId: wb.getId(),
      subUnitId: ws.getSheetId(),
      range: { startRow: rect.r1, endRow: rect.r2, startColumn: rect.c1, endColumn: rect.c2 },
      colIndex: Math.min(Math.max(active.getColumn(), rect.c1), rect.c2),
    };
    sorter.applySort({ orderRules: [{ type: asc ? 'asc' : 'desc', colIndex: loc.colIndex }], range: loc.range, hasTitle: hasHeaderRow(sample) }, loc.unitId, loc.subUnitId);
    return true;
  };

  api.addEvent(api.Event.BeforeCommandExecute, (e) => {
    const m = DIRECT_SORT.exec(e.id);
    if (!m) return;
    e.cancel = true;
    excelSort(m[1] === 'asc');
  });
}

export { DIRECT_SORT };
