/**
 * Tri (Données → Trier, filtres) : Univer réordonne les lignes sans adapter les formules, ce qui fausse les résultats.
 * On ajoute à la commande de réordonnancement les mutations qui écrivent, comme Excel, chaque formule décalée
 * du déplacement de sa ligne (une seule étape d'annulation : ⌘Z annule le tri et les formules ensemble).
 */
import { FormulaDataModel } from '@univerjs/engine-formula';
import { SheetInterceptorService } from '@univerjs/preset-sheets-core';
import type { FUniver } from '@univerjs/presets';
import { sortFormulaFixes, type CellPatch, type SortRange } from '../data/sort-formulas';

const REORDER = 'sheet.command.reorder-range';
const SET_VALUES = 'sheet.mutation.set-range-values';

interface ReorderParams {
  unitId: string;
  subUnitId: string;
  range: SortRange;
  order: Record<number, number>;
}

interface RawCell {
  f?: string | null;
  si?: string | null;
}

interface CoreSheet {
  getCellRaw(row: number, col: number): RawCell | null | undefined;
  getCellMatrix(): { forValue(cb: (row: number, col: number, cell: RawCell | null | undefined) => void): void };
}

const empty = () => ({ redos: [], undos: [] });
const installed = new WeakSet<object>();

export function installSortFormulas(api: FUniver): void {
  if (installed.has(api)) return;
  installed.add(api);
  const injector = (api as unknown as { _injector: { get<T>(id: unknown): T } })._injector;
  const formulas = injector.get<FormulaDataModel>(FormulaDataModel);
  injector.get<SheetInterceptorService>(SheetInterceptorService).interceptCommand({
    getMutations: (command: { id: string; params?: unknown }) => {
      if (command.id !== REORDER || !command.params) return empty();
      const { unitId, subUnitId, range, order } = command.params as ReorderParams;
      const sheet = (api.getWorkbook(unitId)?.getSheetBySheetId(subUnitId)?.getSheet() ?? null) as unknown as CoreSheet | null;
      if (!sheet) return empty();
      const formulaAt = (row: number, col: number) => {
        const raw = sheet.getCellRaw(row, col);
        if (!raw?.f && !raw?.si) return null;
        const formula = formulas.getFormulaStringByCell(row, col, subUnitId, unitId) ?? raw.f ?? null;
        return formula ? { formula, f: raw.f ?? null, si: raw.si ?? null } : null;
      };
      const { redo, undo } = sortFormulaFixes(order, range, formulaAt);

      // Formule partagée dont la première cellule est dans la plage mais d'autres cellules en dehors : ces dernières
      // perdraient leur formule (elle dépend de la position de la première) ; on la rend explicite.
      const inside = (r: number, c: number) => r >= range.startRow && r <= range.endRow && c >= range.startColumn && c <= range.endColumn;
      const sharedInRange = new Set<string>();
      for (let r = range.startRow; r <= range.endRow; r++) {
        for (let c = range.startColumn; c <= range.endColumn; c++) {
          const raw = sheet.getCellRaw(r, c);
          if (raw?.si && raw.f) sharedInRange.add(String(raw.si));
        }
      }
      const outsideRedo: CellPatch = {};
      const outsideUndo: CellPatch = {};
      if (sharedInRange.size) {
        sheet.getCellMatrix().forValue((r, c, raw) => {
          if (!raw?.si || inside(r, c) || !sharedInRange.has(String(raw.si))) return;
          const formula = formulas.getFormulaStringByCell(r, c, subUnitId, unitId);
          if (!formula) return;
          (outsideRedo[r] ??= {})[c] = { f: formula, si: null };
          (outsideUndo[r] ??= {})[c] = { f: raw.f ?? null, si: raw.si ?? null };
        });
      }
      const mutation = (cellValue: CellPatch) => ({ id: SET_VALUES, params: { unitId, subUnitId, cellValue } });
      const has = (p: CellPatch) => Object.keys(p).length > 0;
      return {
        // Les cellules hors plage d'abord (pendant que la première cellule est encore à sa place), puis le tri, puis la plage.
        preRedos: has(outsideRedo) ? [mutation(outsideRedo)] : [],
        redos: has(redo) ? [mutation(redo)] : [],
        // Annulation : on rétablit d'abord les formules d'origine dans la plage triée, puis Univer remet les lignes en place.
        preUndos: has(undo) ? [mutation(undo)] : [],
        undos: has(outsideUndo) ? [mutation(outsideUndo)] : [],
      };
    },
  } as never);
}
