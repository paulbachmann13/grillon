/**
 * Compléments de l'onglet Démarrer pour coller au ruban Accueil d'Excel :
 * Couper/Copier, Coller les formules, Insérer, Supprimer, Format, Effacer, Retrait.
 * Toutes les actions réutilisent des commandes Univer existantes (annulables).
 */
import {
  ClearFormatDoubleIcon,
  ConditionalColorIcon,
  FilterIcon,
  SearchIcon,
  CopyIcon,
  CutIcon,
  DeleteRowDoubleIcon,
  InsertRowAboveDoubleIcon,
  RowHeightMediumIcon,
} from '@univerjs/icons';
import { IconManager } from '@univerjs/preset-sheets-core';
import { CommandType, ICommandService, type FUniver } from '@univerjs/presets';
import type { FWorksheet } from '@univerjs/sheets/facade';
import { createElement, type ComponentType } from 'react';
import { t } from '../i18n';

const START = 'ribbon.start';
const HISTORY = 'ribbon.start.history';
const LAYOUT = 'ribbon.start.layout';
const OTHERS = 'ribbon.start.others';

/** Icônes de retrait (absentes du jeu d'icônes Univer). */
function indentIcon(more: boolean): ComponentType<Record<string, unknown>> {
  return (props) =>
    createElement(
      'svg',
      { viewBox: '0 0 16 16', width: '1em', height: '1em', fill: 'none', stroke: 'currentColor', strokeWidth: 1.3, ...props },
      createElement('path', { d: 'M2 3h12M7 6.5h7M7 10h7M2 13.5h12' }),
      createElement('path', { d: more ? 'M2 6l2.5 2.25L2 10.5' : 'M4.5 6L2 8.25l2.5 2.25' }),
    );
}

/** Retrait en pixels pour un niveau Excel (≈ 3 caractères). */
const INDENT_STEP_PX = 9;

async function changeIndent(api: FUniver, delta: number): Promise<void> {
  const wb = api.getActiveWorkbook();
  const range = wb?.getActiveRange();
  if (!wb || !range) return;
  const data = wb.save();
  const cell = range.getCellData();
  const styleId = cell?.s;
  const style = typeof styleId === 'string' ? data.styles[styleId] : styleId;
  const current = style?.pd?.l ?? 0;
  const next = Math.max(0, Math.round(current / INDENT_STEP_PX + delta) * INDENT_STEP_PX);
  await api.executeCommand('sheet.command.set-style', {
    style: { type: 'pd', value: { ...(style?.pd ?? {}), l: next } },
  });
}

export function installHomeRibbon(api: FUniver): void {
  // Les icônes de menu sont résolues par l'IconManager d'Univer (et non par registerComponent).
  const injector = (api as unknown as { _injector: { get<T>(id: unknown): T } })._injector;
  injector.get<IconManager>(IconManager).register({
    GrillonCutIcon: CutIcon,
    GrillonCopyIcon: CopyIcon,
    GrillonInsertIcon: InsertRowAboveDoubleIcon,
    GrillonDeleteIcon: DeleteRowDoubleIcon,
    GrillonFormatIcon: RowHeightMediumIcon,
    GrillonClearIcon: ClearFormatDoubleIcon,
    GrillonIndentMoreIcon: indentIcon(true),
    GrillonIndentLessIcon: indentIcon(false),
    GrillonCfIcon: ConditionalColorIcon,
    GrillonFilterIcon: FilterIcon,
    GrillonSearchIcon: SearchIcon,
  } as never);

  const run = (id: string, params?: object) => () => void api.executeCommand(id, params);
  const commands = injector.get<ICommandService>(ICommandService);

  // Presse-papiers
  api.createMenu({ id: 'grillon.cut', icon: 'GrillonCutIcon', title: '', tooltip: t('ribbon.cut'), action: run('univer.command.cut'), order: 5, gridLayout: { row: 1, column: 4 } }).appendTo([START, HISTORY]);
  api.createMenu({ id: 'grillon.copy', icon: 'GrillonCopyIcon', title: '', tooltip: t('ribbon.copy'), action: run('univer.command.copy'), order: 6, gridLayout: { row: 2, column: 4 } }).appendTo([START, HISTORY]);
  commands.registerCommand({
    id: 'grillon.paste-formulas',
    type: CommandType.COMMAND,
    handler: () => commands.executeCommand('univer.command.paste', { value: 'special-paste-formula' }),
  });
  api.createMenu({ id: 'grillon.paste-formulas', title: t('ribbon.pasteFormulas'), action: 'grillon.paste-formulas', order: 2 }).appendTo([START, HISTORY, 'sheet.menu.paste']);

  // Retrait
  api.createMenu({ id: 'grillon.indent-less', icon: 'GrillonIndentLessIcon', title: '', tooltip: t('ribbon.indentLess'), action: () => void changeIndent(api, -1), order: 10, gridLayout: { row: 1, column: 4 } }).appendTo([START, LAYOUT]);
  api.createMenu({ id: 'grillon.indent-more', icon: 'GrillonIndentMoreIcon', title: '', tooltip: t('ribbon.indentMore'), action: () => void changeIndent(api, 1), order: 11, gridLayout: { row: 2, column: 4 } }).appendTo([START, LAYOUT]);

  // Cellules : Insérer / Supprimer / Format / Effacer (groupe « Cellules » d'Excel).
  // Le ruban en grille n'affiche pas les groupes imbriqués que crée FSubmenu.addSubmenu :
  // on ajoute donc le bouton déroulant, puis chaque élément directement sous sa clé.
  type Entry = [id: string, title: string, action: () => void];
  const dropdown = (id: string, icon: string, title: string, entries: Entry[]) => {
    api.createSubmenu({ id, icon, title, tooltip: title }).appendTo([START, OTHERS]);
    entries.forEach(([itemId, itemTitle, action], order) => {
      // Les éléments d'un menu déroulant du ruban exécutent la commande qui porte leur identifiant.
      if (!commands.hasCommand(itemId)) commands.registerCommand({ id: itemId, type: CommandType.COMMAND, handler: () => (action(), true) });
      api.createMenu({ id: itemId, title: itemTitle, action: itemId, order }).appendTo([START, OTHERS, id]);
    });
  };

  /** Applique une action à la feuille et à la sélection courantes (via la façade : annulable). */
  const onSelection = (fn: (ws: FWorksheet, row: number, col: number, rows: number, cols: number) => void) => () => {
    const ws = api.getActiveWorkbook()?.getActiveSheet();
    const range = api.getActiveWorkbook()?.getActiveRange();
    if (!ws || !range) return;
    fn(ws, range.getRow(), range.getColumn(), range.getLastRow() - range.getRow() + 1, range.getLastColumn() - range.getColumn() + 1);
  };

  dropdown('grillon.insert', 'GrillonInsertIcon', t('ribbon.insert'), [
    ['grillon.insert.cells-down', t('ribbon.insertCellsDown'), run('sheet.command.insert-range-move-down')],
    ['grillon.insert.cells-right', t('ribbon.insertCellsRight'), run('sheet.command.insert-range-move-right')],
    ['grillon.insert.rows', t('ribbon.insertRows'), onSelection((ws, row, _c, rows) => ws.insertRowsBefore(row, rows))],
    ['grillon.insert.cols', t('ribbon.insertCols'), onSelection((ws, _r, col, _rs, cols) => ws.insertColumnsBefore(col, cols))],
    ['grillon.insert.sheet', t('ribbon.insertSheet'), run('sheet.command.insert-sheet')],
  ]);
  dropdown('grillon.delete', 'GrillonDeleteIcon', t('ribbon.delete'), [
    ['grillon.delete.cells-up', t('ribbon.deleteCellsUp'), run('sheet.command.delete-range-move-up')],
    ['grillon.delete.cells-left', t('ribbon.deleteCellsLeft'), run('sheet.command.delete-range-move-left')],
    ['grillon.delete.rows', t('ribbon.deleteRows'), onSelection((ws, row, _c, rows) => ws.deleteRows(row, rows))],
    ['grillon.delete.cols', t('ribbon.deleteCols'), onSelection((ws, _r, col, _rs, cols) => ws.deleteColumns(col, cols))],
    ['grillon.delete.sheet', t('ribbon.deleteSheet'), run('sheet.command.remove-sheet-confirm')],
  ]);
  dropdown('grillon.format', 'GrillonFormatIcon', t('ribbon.format'), [
    ['grillon.format.row-height', t('ribbon.autoRowHeight'), onSelection((ws, row, _c, rows) => ws.setRowAutoHeight(row, rows))],
    ['grillon.format.col-width', t('ribbon.autoColWidth'), run('sheet.command.set-col-auto-width')],
    ['grillon.format.hide-rows', t('ribbon.hideRows'), onSelection((ws, row, _c, rows) => ws.hideRows(row, rows))],
    ['grillon.format.hide-cols', t('ribbon.hideCols'), onSelection((ws, _r, col, _rs, cols) => ws.hideColumns(col, cols))],
    ['grillon.format.rename', t('ribbon.renameSheet'), run('sheet.operation.rename-sheet')],
  ]);
  dropdown('grillon.clear', 'GrillonClearIcon', t('ribbon.clear'), [
    ['grillon.clear.all', t('ribbon.clearAll'), run('sheet.command.clear-selection-all')],
    ['grillon.clear.formats', t('ribbon.clearFormats'), run('sheet.command.clear-selection-format')],
    ['grillon.clear.contents', t('ribbon.clearContents'), run('sheet.command.clear-selection-content')],
  ]);

  // Styles et Édition (comme le ruban Accueil d'Excel) : raccourcis vers les outils de l'onglet Données.
  api
    .createMenu({ id: 'grillon.cf', icon: 'GrillonCfIcon', title: t('ribbon.conditionalFormatting'), tooltip: t('ribbon.conditionalFormatting'), action: 'sheet.operation.open.conditional.formatting.panel' })
    .appendTo([START, OTHERS]);
  dropdown('grillon.sort-filter', 'GrillonFilterIcon', t('ribbon.sortFilter'), [
    ['grillon.sort.asc', t('ribbon.sortAsc'), run('sheet.command.sort-range-asc')],
    ['grillon.sort.desc', t('ribbon.sortDesc'), run('sheet.command.sort-range-desc')],
    ['grillon.sort.custom', t('ribbon.sortCustom'), run('sheet.command.sort-range-custom')],
    ['grillon.filter', t('ribbon.filter'), run('sheet.command.smart-toggle-filter')],
  ]);
  api.createMenu({ id: 'grillon.find', icon: 'GrillonSearchIcon', title: t('ribbon.find'), tooltip: t('ribbon.find'), action: 'ui.operation.open-find-dialog' }).appendTo([START, OTHERS]);
}
