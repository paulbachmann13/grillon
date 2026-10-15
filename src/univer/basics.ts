/**
 * Bases d'Excel absentes d'Univer : Somme automatique (Σ, Alt+=), Remplir (bas/droite/haut/gauche),
 * Afficher les formules (Ctrl+`), sélection d'une colonne / ligne (Ctrl+Espace / Maj+Espace),
 * Ctrl+Origine / Ctrl+Fin, formats de nombre (Ctrl+Maj+1 à 6) et Insérer / Supprimer des cellules
 * (Ctrl+Maj++ / Ctrl+-). Toutes les écritures passent par la façade Univer : elles sont annulables.
 */
import { INTERCEPTOR_POINT, IEditorBridgeService, IconManager, SheetInterceptorService } from '@univerjs/preset-sheets-core';
import { CommandType, ICommandService, InterceptorEffectEnum, type FUniver } from '@univerjs/presets';
import { createElement, type ComponentType } from 'react';
import { autoSumPlan, type AutoSumFunction } from '../data/autosum';
import { t } from '../i18n';
import { box, openDialog, radio } from '../ui/dialog';

const START_OTHERS = ['ribbon.start', 'ribbon.start.others'];
const FORMULAS_OTHERS = ['ribbon.formulas', 'ribbon.formulas.others'];
const VIEW_DISPLAY = ['ribbon.view', 'ribbon.view.display'];
const STRING = 1;

function icon(path: string): ComponentType<Record<string, unknown>> {
  return (props) =>
    createElement('svg', { viewBox: '0 0 16 16', width: '1em', height: '1em', fill: 'none', stroke: 'currentColor', strokeWidth: 1.3, strokeLinecap: 'round', strokeLinejoin: 'round', ...props }, createElement('path', { d: path }));
}

let showFormulas = false;

type Api = FUniver;

const sheetAndRange = (api: Api) => {
  const wb = api.getActiveWorkbook();
  const ws = wb?.getActiveSheet();
  const range = wb?.getActiveRange();
  return wb && ws && range ? { wb, ws, range } : null;
};

function autoSum(api: Api, fn: AutoSumFunction): void {
  const cur = sheetAndRange(api);
  if (!cur) return;
  const { ws, range } = cur;
  const isNumber = (row: number, col: number) => {
    const cell = ws.getRange(row, col).getCellDataGrid()[0]?.[0];
    return typeof cell?.v === 'number' && cell.t !== STRING;
  };
  const plan = autoSumPlan(isNumber, { r1: range.getRow(), c1: range.getColumn(), r2: range.getLastRow(), c2: range.getLastColumn() }, fn);
  if (!plan.length) return;
  // Les totaux d'une même ligne s'écrivent en une seule opération (un seul « Annuler »).
  ws.getRange(plan[0]!.row, plan[0]!.col, 1, plan.length).setValues([plan.map((w) => w.formula)]);
}

type FillDirection = 'down' | 'right' | 'up' | 'left';

async function fill(api: Api, direction: FillDirection): Promise<void> {
  const cur = sheetAndRange(api);
  if (!cur) return;
  const { ws, range } = cur;
  const [r1, c1, r2, c2] = [range.getRow(), range.getColumn(), range.getLastRow(), range.getLastColumn()];
  if (direction === 'down') return void (await api.executeCommand('sheet.command.copy-down'));
  if (direction === 'right') return void (await api.executeCommand('sheet.command.copy-right'));
  if ((direction === 'up' && r1 === r2) || (direction === 'left' && c1 === c2)) return;
  const source = direction === 'up' ? ws.getRange(r2, c1, 1, c2 - c1 + 1) : ws.getRange(r1, c2, r2 - r1 + 1, 1);
  await source.autoFill(range, 'COPY' as never);
}

/** Boîte « Insérer / Supprimer » d'Excel (Ctrl+Maj++ et Ctrl+-). */
async function cellsDialog(api: Api, mode: 'insert' | 'delete'): Promise<void> {
  const cur = sheetAndRange(api);
  if (!cur) return;
  const insert = mode === 'insert';
  const options: [value: string, label: string, run: () => unknown][] = [
    insert
      ? ['shift-main', t('basics.shiftRight'), () => api.executeCommand('sheet.command.insert-range-move-right')]
      : ['shift-main', t('basics.shiftLeft'), () => api.executeCommand('sheet.command.delete-range-move-left')],
    insert
      ? ['shift-other', t('basics.shiftDown'), () => api.executeCommand('sheet.command.insert-range-move-down')]
      : ['shift-other', t('basics.shiftUp'), () => api.executeCommand('sheet.command.delete-range-move-up')],
    [
      'rows',
      t('basics.wholeRow'),
      () => {
        const { ws, range } = cur;
        const rows = range.getLastRow() - range.getRow() + 1;
        return insert ? ws.insertRowsBefore(range.getRow(), rows) : ws.deleteRows(range.getRow(), rows);
      },
    ],
    [
      'cols',
      t('basics.wholeColumn'),
      () => {
        const { ws, range } = cur;
        const cols = range.getLastColumn() - range.getColumn() + 1;
        return insert ? ws.insertColumnsBefore(range.getColumn(), cols) : ws.deleteColumns(range.getColumn(), cols);
      },
    ],
  ];
  const radios = options.map(([value, label], i) => radio('g-cells-action', label, value, i === 1));
  await openDialog({
    title: insert ? t('basics.insertTitle') : t('basics.deleteTitle'),
    testId: 'dialog-cells',
    body: box(...radios.map((r) => r.el)),
    onOk: () => {
      const index = radios.findIndex((r) => r.input.checked);
      void options[index]?.[2]();
    },
  });
}

/** Formats de nombre : Ctrl+Maj+1 à 6 (touches physiques, valables sur AZERTY comme sur QWERTY). */
const NUMBER_FORMATS: Record<string, string> = {
  Digit1: '#,##0.00',
  Digit2: 'hh:mm',
  Digit3: 'dd/mm/yyyy',
  Digit4: '#,##0.00 "€"',
  Digit5: '0%',
  Digit6: '0.00E+00',
};

function isTypingInCell(editor: IEditorBridgeService): boolean {
  const active = document.activeElement as HTMLElement | null;
  const typing = !!active && (active.tagName === 'INPUT' || active.tagName === 'TEXTAREA' || active.isContentEditable);
  if (typing && active?.id !== '__editor___INTERNAL_EDITOR__DOCS_NORMAL') return true;
  return editor.isVisible().visible || !!document.querySelector('.g-dialog');
}

export function installBasics(api: Api): void {
  const injector = (api as unknown as { _injector: { get<T>(id: unknown): T } })._injector;
  const commands = injector.get<ICommandService>(ICommandService);
  const editor = injector.get<IEditorBridgeService>(IEditorBridgeService);
  injector.get<IconManager>(IconManager).register({
    GrillonSigmaIcon: icon('M12 3H4.5L8.5 8l-4 5H12'),
    GrillonFillIcon: icon('M8 2v9M4.5 8L8 11.5 11.5 8M3 14h10'),
    GrillonFormulaViewIcon: icon('M3 3h10v10H3zM6 10.5l4-5M6 5.5l4 5'),
  } as never);

  // --- Afficher les formules : le texte de la formule remplace le résultat à l'écran ---
  injector.get<SheetInterceptorService>(SheetInterceptorService).intercept(INTERCEPTOR_POINT.CELL_CONTENT, {
    effect: InterceptorEffectEnum.Value,
    priority: 0,
    handler: (cell, _location, next) => (showFormulas && cell?.f ? next({ ...cell, v: cell.f, t: STRING as never }) : next(cell)),
  });
  const toggleFormulas = () => {
    showFormulas = !showFormulas;
    (api.getActiveWorkbook()?.getActiveSheet() as unknown as { refreshCanvas?(): void } | undefined)?.refreshCanvas?.();
  };

  const command = (id: string, run: () => unknown) => {
    if (!commands.hasCommand(id)) commands.registerCommand({ id, type: CommandType.COMMAND, handler: () => (void run(), true) });
  };

  // --- Boutons : Σ et Remplir (Démarrer), Σ et Afficher les formules (Formules), Afficher les formules (Vue) ---
  const sums: [AutoSumFunction, string][] = [
    ['SUM', t('basics.sum')],
    ['AVERAGE', t('basics.average')],
    ['COUNT', t('basics.count')],
    ['MAX', t('basics.max')],
    ['MIN', t('basics.min')],
  ];
  for (const [suffix, path] of [['home', START_OTHERS], ['formulas', FORMULAS_OTHERS]] as const) {
    const id = `grillon.autosum.${suffix}`;
    api.createSubmenu({ id, icon: 'GrillonSigmaIcon', title: t('basics.autoSum'), tooltip: t('basics.autoSumTip') }).appendTo(path);
    sums.forEach(([fn, title], order) => {
      const itemId = `${id}.${fn}`;
      command(itemId, () => autoSum(api, fn));
      api.createMenu({ id: itemId, title, action: itemId, order }).appendTo([...path, id]);
    });
  }

  const fillId = 'grillon.fill';
  api.createSubmenu({ id: fillId, icon: 'GrillonFillIcon', title: t('basics.fill'), tooltip: t('basics.fill') }).appendTo(START_OTHERS);
  ([['down', t('basics.fillDown')], ['right', t('basics.fillRight')], ['up', t('basics.fillUp')], ['left', t('basics.fillLeft')]] as [FillDirection, string][]).forEach(([direction, title], order) => {
    const itemId = `${fillId}.${direction}`;
    command(itemId, () => fill(api, direction));
    api.createMenu({ id: itemId, title, action: itemId, order }).appendTo([...START_OTHERS, fillId]);
  });

  for (const [suffix, path] of [['formulas', FORMULAS_OTHERS], ['view', VIEW_DISPLAY]] as const) {
    command(`grillon.formulas.show.${suffix}`, toggleFormulas);
    api.createMenu({ id: `grillon.formulas.show.${suffix}`, icon: 'GrillonFormulaViewIcon', title: t('basics.showFormulas'), tooltip: t('basics.showFormulasTip'), action: `grillon.formulas.show.${suffix}` }).appendTo(path);
  }

  // --- Raccourcis clavier ---
  window.addEventListener(
    'keydown',
    (e) => {
      const mod = e.ctrlKey || e.metaKey;
      if (e.isComposing) return;
      // Appuyer sur Alt ouvre l'éditeur de cellule d'Univer : pour Alt+= on le referme avant d'écrire la somme.
      const altSum = e.altKey && !mod && e.code === 'Equal' && !document.querySelector('.g-dialog');
      if (!altSum && isTypingInCell(editor)) return;
      const handled = (fn: () => void) => {
        e.preventDefault();
        e.stopPropagation();
        fn();
      };
      const cur = sheetAndRange(api);
      if (!cur) return;
      const { wb, ws, range } = cur;

      if (altSum) {
        return handled(() => {
          if (editor.isVisible().visible) editor.changeVisible({ visible: false, eventType: 4, unitId: wb.getId(), keycode: 27 } as never);
          autoSum(api, 'SUM');
        });
      }
      if (e.altKey) return;
      if (mod && e.shiftKey && NUMBER_FORMATS[e.code]) return handled(() => range.setNumberFormat(NUMBER_FORMATS[e.code]!));
      if (mod && e.code === 'Backquote') return handled(toggleFormulas);
      if (mod && !e.shiftKey && e.key === '-') return handled(() => void cellsDialog(api, 'delete'));
      if (mod && e.key === '+') return handled(() => void cellsDialog(api, 'insert'));
      if (e.code === 'Space' && e.ctrlKey && !e.metaKey && !e.shiftKey) {
        return handled(() => wb.setActiveRange(ws.getRange(0, range.getColumn(), ws.getMaxRows(), range.getLastColumn() - range.getColumn() + 1)));
      }
      if (e.code === 'Space' && e.shiftKey && !mod) {
        return handled(() => wb.setActiveRange(ws.getRange(range.getRow(), 0, range.getLastRow() - range.getRow() + 1, ws.getMaxColumns())));
      }
      if (e.ctrlKey && !e.shiftKey && (e.key === 'Home' || e.key === 'End')) {
        return handled(() => {
          const row = e.key === 'Home' ? 0 : Math.max(0, ws.getLastRow());
          const col = e.key === 'Home' ? 0 : Math.max(0, ws.getLastColumn());
          wb.setActiveRange(ws.getRange(row, col));
          ws.scrollToCell(row, col);
        });
      }
    },
    { capture: true },
  );
}
