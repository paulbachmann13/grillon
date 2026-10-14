/**
 * Onglet Formules : Gestionnaire de noms, Repérer les antécédents / les dépendants ;
 * onglet Données : Valeur cible.
 */
import { IconManager } from '@univerjs/preset-sheets-core';
import { CommandType, ICommandService, type FUniver } from '@univerjs/presets';
import type { FWorkbook, FWorksheet } from '@univerjs/sheets/facade';
import { createElement, type ComponentType } from 'react';
import { extractRefs, refContains, type FormulaRef } from '../data/formula-refs';
import { goalSeek } from '../data/goal-seek';
import { t } from '../i18n';
import { rangeToA1, toA1 } from '../io/xlsx/units';
import type { IWorkbookData } from '../model/univer-types';
import { alertDialog, box, field, openDialog, textInput } from '../ui/dialog';

const FORMULAS = 'ribbon.formulas';
const FORMULAS_OTHERS = 'ribbon.formulas.others';

function icon(path: string): ComponentType<Record<string, unknown>> {
  return (props) =>
    createElement('svg', { viewBox: '0 0 16 16', width: '1em', height: '1em', fill: 'none', stroke: 'currentColor', strokeWidth: 1.3, strokeLinecap: 'round', strokeLinejoin: 'round', ...props }, createElement('path', { d: path }));
}

/** Liste cliquable de références : un clic sélectionne la plage (et active sa feuille). */
function refList(api: FUniver, items: { label: string; sheet: string; range: string }[]): HTMLElement {
  const list = document.createElement('div');
  list.className = 'g-ref-list';
  for (const it of items) {
    const b = document.createElement('button');
    b.type = 'button';
    b.textContent = it.label;
    b.addEventListener('click', () => {
      const wb = api.getActiveWorkbook() as unknown as FWorkbook;
      const ws = wb.getSheetByName(it.sheet);
      if (!ws) return;
      wb.setActiveSheet(ws);
      wb.setActiveRange(ws.getRange(it.range));
    });
    list.append(b);
  }
  return list;
}

function activeCell(api: FUniver): { wb: FWorkbook; ws: FWorksheet; row: number; col: number } | null {
  const wb = api.getActiveWorkbook() as unknown as FWorkbook | null;
  const ws = wb?.getActiveSheet() as unknown as FWorksheet | undefined;
  const r = wb?.getActiveRange();
  if (!wb || !ws || !r) return null;
  return { wb, ws, row: r.getRow(), col: r.getColumn() };
}

async function tracePrecedents(api: FUniver): Promise<void> {
  const a = activeCell(api);
  if (!a) return;
  const formula = a.ws.getRange(a.row, a.col).getFormula();
  if (!formula) return alertDialog(t('formulas.precedents'), t('formulas.noFormula'));
  const sheetName = a.ws.getSheetName();
  const items = extractRefs(formula).map((r: FormulaRef) => {
    const sheet = r.sheet ?? sheetName;
    const range = rangeToA1(r);
    return { label: r.sheet ? `${sheet} ! ${range}` : range, sheet, range };
  });
  const unique = items.filter((it, i) => items.findIndex((o) => o.label === it.label) === i);
  if (!unique.length) return alertDialog(t('formulas.precedents'), t('formulas.noPrecedents'));
  await openDialog({
    title: t('formulas.precedentsOf', { cell: toA1(a.row, a.col) }),
    testId: 'dialog-precedents',
    body: box(t('formulas.clickToSelect'), refList(api, unique)),
    onOk: () => true,
  });
}

async function traceDependents(api: FUniver): Promise<void> {
  const a = activeCell(api);
  if (!a) return;
  const current = a.ws.getSheetName();
  const data = a.wb.save() as unknown as IWorkbookData;
  const items: { label: string; sheet: string; range: string }[] = [];
  for (const id of data.sheetOrder) {
    const sheet = data.sheets[id];
    if (!sheet?.cellData) continue;
    for (const [r, cols] of Object.entries(sheet.cellData)) {
      for (const [c, cell] of Object.entries(cols ?? {})) {
        const f = (cell as { f?: string } | undefined)?.f;
        if (!f) continue;
        const hit = extractRefs(f).some((ref) => (ref.sheet ?? sheet.name) === current && refContains(ref, a.row, a.col));
        if (!hit) continue;
        const range = toA1(Number(r), Number(c));
        items.push({ label: sheet.name === current ? range : `${sheet.name} ! ${range}`, sheet: sheet.name ?? '', range });
      }
    }
  }
  if (!items.length) return alertDialog(t('formulas.dependents'), t('formulas.noDependents'));
  await openDialog({
    title: t('formulas.dependentsOf', { cell: toA1(a.row, a.col) }),
    testId: 'dialog-dependents',
    body: box(t('formulas.clickToSelect'), refList(api, items)),
    onOk: () => true,
  });
}

/** Attend la fin du recalcul (ou un délai maximal). */
function waitCalculation(api: FUniver, timeout = 2000): Promise<void> {
  const formula = (api as unknown as { getFormula(): { calculationEnd(cb: () => void): { dispose(): void } } }).getFormula();
  return new Promise((resolve) => {
    let done = false;
    const finish = () => {
      if (done) return;
      done = true;
      sub.dispose();
      resolve();
    };
    const sub = formula.calculationEnd(finish);
    setTimeout(finish, timeout);
  });
}

async function goalSeekTool(api: FUniver): Promise<void> {
  const a = activeCell(api);
  if (!a) return;
  const setCell = textInput(toA1(a.row, a.col));
  const toValue = textInput('');
  const byChanging = textInput('');
  let summary = '';
  let restore: (() => void) | null = null;
  const ok = await openDialog({
    title: t('formulas.goalSeek'),
    testId: 'dialog-goal-seek',
    body: box(field(t('formulas.setCell'), setCell), field(t('formulas.toValue'), toValue), field(t('formulas.byChanging'), byChanging)),
    onOk: async () => {
      const target = Number(toValue.value.replace(/\s/g, '').replace(',', '.'));
      if (!Number.isFinite(target)) return t('formulas.invalidNumber');
      let targetRange;
      let changeRange;
      try {
        targetRange = a.ws.getRange(setCell.value.replace(/\$/g, '').trim());
        changeRange = a.ws.getRange(byChanging.value.replace(/\$/g, '').trim());
      } catch {
        return t('data.invalidRange', { ref: `${setCell.value} / ${byChanging.value}` });
      }
      if (!targetRange.getFormula()) return t('formulas.mustBeFormula');
      if (changeRange.getFormula()) return t('formulas.mustBeValue');
      const original = changeRange.getCellDataGrid()[0]?.[0]?.v;
      const x0 = Number(original ?? 0) || 0;
      restore = () => changeRange.setValue(original === undefined || original === null ? ({ v: null } as never) : original);
      const evaluate = async (x: number) => {
        changeRange.setValue(x);
        await waitCalculation(api);
        // Valeur brute (getCellData passe par l'affichage à la française : « 4,1 »).
        return Number(targetRange.getCellDataGrid()[0]?.[0]?.v);
      };
      const res = await goalSeek(evaluate, target, x0);
      if (res.found) await evaluate(res.x);
      summary = (res.found ? t('formulas.goalFound', { cell: setCell.value }) : t('formulas.goalNotFound', { cell: setCell.value })) +
        ` ${t('formulas.goalValues', { target: String(target).replace('.', ','), current: String(Number(res.value.toPrecision(12))).replace('.', ',') })}`;
    },
  });
  if (!ok) return;
  const keep = await openDialog({ title: t('formulas.goalSeekStatus'), testId: 'dialog-goal-status', body: box(summary), onOk: () => true });
  if (!keep) (restore as (() => void) | null)?.();
}

export function installFormulasRibbon(api: FUniver): void {
  const injector = (api as unknown as { _injector: { get<T>(id: unknown): T } })._injector;
  const commands = injector.get<ICommandService>(ICommandService);
  injector.get<IconManager>(IconManager).register({
    GrillonNamesIcon: icon('M2 4h12v8H2zM4.5 7h4M4.5 9.5h7'),
    GrillonPrecedentsIcon: icon('M2 4h4v3H2zM10 9h4v3h-4zM6 5.5c3 0 3 5 4 5M8.5 9l1.5 1.5L8.5 12'),
    GrillonDependentsIcon: icon('M10 4h4v3h-4zM2 9h4v3H2zM10 5.5c-3 0-3 5-4 5M7.5 9L6 10.5 7.5 12'),
    GrillonGoalIcon: icon('M8 14A6 6 0 108 2a6 6 0 000 12zM8 11a3 3 0 100-6 3 3 0 000 6zM8 8h.01'),
  } as never);
  const button = (id: string, iconName: string, title: string, run: () => unknown, path: string[]) => {
    if (!commands.hasCommand(id)) commands.registerCommand({ id, type: CommandType.COMMAND, handler: () => (void run(), true) });
    api.createMenu({ id, icon: iconName, title, tooltip: title, action: id }).appendTo(path);
  };
  api.createMenu({ id: 'grillon.formulas.names', icon: 'GrillonNamesIcon', title: t('formulas.nameManager'), tooltip: t('formulas.nameManager'), action: 'sidebar.operation.defined-name' }).appendTo([FORMULAS, FORMULAS_OTHERS]);
  button('grillon.formulas.precedents', 'GrillonPrecedentsIcon', t('formulas.precedents'), () => tracePrecedents(api), [FORMULAS, FORMULAS_OTHERS]);
  button('grillon.formulas.dependents', 'GrillonDependentsIcon', t('formulas.dependents'), () => traceDependents(api), [FORMULAS, FORMULAS_OTHERS]);
  button('grillon.data.goal-seek', 'GrillonGoalIcon', t('formulas.goalSeek'), () => goalSeekTool(api), ['ribbon.data', 'ribbon.data.others']);
}
