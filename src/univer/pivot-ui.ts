/**
 * Tableau croisé dynamique : création (Insérer), modification et actualisation (Données).
 * Le résultat est écrit dans une feuille dédiée ; sa définition est gardée dans les
 * métadonnées de cette feuille (custom.grillon.pivot) pour pouvoir l'actualiser.
 */
import { IconManager } from '@univerjs/preset-sheets-core';
import { CommandType, ICommandService, type FUniver } from '@univerjs/presets';
import type { FWorkbook, FWorksheet } from '@univerjs/sheets/facade';
import { createElement, type ComponentType } from 'react';
import { cellText, type Grid } from '../data/cells';
import { computePivot, distinctValues, FR_PIVOT_LABELS, type PivotConfig, type PivotFunction } from '../data/pivot';
import { t } from '../i18n';
import { columnToLetters, rangeFromA1 } from '../io/xlsx/units';
import type { ICellData } from '../model/univer-types';
import { alertDialog, box, field, group, openDialog, select, textInput, checkbox } from '../ui/dialog';
import { currentRegionRect } from './region';

interface PivotDefinition {
  sourceSheet: string;
  sourceRange: string;
  config: PivotConfig;
  /** Taille de la dernière sortie (pour l'effacer avant d'actualiser). */
  outRows: number;
  outCols: number;
}

const OUT_ROW = 2; // la sortie commence en A3 ; A1 décrit la source et les filtres
const CLEAR: ICellData = { v: null, f: null, p: null, s: null } as unknown as ICellData;

function wbOf(api: FUniver): FWorkbook | null {
  return api.getActiveWorkbook() as unknown as FWorkbook | null;
}

function getDefinition(ws: FWorksheet): PivotDefinition | null {
  const custom = ws.getCustomMetadata() as { grillon?: { pivot?: PivotDefinition } } | undefined;
  return custom?.grillon?.pivot ?? null;
}

function readSource(api: FUniver, def: Pick<PivotDefinition, 'sourceSheet' | 'sourceRange'>): Grid | null {
  const ws = wbOf(api)?.getSheetByName(def.sourceSheet) as unknown as FWorksheet | null;
  if (!ws) return null;
  try {
    const r = rangeFromA1(def.sourceRange.replace(/\$/g, ''));
    return ws.getRange(r.startRow, r.startColumn, r.endRow - r.startRow + 1, r.endColumn - r.startColumn + 1).getCellDataGrid() as Grid;
  } catch {
    return null;
  }
}

function writePivot(api: FUniver, target: FWorksheet, def: PivotDefinition): boolean {
  const grid = readSource(api, def);
  if (!grid) return false;
  const out = computePivot(grid, def.config, FR_PIVOT_LABELS);
  const width = Math.max(1, ...out.map((r) => r.length));
  // Efface l'ancienne sortie.
  if (def.outRows && def.outCols) {
    target.getRange(OUT_ROW, 0, def.outRows, def.outCols).setValues(Array.from({ length: def.outRows }, () => Array.from({ length: def.outCols }, () => CLEAR)));
  }
  const header = [cellText(grid[0]?.[0]) ? grid[0] : []];
  const filters = def.config.filters.map((f) => `${cellText(header[0][f.col]) || columnToLetters(f.col)} = ${f.include.join(', ')}`);
  target.getRange(0, 0, 1, 1).setValues([[{ v: `${t('pivot.source')} : ${def.sourceSheet}!${def.sourceRange}${filters.length ? ` — ${t('pivot.filters')} : ${filters.join(' ; ')}` : ''}`, t: 1, s: { it: 1 } } as ICellData]]);
  target.getRange(OUT_ROW, 0, out.length, width).setValues(out.map((r) => Array.from({ length: width }, (_, i) => r[i] ?? CLEAR)) as ICellData[][]);
  target.setCustomMetadata({ ...(target.getCustomMetadata() ?? {}), grillon: { pivot: { ...def, outRows: out.length, outCols: width } } });
  return true;
}

/** Boîte de création / modification. */
async function pivotDialog(api: FUniver, existing?: { ws: FWorksheet; def: PivotDefinition }): Promise<void> {
  const wb = wbOf(api);
  const ws = wb?.getActiveSheet() as unknown as FWorksheet | undefined;
  if (!wb || !ws) return;
  let sourceSheet = existing?.def.sourceSheet ?? ws.getSheetName();
  let sourceRange = existing?.def.sourceRange;
  if (!sourceRange) {
    const r = wb.getActiveRange();
    let rect = r ? { r1: r.getRow(), c1: r.getColumn(), r2: r.getLastRow(), c2: r.getLastColumn() } : { r1: 0, c1: 0, r2: 0, c2: 0 };
    if (rect.r1 === rect.r2 && rect.c1 === rect.c2) rect = currentRegionRect(ws, rect);
    sourceRange = `$${columnToLetters(rect.c1)}$${rect.r1 + 1}:$${columnToLetters(rect.c2)}$${rect.r2 + 1}`;
  }
  const rangeInput = textInput(`${sourceSheet}!${sourceRange}`);
  const fieldsBox = group(t('pivot.fields'));
  const roles: { role: HTMLSelectElement; fn: HTMLSelectElement; filterBox: HTMLElement; checks: HTMLInputElement[]; values: string[] }[] = [];
  const cfg = existing?.def.config;

  const renderFields = () => {
    fieldsBox.querySelectorAll('.g-pivot-field').forEach((e) => e.remove());
    roles.length = 0;
    const m = /^(?:'?(.*?)'?!)?(.+)$/.exec(rangeInput.value.trim());
    if (m?.[1]) sourceSheet = m[1].replace(/''/g, "'");
    sourceRange = m?.[2] ?? sourceRange;
    const grid = readSource(api, { sourceSheet, sourceRange: sourceRange ?? '' });
    if (!grid?.length) return;
    grid[0].forEach((h, col) => {
      const row = document.createElement('div');
      row.className = 'g-pivot-field';
      const label = document.createElement('span');
      label.textContent = cellText(h) || `${t('data.column', { letter: columnToLetters(col) })}`;
      const initial = cfg?.rows.includes(col) ? 'rows' : cfg?.columns.includes(col) ? 'columns' : cfg?.values.find((v) => v.col === col) ? 'values' : cfg?.filters.find((f) => f.col === col) ? 'filter' : '';
      const role = select(
        [
          ['', '—'],
          ['rows', t('pivot.rows')],
          ['columns', t('pivot.columns')],
          ['values', t('pivot.values')],
          ['filter', t('pivot.filter')],
        ],
        initial,
      );
      const fn = select(
        [
          ['sum', t('data.fnSum')],
          ['count', t('data.fnCount')],
          ['average', t('data.fnAverage')],
          ['min', t('data.fnMin')],
          ['max', t('data.fnMax')],
        ],
        cfg?.values.find((v) => v.col === col)?.fn ?? 'sum',
      );
      const filterBox = document.createElement('div');
      filterBox.className = 'g-pivot-filter';
      const values = distinctValues(grid, col);
      const included = cfg?.filters.find((f) => f.col === col)?.include;
      const checks = values.slice(0, 200).map((v) => {
        const c = checkbox(v, !included || included.includes(v));
        filterBox.append(c.el);
        return c.input;
      });
      const sync = () => {
        fn.hidden = role.value !== 'values';
        filterBox.hidden = role.value !== 'filter';
      };
      role.addEventListener('change', sync);
      sync();
      row.append(label, role, fn, filterBox);
      fieldsBox.append(row);
      roles.push({ role, fn, filterBox, checks, values });
    });
  };
  rangeInput.addEventListener('change', renderFields);
  renderFields();

  await openDialog({
    title: existing ? t('pivot.edit') : t('pivot.create'),
    testId: 'dialog-pivot',
    body: box(field(t('pivot.sourceRange'), rangeInput), fieldsBox),
    onOk: () => {
      const config: PivotConfig = { rows: [], columns: [], values: [], filters: [] };
      roles.forEach((r, col) => {
        if (r.role.value === 'rows') config.rows.push(col);
        else if (r.role.value === 'columns') config.columns.push(col);
        else if (r.role.value === 'values') config.values.push({ col, fn: r.fn.value as PivotFunction });
        else if (r.role.value === 'filter') config.filters.push({ col, include: r.values.filter((_, i) => r.checks[i]?.checked) });
      });
      if (!config.values.length) return t('pivot.needValues');
      const def: PivotDefinition = { sourceSheet, sourceRange: sourceRange ?? '', config, outRows: existing?.def.outRows ?? 0, outCols: existing?.def.outCols ?? 0 };
      if (!readSource(api, def)) return t('data.invalidRange', { ref: rangeInput.value });
      let target = existing?.ws;
      if (!target) {
        const names = new Set((wb.getSheets() as unknown as FWorksheet[]).map((s) => s.getSheetName()));
        let n = 1;
        while (names.has(`${t('pivot.sheetName')}${n}`)) n++;
        target = wb.insertSheet(`${t('pivot.sheetName')}${n}`) as unknown as FWorksheet;
        wb.setActiveSheet(target as never);
      }
      writePivot(api, target, def);
    },
  });
}

function refresh(api: FUniver): void {
  const ws = wbOf(api)?.getActiveSheet() as unknown as FWorksheet | undefined;
  const def = ws && getDefinition(ws);
  if (!ws || !def) return void alertDialog(t('pivot.refresh'), t('pivot.notPivot'));
  if (!writePivot(api, ws, def)) void alertDialog(t('pivot.refresh'), t('pivot.sourceMissing'));
}

function edit(api: FUniver): void {
  const ws = wbOf(api)?.getActiveSheet() as unknown as FWorksheet | undefined;
  const def = ws && getDefinition(ws);
  if (!ws || !def) return void alertDialog(t('pivot.edit'), t('pivot.notPivot'));
  void pivotDialog(api, { ws, def });
}

function icon(path: string): ComponentType<Record<string, unknown>> {
  return (props) =>
    createElement('svg', { viewBox: '0 0 16 16', width: '1em', height: '1em', fill: 'none', stroke: 'currentColor', strokeWidth: 1.3, strokeLinecap: 'round', strokeLinejoin: 'round', ...props }, createElement('path', { d: path }));
}

export function installPivot(api: FUniver): void {
  const injector = (api as unknown as { _injector: { get<T>(id: unknown): T } })._injector;
  const commands = injector.get<ICommandService>(ICommandService);
  injector.get<IconManager>(IconManager).register({ GrillonPivotIcon: icon('M2 2h12v12H2zM2 5.5h12M5.5 2v12M8 8.5h4M10 10.5V6.5') } as never);
  const register = (id: string, run: () => unknown) => {
    if (!commands.hasCommand(id)) commands.registerCommand({ id, type: CommandType.COMMAND, handler: () => (void run(), true) });
  };
  register('grillon.pivot.create', () => pivotDialog(api));
  api.createMenu({ id: 'grillon.pivot.create', icon: 'GrillonPivotIcon', title: t('pivot.short'), tooltip: t('pivot.create'), action: 'grillon.pivot.create' }).appendTo(['ribbon.insert', 'ribbon.insert.media']);
  api.createSubmenu({ id: 'grillon.pivot', icon: 'GrillonPivotIcon', title: t('pivot.short'), tooltip: t('pivot.short') }).appendTo(['ribbon.data', 'ribbon.data.others']);
  const items: [string, string, () => unknown][] = [
    ['grillon.pivot.new', t('pivot.create'), () => pivotDialog(api)],
    ['grillon.pivot.refresh', t('pivot.refresh'), () => refresh(api)],
    ['grillon.pivot.edit', t('pivot.edit'), () => edit(api)],
  ];
  items.forEach(([id, title, run], order) => {
    register(id, run);
    api.createMenu({ id, title, action: id, order }).appendTo(['ribbon.data', 'ribbon.data.others', 'grillon.pivot']);
  });
}
