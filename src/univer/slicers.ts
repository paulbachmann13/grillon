/**
 * Segments et chronologies (Insertion → Filtres) : des « DOM flottants » Univer qui filtrent un
 * tableau (en masquant les lignes) ou un tableau croisé dynamique (en posant un filtre de champ).
 * La définition est stockée dans les données du DOM flottant, donc enregistrée avec le classeur.
 */
import { IconManager } from '@univerjs/preset-sheets-core';
import { CommandType, ICommandService, type FUniver } from '@univerjs/presets';
import type { FWorkbook, FWorksheet } from '@univerjs/sheets/facade';
import { createElement, useEffect, useState, type ComponentType } from 'react';
import { cellText, cellValue, type Grid } from '../data/cells';
import {
  clickItem,
  clickPeriod,
  EMPTY_LABEL,
  filterOf,
  isOpen,
  pivotInclude,
  periodLabel,
  slicerItems,
  timelinePeriods,
  visibleRows,
  type ColumnCell,
  type SlicerFilter,
  type SlicerSpec,
  type SlicerTarget,
  type TimelineLevel,
} from '../data/slicer';
import type { PivotFilter } from '../data/pivot';
import { t } from '../i18n';
import { columnToLetters, rangeFromA1 } from '../io/xlsx/units';
import { GRILLON_CUSTOM_KEY } from '../model/univer-types';
import { alertDialog, box, field, openDialog, select } from '../ui/dialog';
import { getDefinition, readSource, writePivot } from './pivot-ui';
import { currentRegionRect } from './region';

export const SLICER_COMPONENT = 'GrillonSlicer';
const LEVELS: TimelineLevel[] = ['years', 'quarters', 'months', 'days'];

type FloatDom = { id: string; componentKey: string; data?: SlicerSpec };
type FloatSheet = FWorksheet;
interface FloatApi {
  getAllFloatDoms(): FloatDom[];
  updateFloatDom(id: string, config: object): void;
  addFloatDomToPosition(config: object): unknown;
}
const float = (ws: FWorksheet) => ws as unknown as FloatApi;
type Custom = { [GRILLON_CUSTOM_KEY]?: { slicerHidden?: Record<string, number[]> } & Record<string, unknown> };
interface CoreSheet {
  getRowVisible(r: number): boolean;
}

const wbOf = (api: FUniver) => api.getActiveWorkbook() as unknown as FWorkbook | null;
const sheetsOf = (api: FUniver) => (wbOf(api)?.getSheets() ?? []) as unknown as FloatSheet[];

export function listSlicers(api: FUniver): { ws: FloatSheet; domId: string; spec: SlicerSpec }[] {
  const out: { ws: FloatSheet; domId: string; spec: SlicerSpec }[] = [];
  for (const ws of sheetsOf(api)) for (const d of float(ws).getAllFloatDoms?.() ?? []) if (d.componentKey === SLICER_COMPONENT && d.data) out.push({ ws, domId: d.id, spec: d.data });
  return out;
}

/** Clé commune aux segments qui filtrent la même chose. */
export const groupKey = (t: SlicerTarget) => (t.kind === 'table' ? `t|${t.sheet}|${t.ref.replace(/\$/g, '')}` : `p|${t.sheet}`);

// --- Lecture des données ---

interface Loaded {
  ws: FWorksheet;
  headers: string[];
  rows: ColumnCell[][];
  /** Ligne de la feuille de la première ligne de données (tableau). */
  firstRow: number;
}

export function loadTarget(api: FUniver, target: SlicerTarget): Loaded | null {
  const wb = wbOf(api);
  const ws = wb?.getSheetByName(target.sheet) as unknown as FWorksheet | null;
  if (!ws) return null;
  try {
    if (target.kind === 'table') {
      const r = rangeFromA1(target.ref.replace(/\$/g, ''));
      const range = ws.getRange(r.startRow, r.startColumn, r.endRow - r.startRow + 1, r.endColumn - r.startColumn + 1);
      const grid = range.getCellDataGrid() as Grid;
      const shown = (range as unknown as { getDisplayValues(): string[][] }).getDisplayValues();
      const headers = (grid[0] ?? []).map((c, j) => cellText(c) || t('data.column', { letter: columnToLetters(r.startColumn + j) }));
      const rows = grid.slice(1).map((row, i) => row.map((c, j) => ({ label: shown[i + 1]?.[j] ?? cellText(c), value: cellValue(c) })));
      return { ws, headers, rows, firstRow: r.startRow + 1 };
    }
    const def = getDefinition(ws);
    const grid = def && readSource(api, def);
    if (!def || !grid) return null;
    const headers = (grid[0] ?? []).map((c, j) => cellText(c) || t('data.column', { letter: columnToLetters(j) }));
    return { ws, headers, rows: grid.slice(1).map((row) => row.map((c) => ({ label: cellText(c), value: cellValue(c) }))), firstRow: 0 };
  } catch {
    return null;
  }
}

/** Colonne qui contient des dates (numéros de série d'une période plausible, nombres entiers). */
export function looksLikeDates(rows: readonly (readonly ColumnCell[])[], col: number): boolean {
  const values = rows.map((r) => r[col]?.value).filter((v) => v !== null && v !== undefined && v !== '');
  return values.length > 0 && values.every((v) => typeof v === 'number' && Number.isInteger(v) && v >= 20000 && v <= 80000);
}

// --- Application des filtres ---

function managedRows(ws: FWorksheet, key: string): Set<number> {
  const custom = (ws.getCustomMetadata() ?? {}) as Custom;
  return new Set(custom[GRILLON_CUSTOM_KEY]?.slicerHidden?.[key] ?? []);
}

function saveManaged(ws: FWorksheet, key: string, rows: Set<number>): void {
  const custom = (ws.getCustomMetadata() ?? {}) as Custom;
  const grillon = { ...(custom[GRILLON_CUSTOM_KEY] ?? {}) };
  const map = { ...(grillon.slicerHidden ?? {}) };
  if (rows.size) map[key] = [...rows].sort((a, b) => a - b);
  else delete map[key];
  if (Object.keys(map).length) grillon.slicerHidden = map;
  else delete grillon.slicerHidden;
  const next = { ...custom, [GRILLON_CUSTOM_KEY]: grillon };
  if (JSON.stringify(next) !== JSON.stringify(custom)) ws.setCustomMetadata(next);
}

/** Masque ou affiche des lignes par séries contiguës. */
function setRowsHidden(ws: FWorksheet, rows: number[], hide: boolean): void {
  const sorted = [...new Set(rows)].sort((a, b) => a - b);
  let i = 0;
  while (i < sorted.length) {
    let j = i;
    while (j + 1 < sorted.length && sorted[j + 1] === sorted[j] + 1) j++;
    if (hide) ws.hideRows(sorted[i], j - i + 1);
    else ws.showRows(sorted[i], j - i + 1);
    i = j + 1;
  }
}

function applyTable(api: FUniver, key: string, specs: SlicerSpec[]): void {
  const loaded = loadTarget(api, specs[0].target);
  if (!loaded) return;
  const filters = specs.map(filterOf);
  const mask = visibleRows(loaded.rows, filters);
  const want = new Set<number>();
  mask.forEach((visible, i) => !visible && want.add(loaded.firstRow + i));
  const managed = managedRows(loaded.ws, key);
  const core = loaded.ws.getSheet() as unknown as CoreSheet;
  setRowsHidden(loaded.ws, [...want].filter((r) => core.getRowVisible(r)), true);
  setRowsHidden(loaded.ws, [...managed].filter((r) => !want.has(r)), false);
  saveManaged(loaded.ws, key, want);
}

function applyPivot(api: FUniver, sheet: string, specs: SlicerSpec[]): void {
  const ws = wbOf(api)?.getSheetByName(sheet) as unknown as FWorksheet | null;
  const def = ws && getDefinition(ws);
  const loaded = def && loadTarget(api, specs[0].target);
  if (!ws || !def || !loaded) return;
  const byCol = new Map<number, string[]>();
  for (const spec of specs) {
    const include = pivotInclude(loaded.rows.map((r) => r[spec.target.col] ?? { label: '', value: null }), filterOf(spec));
    if (!include) continue;
    const before = byCol.get(spec.target.col);
    byCol.set(spec.target.col, before ? before.filter((l) => include.includes(l)) : include);
  }
  rewritePivotFilters(api, ws, def, byCol);
}

function rewritePivotFilters(api: FUniver, ws: FWorksheet, def: NonNullable<ReturnType<typeof getDefinition>>, byCol: Map<number, string[]>): void {
  const owned = new Set([...(def.slicerCols ?? []), ...byCol.keys()]);
  const filters: PivotFilter[] = [...def.config.filters.filter((f) => !owned.has(f.col)), ...[...byCol].map(([col, include]) => ({ col, include }))];
  const slicerCols = [...byCol.keys()].sort((a, b) => a - b);
  const same = JSON.stringify(filters) === JSON.stringify(def.config.filters) && JSON.stringify(slicerCols) === JSON.stringify(def.slicerCols ?? []);
  if (same) return;
  writePivot(api, ws, { ...def, config: { ...def.config, filters }, slicerCols });
}

let applying = false;

/** Remet les tableaux et tableaux croisés en accord avec les segments et chronologies de la feuille. */
export function applySlicers(api: FUniver): void {
  if (applying) return;
  applying = true;
  try {
    const groups = new Map<string, SlicerSpec[]>();
    for (const { spec } of listSlicers(api)) groups.set(groupKey(spec.target), [...(groups.get(groupKey(spec.target)) ?? []), spec]);
    for (const [key, specs] of groups) {
      if (specs[0].target.kind === 'table') applyTable(api, key, specs);
      else applyPivot(api, specs[0].target.sheet, specs);
    }
    // Segments supprimés : on rend leurs lignes et leurs filtres.
    for (const ws of sheetsOf(api)) {
      const custom = (ws.getCustomMetadata() ?? {}) as Custom;
      for (const key of Object.keys(custom[GRILLON_CUSTOM_KEY]?.slicerHidden ?? {})) {
        if (groups.has(key)) continue;
        setRowsHidden(ws, [...managedRows(ws, key)], false);
        saveManaged(ws, key, new Set());
      }
      const def = getDefinition(ws);
      if (def?.slicerCols?.length && !groups.has(`p|${ws.getSheetName()}`)) rewritePivotFilters(api, ws, def, new Map());
    }
  } finally {
    applying = false;
  }
}

export function updateSlicer(api: FUniver, id: string, patch: Partial<SlicerSpec>): void {
  const found = listSlicers(api).find((s) => s.spec.id === id);
  if (!found) return;
  float(found.ws).updateFloatDom(found.domId, { data: { ...found.spec, ...patch } });
  applySlicers(api);
}

// --- Composants ---

interface Model {
  filter: SlicerFilter;
  items?: { label: string; count: number }[];
  periods?: string[];
  headers: string[];
  missing: boolean;
}

function modelOf(api: FUniver, spec: SlicerSpec): Model {
  const loaded = loadTarget(api, spec.target);
  if (!loaded) return { filter: filterOf(spec), headers: [], missing: true };
  const key = groupKey(spec.target);
  const specs = listSlicers(api).map((s) => s.spec).filter((s) => groupKey(s.target) === key);
  const index = Math.max(0, specs.findIndex((s) => s.id === spec.id));
  const filters = specs.map(filterOf);
  filters[index] = filterOf(spec);
  if (spec.type === 'timeline') return { filter: filters[index], periods: timelinePeriods(loaded.rows.map((r) => r[spec.target.col]?.value ?? null), spec.level ?? 'months'), headers: loaded.headers, missing: false };
  return { filter: filters[index], items: slicerItems(loaded.rows, filters, index), headers: loaded.headers, missing: false };
}

const stop = (e: Event) => e.stopPropagation();

export function createSlicerComponent(api: FUniver): ComponentType<{ data?: SlicerSpec }> {
  return function SlicerView(props: { data?: SlicerSpec }) {
    const spec = props.data;
    const [tick, setTick] = useState(0);
    useEffect(() => {
      let timer: ReturnType<typeof setTimeout> | undefined;
      const sub = api.addEvent(api.Event.CommandExecuted, (e) => {
        if (e.type === 2 && /set-range-values|formula-calculation-result|insert-row|remove-row|insert-col|remove-col|row-visible|rows-visible|row-hidden/.test(e.id)) {
          clearTimeout(timer);
          timer = setTimeout(() => setTick((n) => n + 1), 150);
        }
      });
      return () => {
        clearTimeout(timer);
        sub.dispose();
      };
    }, []);
    if (!spec) return null;
    void tick;
    const model = modelOf(api, spec);
    const open = isOpen(model.filter);
    const clear = createElement(
      'button',
      {
        type: 'button',
        className: 'g-slicer-clear',
        title: t('slicer.clear'),
        'aria-label': t('slicer.clear'),
        'data-testid': 'slicer-clear',
        disabled: open,
        onPointerDown: stop,
        onClick: () => updateSlicer(api, spec.id, spec.type === 'timeline' ? { from: null, to: null } : { selected: null }),
      },
      '✕',
    );
    const header = createElement('div', { className: 'g-slicer-head' }, createElement('span', { className: 'g-slicer-title' }, spec.title), clear);
    if (model.missing) return createElement('div', { className: 'g-slicer', 'data-testid': 'slicer' }, header, createElement('p', { className: 'g-slicer-empty' }, t('slicer.missing')));

    if (spec.type === 'timeline') {
      const level = spec.level ?? 'months';
      const range = { from: spec.from ?? null, to: spec.to ?? null };
      const picker = createElement(
        'select',
        { className: 'g-slicer-level', value: level, 'data-testid': 'timeline-level', onPointerDown: stop, onChange: (e: { target: { value: string } }) => updateSlicer(api, spec.id, { level: e.target.value as TimelineLevel, from: null, to: null }) },
        ...LEVELS.map((l) => createElement('option', { key: l, value: l }, t(`slicer.level.${l}`))),
      );
      return createElement(
        'div',
        { className: 'g-slicer g-timeline', 'data-testid': 'timeline' },
        header,
        createElement('div', { className: 'g-slicer-sub' }, createElement('span', { 'data-testid': 'timeline-range' }, range.from === null ? t('slicer.allPeriods') : range.from === range.to ? periodLabel(range.from) : `${periodLabel(range.from)} – ${periodLabel(range.to!)}`), picker),
        createElement(
          'div',
          { className: 'g-timeline-strip' },
          ...(model.periods ?? []).map((key) =>
            createElement(
              'button',
              {
                key,
                type: 'button',
                'data-testid': 'timeline-period',
                'data-period': key,
                className: range.from !== null && key >= range.from && key <= (range.to ?? range.from) ? 'selected' : '',
                onPointerDown: stop,
                onClick: (e: { shiftKey: boolean }) => updateSlicer(api, spec.id, clickPeriod(range, key, e.shiftKey)),
              },
              periodLabel(key),
            ),
          ),
        ),
      );
    }

    const all = (model.items ?? []).map((i) => i.label);
    return createElement(
      'div',
      { className: 'g-slicer', 'data-testid': 'slicer' },
      header,
      createElement(
        'div',
        { className: 'g-slicer-items' },
        ...(model.items ?? []).map((item) => {
          const selected = spec.selected === null || spec.selected.includes(item.label);
          return createElement(
            'button',
            {
              key: item.label,
              type: 'button',
              'data-testid': 'slicer-item',
              className: `${selected ? 'selected' : ''}${item.count === 0 ? ' empty' : ''}`,
              onPointerDown: stop,
              onClick: (e: { metaKey: boolean; ctrlKey: boolean; shiftKey: boolean }) => updateSlicer(api, spec.id, { selected: clickItem(spec.selected, item.label, all, e.metaKey || e.ctrlKey || e.shiftKey) }),
            },
            item.label,
          );
        }),
      ),
    );
  };
}

// --- Insertion ---

interface Context {
  ws: FloatSheet;
  target: Omit<Extract<SlicerTarget, { kind: 'table' }>, 'col'> | Omit<Extract<SlicerTarget, { kind: 'pivot' }>, 'col'>;
  headers: string[];
  dateCols: boolean[];
  /** Position (rectangle de la plage ou de la sortie du tableau croisé) pour placer l'objet à droite. */
  rect: { r1: number; c1: number; r2: number; c2: number };
}

function contextOf(api: FUniver): Context | null {
  const wb = wbOf(api);
  const ws = wb?.getActiveSheet() as unknown as FloatSheet | undefined;
  const sel = wb?.getActiveRange();
  if (!wb || !ws || !sel) return null;
  const def = getDefinition(ws);
  let target: Context['target'];
  let rect: Context['rect'];
  if (def) {
    target = { kind: 'pivot', sheet: ws.getSheetName() };
    const top = def.outRow ?? 2;
    const left = def.outCol ?? 0;
    rect = { r1: top, c1: left, r2: top + Math.max(0, def.outRows - 1), c2: left + Math.max(0, def.outCols - 1) };
  } else {
    rect = { r1: sel.getRow(), c1: sel.getColumn(), r2: sel.getLastRow(), c2: sel.getLastColumn() };
    if (rect.r1 === rect.r2 && rect.c1 === rect.c2) rect = currentRegionRect(ws, rect);
    target = { kind: 'table', sheet: ws.getSheetName(), ref: `${columnToLetters(rect.c1)}${rect.r1 + 1}:${columnToLetters(rect.c2)}${rect.r2 + 1}` };
  }
  const loaded = loadTarget(api, { ...target, col: 0 } as SlicerTarget);
  if (!loaded || !loaded.rows.length) return null;
  return { ws, target, headers: loaded.headers, dateCols: loaded.headers.map((_, c) => looksLikeDates(loaded.rows, c)), rect };
}

async function insertDialog(api: FUniver, type: 'slicer' | 'timeline'): Promise<void> {
  const ctx = contextOf(api);
  if (!ctx) return void alertDialog(t(`slicer.insert.${type}`), t('slicer.noData'));
  const candidates = ctx.headers.map((h, i) => ({ h, i })).filter(({ i }) => type === 'slicer' || ctx.dateCols[i]);
  if (!candidates.length) return void alertDialog(t(`slicer.insert.${type}`), t('slicer.noDates'));
  const picker = select(candidates.map(({ h, i }) => [String(i), h] as [string, string]), String(candidates[0].i));
  picker.dataset.testid = 'slicer-field';
  const ok = await openDialog({ title: t(`slicer.insert.${type}`), testId: 'dialog-slicer', body: box(t(`slicer.pick.${type}`), field(t('slicer.field'), picker)), onOk: () => true });
  if (!ok) return;
  const col = Number(picker.value);
  const config = (ctx.ws as unknown as { getSheet(): { getConfig(): { rowHeader?: { width?: number }; columnHeader?: { height?: number } } } }).getSheet().getConfig();
  let left = config.rowHeader?.width ?? 46;
  for (let c = 0; c <= ctx.rect.c2; c++) left += ctx.ws.getColumnWidth(c);
  let top = config.columnHeader?.height ?? 20;
  for (let r = 0; r < ctx.rect.r1; r++) top += ctx.ws.getRowHeight(r);
  const width = type === 'timeline' ? 520 : 200;
  const height = type === 'timeline' ? 120 : 260;
  // Plusieurs segments : côte à côte, sans se recouvrir.
  const offset = listSlicers(api)
    .filter((s) => s.ws.getSheetId() === ctx.ws.getSheetId())
    .reduce((sum, s) => sum + (s.spec.type === 'timeline' ? 520 : 200) + 12, 0);
  const spec: SlicerSpec = {
    id: `slicer-${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`,
    type,
    title: ctx.headers[col],
    target: { ...ctx.target, col } as SlicerTarget,
    selected: null,
    ...(type === 'timeline' ? { level: 'months' as TimelineLevel, from: null, to: null } : {}),
  };
  float(ctx.ws).addFloatDomToPosition({
    componentKey: SLICER_COMPONENT,
    initPosition: { startX: left + 24 + offset, endX: left + 24 + offset + width, startY: top + 8, endY: top + 8 + height },
    data: spec,
    allowTransform: true,
  });
}

function icon(path: string): ComponentType<Record<string, unknown>> {
  return (props) =>
    createElement('svg', { viewBox: '0 0 16 16', width: '1em', height: '1em', fill: 'none', stroke: 'currentColor', strokeWidth: 1.3, strokeLinecap: 'round', strokeLinejoin: 'round', ...props }, createElement('path', { d: path }));
}

export function installSlicers(api: FUniver): void {
  const injector = (api as unknown as { _injector: { get<T>(id: unknown): T } })._injector;
  const commands = injector.get<ICommandService>(ICommandService);
  injector.get<IconManager>(IconManager).register({
    GrillonSlicerIcon: icon('M2 2.5h12v3H2zM2 7h12v3H2zM2 11.5h7v2.5H2zM11 12l1.3 1.3L14.5 11'),
    GrillonTimelineIcon: icon('M2 4h12v8H2zM2 8h12M5 4v8M9 4v8M12 4v8'),
  } as never);
  const register = (id: string, run: () => unknown) => {
    if (!commands.hasCommand(id)) commands.registerCommand({ id, type: CommandType.COMMAND, handler: () => (void run(), true) });
  };
  for (const [type, iconName] of [['slicer', 'GrillonSlicerIcon'], ['timeline', 'GrillonTimelineIcon']] as const) {
    const id = `grillon.insert.${type}`;
    register(id, () => insertDialog(api, type));
    api.createMenu({ id, icon: iconName, title: t(`slicer.insert.${type}`), tooltip: t(`slicer.insert.${type}`), action: id }).appendTo(['ribbon.insert', 'ribbon.insert.media']);
  }
  // Un segment supprimé (ou un classeur chargé) : les filtres suivent.
  let timer: ReturnType<typeof setTimeout> | undefined;
  api.addEvent(api.Event.CommandExecuted, (e) => {
    if (/drawing|float-dom|floatdom/i.test(e.id)) {
      clearTimeout(timer);
      timer = setTimeout(() => applySlicers(api), 100);
    }
  });
}

export { EMPTY_LABEL };
