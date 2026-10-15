/**
 * Mini-graphiques dans les cellules : un « DOM flottant » Univer posé sur une cellule (non
 * déplaçable), qui dessine un petit graphique ECharts des données d'une ligne ou d'une colonne.
 */
import { IconManager } from '@univerjs/preset-sheets-core';
import { CommandType, ICommandService, type FUniver } from '@univerjs/presets';
import type { FWorksheet } from '@univerjs/sheets/facade';
import { createElement, useEffect, useRef, type ComponentType } from 'react';
import { t } from '../i18n';
import { rangeFromA1 } from '../io/xlsx/units';
import { box, checkbox, field, group, openDialog, select, textInput } from '../ui/dialog';
import { isDark, loadECharts, readRef, workbook } from './chart-store';
import { sheetRef } from './spec';
import { SPARK_COLOR, SPARK_HIGHLIGHT, sparkTargets, sparklineOption, type SparklineSpec, type SparklineType } from './sparkline';

export const SPARKLINE_COMPONENT = 'GrillonSparkline';

export interface SparklineData {
  sparkId: string;
  spec: SparklineSpec;
}

const toNumber = (v: unknown): number | null => (typeof v === 'number' && Number.isFinite(v) ? v : typeof v === 'string' && v.trim() !== '' && Number.isFinite(Number(v.replace(',', '.'))) ? Number(v.replace(',', '.')) : null);

export function createSparklineComponent(api: FUniver): ComponentType<{ data?: SparklineData }> {
  return function SparklineView(props: { data?: SparklineData }) {
    const holder = useRef<HTMLDivElement>(null);
    const spec = props.data?.spec;
    const key = JSON.stringify(spec ?? null);
    useEffect(() => {
      if (!spec || !holder.current) return;
      let disposed = false;
      let chart: { setOption(o: never, notMerge: boolean): void; resize(): void; dispose(): void } | null = null;
      let timer: ReturnType<typeof setTimeout> | undefined;
      const render = () => chart?.setOption(sparklineOption(spec, readRef(api, spec.dataRef, { visibleOnly: false }).map(toNumber), isDark()) as never, true);
      const schedule = () => {
        clearTimeout(timer);
        timer = setTimeout(render, 120);
      };
      void loadECharts().then((ec) => {
        if (disposed || !holder.current) return;
        chart = ec.init(holder.current, undefined, { renderer: 'canvas' }) as never;
        render();
      });
      const sub = api.addEvent(api.Event.CommandExecuted, (e) => {
        if (e.type === 2 && /set-range-values|formula-calculation-result|insert-row|remove-row|insert-col|remove-col|move-/.test(e.id)) schedule();
      });
      const ro = new ResizeObserver(() => chart?.resize());
      ro.observe(holder.current);
      return () => {
        disposed = true;
        clearTimeout(timer);
        sub.dispose();
        ro.disconnect();
        chart?.dispose();
      };
    }, [key]);
    // Les clics traversent le mini-graphique : la cellule reste sélectionnable et modifiable.
    return createElement('div', { ref: holder, className: 'g-spark', 'data-testid': 'sparkline' });
  };
}

interface FloatDom {
  id: string;
  componentKey: string;
  data?: SparklineData;
  position?: { left: number; top: number };
}

const floatDoms = (ws: FWorksheet) => (ws as unknown as { getAllFloatDoms(): FloatDom[] }).getAllFloatDoms();

/** Cellule (ligne, colonne) sous un point de la feuille, en pixels (en-têtes compris). */
function cellAtPoint(ws: FWorksheet, x: number, y: number): { row: number; col: number } {
  const config = (ws as unknown as { getSheet(): { getConfig(): { rowHeader?: { width?: number }; columnHeader?: { height?: number } } } }).getSheet().getConfig();
  let col = 0;
  for (let edge = config.rowHeader?.width ?? 46; col < 16384; col++) {
    edge += ws.getColumnWidth(col);
    if (edge > x) break;
  }
  let row = 0;
  for (let edge = config.columnHeader?.height ?? 20; row < 1048576; row++) {
    edge += ws.getRowHeight(row);
    if (edge > y) break;
  }
  return { row, col };
}

/** Mini-graphique posé sur une cellule. */
export function findSparkline(ws: FWorksheet, row: number, col: number): FloatDom | null {
  return (
    floatDoms(ws).find((d) => {
      if (d.componentKey !== SPARKLINE_COMPONENT || !d.position) return false;
      const at = cellAtPoint(ws, d.position.left + 2, d.position.top + 2);
      return at.row === row && at.col === col;
    }) ?? null
  );
}

function removeSparkline(ws: FWorksheet, id: string): void {
  (ws as unknown as { removeFloatDom(id: string): void }).removeFloatDom(id);
}

/** Crée (ou remplace) des mini-graphiques ; chaque cible reçoit un exemplaire de `spec`. */
export function createSparklines(api: FUniver, targets: { row: number; col: number; dataRef: string }[], spec: Omit<SparklineSpec, 'dataRef'>): number {
  const wb = workbook(api);
  const ws = wb?.getActiveSheet() as unknown as FWorksheet | undefined;
  if (!ws) return 0;
  const config = (ws as unknown as { getSheet(): { getConfig(): { rowHeader?: { width?: number }; columnHeader?: { height?: number } } } }).getSheet().getConfig();
  for (const target of targets) {
    const existing = findSparkline(ws, target.row, target.col);
    if (existing) removeSparkline(ws, existing.id);
    let left = config.rowHeader?.width ?? 46;
    for (let c = 0; c < target.col; c++) left += ws.getColumnWidth(c);
    let top = config.columnHeader?.height ?? 20;
    for (let r = 0; r < target.row; r++) top += ws.getRowHeight(r);
    const width = ws.getColumnWidth(target.col);
    const height = ws.getRowHeight(target.row);
    const sparkId = `spark-${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`;
    (ws as unknown as { addFloatDomToPosition(l: object): unknown }).addFloatDomToPosition({
      componentKey: SPARKLINE_COMPONENT,
      // Marge à gauche : on peut cliquer dans cette bande pour sélectionner la cellule (le mini-graphique, lui, capte les clics).
      initPosition: { startX: left + 9, endX: left + width - 2, startY: top + 2, endY: top + height - 2 },
      data: { sparkId, spec: { ...spec, dataRef: target.dataRef } } satisfies SparklineData,
      allowTransform: false,
    });
  }
  return targets.length;
}

const a1 = (r: { r1: number; c1: number; r2: number; c2: number }) => {
  const col = (c: number) => {
    let n = '';
    for (let x = c + 1; x > 0; x = Math.floor((x - 1) / 26)) n = String.fromCharCode(65 + ((x - 1) % 26)) + n;
    return n;
  };
  return r.r1 === r.r2 && r.c1 === r.c2 ? `${col(r.c1)}${r.r1 + 1}` : `${col(r.c1)}${r.r1 + 1}:${col(r.c2)}${r.r2 + 1}`;
};

const parseRect = (text: string) => {
  try {
    const r = rangeFromA1(text.replace(/\$/g, '').replace(/^.*!/, ''));
    return { r1: r.startRow, c1: r.startColumn, r2: r.endRow, c2: r.endColumn };
  } catch {
    return null;
  }
};

/** Boîte « Créer des mini-graphiques » (ou « Modifier » si la cellule active en contient un). */
export async function sparklineDialog(api: FUniver, preset: { type?: SparklineType } = {}): Promise<void> {
  const wb = workbook(api);
  const ws = wb?.getActiveSheet() as unknown as FWorksheet | undefined;
  const range = wb?.getActiveRange();
  if (!wb || !ws || !range) return;
  const sheetName = ws.getSheetName();
  const sel = { r1: range.getRow(), c1: range.getColumn(), r2: range.getLastRow(), c2: range.getLastColumn() };
  const existing = findSparkline(ws, sel.r1, sel.c1);
  const editing = existing?.data?.spec;
  const type = select([['line', t('spark.line')], ['column', t('spark.column')], ['winloss', t('spark.winloss')]], editing?.type ?? preset.type ?? 'line');
  const dataRefText = editing ? (editing.dataRef.replace(/^.*!/, '').replace(/\$/g, '')) : a1(sel);
  const dataInput = textInput(dataRefText);
  // Emplacement par défaut : la colonne à droite des données (une cellule par ligne).
  const defaultLocation = editing ? a1({ r1: sel.r1, c1: sel.c1, r2: sel.r1, c2: sel.c1 }) : a1({ r1: sel.r1, c1: sel.c2 + 1, r2: sel.r2, c2: sel.c2 + 1 });
  const locationInput = textInput(defaultLocation);
  locationInput.disabled = !!editing;
  const color = document.createElement('input');
  color.type = 'color';
  color.value = editing?.color ?? SPARK_COLOR;
  const highlight = document.createElement('input');
  highlight.type = 'color';
  highlight.value = editing?.highlight ?? SPARK_HIGHLIGHT;
  const opts = {
    markers: checkbox(t('spark.markers'), !!editing?.markers),
    high: checkbox(t('spark.high'), !!editing?.high),
    low: checkbox(t('spark.low'), !!editing?.low),
    first: checkbox(t('spark.first'), !!editing?.first),
    last: checkbox(t('spark.last'), !!editing?.last),
    negative: checkbox(t('spark.negative'), !!editing?.negative),
  };
  const empty = select([['gap', t('chartUi.emptyGap')], ['zero', t('chartUi.emptyZero')], ['connect', t('chartUi.emptyConnect')]], editing?.emptyCells ?? 'gap');
  const colors = document.createElement('div');
  colors.className = 'g-row';
  colors.append(field(t('spark.color'), color), field(t('spark.highlight'), highlight));
  let removed = false;
  const body: HTMLElement[] = [field(t('spark.type'), type), field(t('spark.dataRange'), dataInput), field(t('spark.location'), locationInput), group(t('spark.show'), ...Object.values(opts).map((o) => o.el)), colors, field(t('spark.emptyCells'), empty)];
  if (editing && existing) {
    const del = document.createElement('button');
    del.type = 'button';
    del.className = 'g-danger';
    del.textContent = t('spark.delete');
    del.addEventListener('click', () => {
      removed = true;
      removeSparkline(ws, existing.id);
      (del.closest('.g-dialog')?.querySelector('.g-dialog-actions button:not(.primary)') as HTMLButtonElement | null)?.click();
    });
    body.push(del);
  }
  await openDialog({
    title: editing ? t('spark.editTitle') : t('spark.createTitle'),
    testId: 'dialog-sparkline',
    wide: true,
    body: box(...body),
    onOk: () => {
      if (removed) return true;
      const data = parseRect(dataInput.value);
      const location = parseRect(locationInput.value);
      if (!data || !location) return t('spark.badRange');
      const targets = sparkTargets(data, location, (r) => sheetRef(sheetName, r.r1, r.c1, r.r2, r.c2));
      if (!targets) return t('spark.mismatch');
      createSparklines(api, targets, {
        type: type.value as SparklineType,
        color: color.value === SPARK_COLOR ? undefined : color.value,
        highlight: highlight.value === SPARK_HIGHLIGHT ? undefined : highlight.value,
        markers: opts.markers.input.checked || undefined,
        high: opts.high.input.checked || undefined,
        low: opts.low.input.checked || undefined,
        first: opts.first.input.checked || undefined,
        last: opts.last.input.checked || undefined,
        negative: opts.negative.input.checked || undefined,
        emptyCells: empty.value === 'gap' ? undefined : (empty.value as 'zero' | 'connect'),
      });
      return true;
    },
  });
}

/** Mini-graphiques sans boîte de dialogue (Analyse rapide) : une ligne de données par cellule à droite. */
export function quickSparklines(api: FUniver, type: SparklineType): boolean {
  const wb = workbook(api);
  const ws = wb?.getActiveSheet() as unknown as FWorksheet | undefined;
  const range = wb?.getActiveRange();
  if (!ws || !range) return false;
  const data = { r1: range.getRow(), c1: range.getColumn(), r2: range.getLastRow(), c2: range.getLastColumn() };
  const location = { r1: data.r1, c1: data.c2 + 1, r2: data.r2, c2: data.c2 + 1 };
  const targets = sparkTargets(data, location, (r) => sheetRef(ws.getSheetName(), r.r1, r.c1, r.r2, r.c2));
  if (!targets) return false;
  createSparklines(api, targets, { type });
  return true;
}

function icon(path: string): ComponentType<Record<string, unknown>> {
  return (props) =>
    createElement('svg', { viewBox: '0 0 16 16', width: '1em', height: '1em', fill: 'none', stroke: 'currentColor', strokeWidth: 1.3, strokeLinecap: 'round', strokeLinejoin: 'round', ...props }, createElement('path', { d: path }));
}

export function installSparklines(api: FUniver): void {
  const injector = (api as unknown as { _injector: { get<T>(id: unknown): T } })._injector;
  const commands = injector.get<ICommandService>(ICommandService);
  injector.get<IconManager>(IconManager).register({ GrillonSparkIcon: icon('M1.5 11l3-4 3 2.5 3-5 4 4.5') } as never);
  const id = 'grillon.insert.sparkline';
  api.createSubmenu({ id, icon: 'GrillonSparkIcon', title: t('spark.menu'), tooltip: t('spark.menu') }).appendTo(['ribbon.insert', 'ribbon.insert.media']);
  const items: [SparklineType, string][] = [['line', t('spark.line')], ['column', t('spark.column')], ['winloss', t('spark.winloss')]];
  items.forEach(([type, label], order) => {
    const itemId = `${id}.${type}`;
    if (!commands.hasCommand(itemId)) commands.registerCommand({ id: itemId, type: CommandType.COMMAND, handler: () => (void sparklineDialog(api, { type }), true) });
    api.createMenu({ id: itemId, title: label, action: itemId, order }).appendTo(['ribbon.insert', 'ribbon.insert.media', id]);
  });
}

