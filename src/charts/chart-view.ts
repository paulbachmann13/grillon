/**
 * Graphiques dans la feuille : un « DOM flottant » Univer (déplaçable, redimensionnable,
 * enregistré avec le classeur) qui affiche un graphique ECharts. La définition du graphique
 * (ChartSpec) est stockée dans les données du DOM flottant.
 */
import { IconManager } from '@univerjs/preset-sheets-core';
import { CommandType, ICommandService, type FUniver } from '@univerjs/presets';
import type { FWorkbook, FWorksheet } from '@univerjs/sheets/facade';
import { createElement, useEffect, useRef, type ComponentType } from 'react';
import { t } from '../i18n';
import { box, checkbox, field, group, openDialog, select, textInput } from '../ui/dialog';
import { currentRegionRect } from '../univer/region';
import { buildEChartsOption, DEFAULT_COLORS, parseSheetRef, resolveSeries, seriesFromRange, type CellScalar, type ChartSpec, type ChartType } from './spec';

export const CHART_COMPONENT = 'GrillonChart';

export interface ChartData {
  chartId: string;
  spec: ChartSpec;
}

type EChartsModule = typeof import('echarts');
let echartsPromise: Promise<EChartsModule> | null = null;
const loadECharts = () => (echartsPromise ??= import('echarts'));

function workbook(api: FUniver): FWorkbook | null {
  return api.getActiveWorkbook() as unknown as FWorkbook | null;
}

/** Lit les valeurs d'une référence « Feuille!A1:B3 » (valeurs brutes, aplaties). */
export function readRef(api: FUniver, ref: string, defaultSheet?: string): CellScalar[] {
  const parsed = parseSheetRef(ref);
  const wb = workbook(api);
  if (!parsed || !wb) return [];
  const ws = (parsed.sheet ? wb.getSheetByName(parsed.sheet) : defaultSheet ? wb.getSheetByName(defaultSheet) : wb.getActiveSheet()) as unknown as FWorksheet | null;
  if (!ws) return [];
  const grid = ws.getRange(parsed.startRow, parsed.startColumn, parsed.endRow - parsed.startRow + 1, parsed.endColumn - parsed.startColumn + 1).getCellDataGrid();
  return grid.flat().map((c) => {
    const v = c?.v;
    if (v === undefined || v === null || v === '') return null;
    if (c?.t === 2 && typeof v === 'string' && Number.isFinite(Number(v))) return Number(v);
    return v as CellScalar;
  });
}

const isDark = () => document.documentElement.dataset.theme === 'dark' || (document.documentElement.dataset.theme !== 'light' && matchMedia('(prefers-color-scheme: dark)').matches);

export function createChartComponent(api: FUniver): ComponentType<{ data?: ChartData }> {
  return function ChartView(props: { data?: ChartData }) {
    const holder = useRef<HTMLDivElement>(null);
    const spec = props.data?.spec;
    const key = JSON.stringify(spec ?? null);
    useEffect(() => {
      if (!spec || !holder.current) return;
      let disposed = false;
      let chart: ReturnType<EChartsModule['init']> | null = null;
      let timer: ReturnType<typeof setTimeout> | undefined;
      const render = () => {
        if (!chart) return;
        const series = resolveSeries(spec, (ref) => readRef(api, ref));
        chart.setOption(buildEChartsOption(spec, series, isDark()) as never, true);
      };
      const schedule = () => {
        clearTimeout(timer);
        timer = setTimeout(render, 120);
      };
      void loadECharts().then((ec) => {
        if (disposed || !holder.current) return;
        chart = ec.init(holder.current, undefined, { renderer: 'canvas' });
        render();
      });
      // Les données changent : saisie, recalcul des formules, lignes insérées…
      const sub = api.addEvent(api.Event.CommandExecuted, (e) => {
        if (e.type === 2 && /set-range-values|formula-calculation-result|insert-row|remove-row|insert-col|remove-col|move-/.test(e.id)) schedule();
      });
      const ro = new ResizeObserver(() => chart?.resize());
      ro.observe(holder.current);
      const mq = matchMedia('(prefers-color-scheme: dark)');
      mq.addEventListener('change', schedule);
      return () => {
        disposed = true;
        clearTimeout(timer);
        sub.dispose();
        ro.disconnect();
        mq.removeEventListener('change', schedule);
        chart?.dispose();
      };
    }, [key]);
    return createElement(
      'div',
      { className: 'g-chart', 'data-testid': 'chart' },
      createElement('div', { ref: holder, className: 'g-chart-canvas' }),
      createElement(
        'button',
        {
          type: 'button',
          className: 'g-chart-edit',
          title: t('charts.edit'),
          'aria-label': t('charts.edit'),
          onPointerDown: (e: Event) => e.stopPropagation(),
          onClick: () => props.data && void editChart(api, props.data.chartId),
        },
        '⚙',
      ),
    );
  };
}

function findChart(api: FUniver, chartId: string): { ws: FWorksheet; id: string; data: ChartData } | null {
  const wb = workbook(api);
  if (!wb) return null;
  for (const ws of wb.getSheets() as unknown as FWorksheet[]) {
    const doms = (ws as unknown as { getAllFloatDoms(): { id: string; componentKey: string; data?: ChartData }[] }).getAllFloatDoms();
    const dom = doms.find((d) => d.componentKey === CHART_COMPONENT && d.data?.chartId === chartId);
    if (dom?.data) return { ws, id: dom.id, data: dom.data };
  }
  return null;
}

const TYPE_LABELS = (): [ChartType, string][] => [
  ['column', t('charts.column')],
  ['bar', t('charts.bar')],
  ['line', t('charts.line')],
  ['area', t('charts.area')],
  ['pie', t('charts.pie')],
  ['scatter', t('charts.scatter')],
];

/** Boîte « Modifier le graphique » : type, titres, légende, séries (couleur, type, axe secondaire). */
export async function editChart(api: FUniver, chartId: string): Promise<void> {
  const found = findChart(api, chartId);
  if (!found) return;
  const spec: ChartSpec = structuredClone(found.data.spec);
  const type = select(TYPE_LABELS(), spec.type);
  const title = textInput(spec.title ?? '');
  const legend = select(
    [
      ['bottom', t('charts.legendBottom')],
      ['top', t('charts.legendTop')],
      ['right', t('charts.legendRight')],
      ['none', t('charts.legendNone')],
    ],
    spec.legend ?? 'bottom',
  );
  const xTitle = textInput(spec.xAxisTitle ?? '');
  const yTitle = textInput(spec.yAxisTitle ?? '');
  const stacked = checkbox(t('charts.stacked'), !!spec.stacked);
  const seriesRows = spec.series.map((s, i) => {
    const name = textInput(s.name ?? '', s.nameRef ?? '');
    const color = document.createElement('input');
    color.type = 'color';
    color.value = s.color ?? DEFAULT_COLORS[i % DEFAULT_COLORS.length];
    const stype = select([['', t('charts.sameType')], ...TYPE_LABELS().filter(([v]) => v !== 'pie' && v !== 'scatter')], s.type ?? '');
    const secondary = checkbox(t('charts.secondaryAxis'), !!s.secondaryAxis);
    const row = document.createElement('div');
    row.className = 'g-series-row';
    row.append(name, color, stype, secondary.el);
    return { row, name, color, stype, secondary };
  });
  let removed = false;
  const del = document.createElement('button');
  del.type = 'button';
  del.className = 'g-danger';
  del.textContent = t('charts.delete');
  del.addEventListener('click', () => {
    removed = true;
    removeChart(api, chartId);
    // Ferme la boîte (comme « Annuler ») : le graphique n'existe plus.
    (del.closest('.g-dialog')?.querySelector('.g-dialog-actions button:not(.primary)') as HTMLButtonElement | null)?.click();
  });
  const ok = await openDialog({
    title: t('charts.edit'),
    testId: 'dialog-chart',
    body: box(
      field(t('charts.type'), type),
      field(t('charts.title'), title),
      field(t('charts.legend'), legend),
      field(t('charts.xAxisTitle'), xTitle),
      field(t('charts.yAxisTitle'), yTitle),
      stacked.el,
      group(t('charts.series'), ...seriesRows.map((r) => r.row)),
      del,
    ),
    onOk: () => true,
  });
  if (!ok || removed) return;
  spec.type = type.value as ChartType;
  spec.title = title.value || undefined;
  spec.legend = legend.value as ChartSpec['legend'];
  spec.xAxisTitle = xTitle.value || undefined;
  spec.yAxisTitle = yTitle.value || undefined;
  spec.stacked = stacked.input.checked || undefined;
  seriesRows.forEach((r, i) => {
    const s = spec.series[i];
    if (r.name.value) s.name = r.name.value;
    else delete s.name;
    s.color = r.color.value;
    s.type = (r.stype.value || undefined) as ChartType | undefined;
    s.secondaryAxis = r.secondary.input.checked || undefined;
  });
  updateChart(api, chartId, spec);
}

export function updateChart(api: FUniver, chartId: string, spec: ChartSpec): void {
  const found = findChart(api, chartId);
  if (!found) return;
  (found.ws as unknown as { updateFloatDom(id: string, c: object): void }).updateFloatDom(found.id, { data: { chartId, spec } });
}

export function removeChart(api: FUniver, chartId: string): void {
  const found = findChart(api, chartId);
  if (found) (found.ws as unknown as { removeFloatDom(id: string): void }).removeFloatDom(found.id);
}

/** Insère un graphique à partir de la sélection (ou de la zone de données en cours). */
export function insertChart(api: FUniver, type: ChartType): string | null {
  const wb = workbook(api);
  const ws = wb?.getActiveSheet() as unknown as FWorksheet | undefined;
  const r = wb?.getActiveRange();
  if (!wb || !ws || !r) return null;
  let rect = { r1: r.getRow(), c1: r.getColumn(), r2: r.getLastRow(), c2: r.getLastColumn() };
  if (rect.r1 === rect.r2 && rect.c1 === rect.c2) rect = currentRegionRect(ws, rect);
  const grid = ws
    .getRange(rect.r1, rect.c1, rect.r2 - rect.r1 + 1, rect.c2 - rect.c1 + 1)
    .getCellDataGrid()
    .map((row) => row.map((c) => (c?.v === undefined || c?.v === null || c?.v === '' ? null : c?.t === 2 && typeof c.v === 'string' ? Number(c.v) : (c.v as CellScalar))));
  const series = seriesFromRange(ws.getSheetName(), grid, { row: rect.r1, col: rect.c1 }, type);
  if (!series.length) return null;
  const spec: ChartSpec = { version: 1, type, series, legend: type === 'pie' || series.length > 1 ? 'bottom' : 'none' };
  // À droite des données, comme Excel. Les positions incluent les en-têtes de lignes/colonnes.
  const config = (ws as unknown as { getSheet(): { getConfig(): { rowHeader?: { width?: number }; columnHeader?: { height?: number } } } }).getSheet().getConfig();
  let left = config.rowHeader?.width ?? 46;
  for (let c = 0; c <= rect.c2; c++) left += ws.getColumnWidth(c);
  let top = config.columnHeader?.height ?? 20;
  for (let i = 0; i < rect.r1; i++) top += ws.getRowHeight(i);
  const chartId = `chart-${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`;
  const data: ChartData = { chartId, spec };
  (ws as unknown as { addFloatDomToPosition(l: object): unknown }).addFloatDomToPosition({
    componentKey: CHART_COMPONENT,
    initPosition: { startX: left + 24, endX: left + 24 + 480, startY: top + 8, endY: top + 8 + 288 },
    data,
    allowTransform: true,
  });
  return chartId;
}

function icon(path: string): ComponentType<Record<string, unknown>> {
  return (props) =>
    createElement('svg', { viewBox: '0 0 16 16', width: '1em', height: '1em', fill: 'none', stroke: 'currentColor', strokeWidth: 1.3, strokeLinecap: 'round', strokeLinejoin: 'round', ...props }, createElement('path', { d: path }));
}

export function installCharts(api: FUniver): void {
  const injector = (api as unknown as { _injector: { get<T>(id: unknown): T } })._injector;
  const commands = injector.get<ICommandService>(ICommandService);
  injector.get<IconManager>(IconManager).register({ GrillonChartIcon: icon('M2 14h12M4 11V7M7 11V3M10 11V8M13 11V5') } as never);
  const id = 'grillon.insert.chart';
  api.createSubmenu({ id, icon: 'GrillonChartIcon', title: t('charts.insert'), tooltip: t('charts.insert') }).appendTo(['ribbon.insert', 'ribbon.insert.media']);
  TYPE_LABELS().forEach(([type, label], order) => {
    const itemId = `grillon.insert.chart.${type}`;
    if (!commands.hasCommand(itemId)) commands.registerCommand({ id: itemId, type: CommandType.COMMAND, handler: () => insertChart(api, type) !== null });
    api.createMenu({ id: itemId, title: label, action: itemId, order }).appendTo(['ribbon.insert', 'ribbon.insert.media', id]);
  });
}
