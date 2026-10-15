/**
 * Accès aux graphiques de la feuille : lecture des données, recherche, mise à jour, insertion.
 * (Les graphiques sont des « DOM flottants » Univer qui portent leur ChartSpec.)
 */
import type { FUniver } from '@univerjs/presets';
import type { FWorkbook, FWorksheet } from '@univerjs/sheets/facade';
import { currentRegionRect } from '../univer/region';
import { applyKind, kindById, type ChartKind } from './kinds';
import { parseSheetRef, seriesFromRange, type CellScalar, type ChartSpec } from './spec';

export const CHART_COMPONENT = 'GrillonChart';

export interface ChartData {
  chartId: string;
  spec: ChartSpec;
}

type EChartsModule = typeof import('echarts');
let echartsPromise: Promise<EChartsModule> | null = null;
export const loadECharts = () => (echartsPromise ??= import('echarts'));

export const isDark = () => document.documentElement.dataset.theme === 'dark' || (document.documentElement.dataset.theme !== 'light' && matchMedia('(prefers-color-scheme: dark)').matches);

export function workbook(api: FUniver): FWorkbook | null {
  return api.getActiveWorkbook() as unknown as FWorkbook | null;
}

/**
 * Lit les valeurs d'une référence « Feuille!A1:B3 » (aplaties). `display` renvoie le texte affiché
 * (dates, formats) ; `visibleOnly` ignore les lignes et colonnes masquées, comme Excel.
 */
export function readRef(api: FUniver, ref: string, options: { defaultSheet?: string; display?: boolean; visibleOnly?: boolean } = {}): CellScalar[] {
  const parsed = parseSheetRef(ref);
  const wb = workbook(api);
  if (!parsed || !wb) return [];
  const sheetName = parsed.sheet ?? options.defaultSheet;
  const ws = (sheetName ? wb.getSheetByName(sheetName) : wb.getActiveSheet()) as unknown as FWorksheet | null;
  if (!ws) return [];
  const rows = parsed.endRow - parsed.startRow + 1;
  const cols = parsed.endColumn - parsed.startColumn + 1;
  const range = ws.getRange(parsed.startRow, parsed.startColumn, rows, cols);
  const grid = range.getCellDataGrid();
  const shown = options.display ? (range as unknown as { getDisplayValues(): string[][] }).getDisplayValues() : null;
  const model = (ws as unknown as { getSheet(): { getRowVisible(r: number): boolean; getColVisible(c: number): boolean } }).getSheet();
  const out: CellScalar[] = [];
  grid.forEach((row, r) => {
    if (options.visibleOnly && !model.getRowVisible(parsed.startRow + r)) return;
    row.forEach((c, k) => {
      if (options.visibleOnly && !model.getColVisible(parsed.startColumn + k)) return;
      const v = c?.v;
      if (v === undefined || v === null || v === '') return void out.push(null);
      const isNumber = typeof v === 'number' || (c?.t === 2 && typeof v === 'string' && Number.isFinite(Number(v)));
      if (shown && isNumber) {
        const text = shown[r]?.[k];
        return void out.push(text === undefined || text === '' ? Number(v) : text);
      }
      out.push(isNumber ? Number(v) : (v as CellScalar));
    });
  });
  return out;
}

export function findChart(api: FUniver, chartId: string): { ws: FWorksheet; id: string; data: ChartData } | null {
  const wb = workbook(api);
  if (!wb) return null;
  for (const ws of wb.getSheets() as unknown as FWorksheet[]) {
    const doms = (ws as unknown as { getAllFloatDoms(): { id: string; componentKey: string; data?: ChartData }[] }).getAllFloatDoms();
    const dom = doms.find((d) => d.componentKey === CHART_COMPONENT && d.data?.chartId === chartId);
    if (dom?.data) return { ws, id: dom.id, data: dom.data };
  }
  return null;
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

/** Plage de données de la sélection (ou de la zone de données autour de la cellule active). */
export function selectedData(api: FUniver): { ws: FWorksheet; rect: { r1: number; c1: number; r2: number; c2: number }; grid: CellScalar[][] } | null {
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
  return { ws, rect, grid };
}

/** Spécification d'un nouveau graphique construit à partir de la sélection. */
export function specFromSelection(api: FUniver, kind: ChartKind | string): ChartSpec | null {
  const k = typeof kind === 'string' ? kindById(kind) : kind;
  const data = selectedData(api);
  if (!data) return null;
  const series = seriesFromRange(data.ws.getSheetName(), data.grid, { row: data.rect.r1, col: data.rect.c1 }, k.type);
  if (!series.length) return null;
  const base: ChartSpec = { version: 1, type: k.type, series, legend: k.type === 'pie' || k.type === 'doughnut' || series.length > 1 ? 'bottom' : 'none' };
  // Un graphique à une série prend le nom de la série pour titre, comme Excel.
  if (series.length === 1 && series[0]!.nameRef) {
    const name = readRef(api, series[0]!.nameRef, { display: true })[0];
    if (typeof name === 'string' && name) base.title = name;
  }
  return applyKind(base, k);
}

/** Insère un graphique à droite des données, comme Excel. Renvoie son identifiant. */
export function insertChartSpec(api: FUniver, spec: ChartSpec): string | null {
  const data = selectedData(api);
  if (!data) return null;
  const { ws, rect } = data;
  // Les positions incluent les en-têtes de lignes/colonnes.
  const config = (ws as unknown as { getSheet(): { getConfig(): { rowHeader?: { width?: number }; columnHeader?: { height?: number } } } }).getSheet().getConfig();
  let left = config.rowHeader?.width ?? 46;
  for (let c = 0; c <= rect.c2; c++) left += ws.getColumnWidth(c);
  let top = config.columnHeader?.height ?? 20;
  for (let i = 0; i < rect.r1; i++) top += ws.getRowHeight(i);
  const chartId = `chart-${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`;
  (ws as unknown as { addFloatDomToPosition(l: object): unknown }).addFloatDomToPosition({
    componentKey: CHART_COMPONENT,
    initPosition: { startX: left + 24, endX: left + 24 + 480, startY: top + 8, endY: top + 8 + 288 },
    data: { chartId, spec } satisfies ChartData,
    allowTransform: true,
  });
  return chartId;
}

export function insertChart(api: FUniver, kind: ChartKind | string): string | null {
  const spec = specFromSelection(api, kind);
  return spec ? insertChartSpec(api, spec) : null;
}
