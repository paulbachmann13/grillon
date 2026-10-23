/**
 * Graphiques : DrawingML (xl/drawings + xl/charts) ⇄ graphiques Grillon (DOM flottants Univer,
 * ressource SHEET_DRAWING_PLUGIN, données { chartId, spec }).
 *
 * Pris en charge : histogramme, barres, courbes, aires (groupés ou empilés), secteurs,
 * nuage de points, graphiques combinés et axe secondaire, titres, légende, couleurs.
 */
import type JSZip from 'jszip';
import type { AxisSpec, ChartSpec, ChartType, DataLabelsSpec, ErrorBarsSpec, Grouping, LabelPosition, MarkerShape, SeriesSpec, TrendlineSpec } from '../../charts/spec';
import { DEFAULT_COLORS, groupingOf, isStatChart, PALETTES, parseSheetRef } from '../../charts/spec';
import { SLICER_COMPONENT_KEY, slicerFrameXml } from './slicers';
import { buildChartExXml, chartColorsXml, chartExFrameXml, chartStyleXml, CT_CHART_COLORS, CT_CHART_STYLE, CT_CHARTEX, parseChartExXml, REL_CHART_COLORS, REL_CHART_STYLE, REL_CHARTEX } from './chartex';
import type { ShapeSpec } from '../../charts/shape';
import type { ICellData, IWorkbookData, IWorksheetData } from '../../model/univer-types';
import { parseShapeAnchor, shapeXml } from './shapes';
import { addContentTypeOverride, addRelationship, attrs, escapeXml, extractAll, extractElement, insertWorksheetElement, parseXml, readRelationships, resolveRelTarget } from './ooxml';

export const DRAWING_RESOURCE = 'SHEET_DRAWING_PLUGIN';
export const CHART_COMPONENT_KEY = 'GrillonChart';
export const SHAPE_COMPONENT_KEY = 'GrillonShape';
const DRAWING_DOM_TYPE = 8;
const DRAWING_IMAGE_TYPE = 0;
const REL_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
const IMAGE_TYPES: Record<string, string> = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', webp: 'image/webp', svg: 'image/svg+xml' };
const mimeOf = (path: string) => IMAGE_TYPES[path.split('.').pop()?.toLowerCase() ?? ''] ?? 'image/png';
const EMU_PER_PX = 9525;
const REL_DRAWING = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing';
const REL_CHART = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart';
const CT_DRAWING = 'application/vnd.openxmlformats-officedocument.drawing+xml';
const CT_CHART = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
const ROW_HEADER = 46;
const COL_HEADER = 20;

export interface CellAnchor {
  column: number;
  columnOffset: number;
  row: number;
  rowOffset: number;
}

interface DrawingItem {
  unitId?: string;
  subUnitId?: string;
  drawingId: string;
  drawingType: number;
  componentKey?: string;
  sheetTransform: { from: CellAnchor; to: CellAnchor };
  axisAlignSheetTransform?: { from: CellAnchor; to: CellAnchor };
  transform: { left: number; top: number; width: number; height: number };
  data?: { chartId?: string; spec?: ChartSpec };
  imageSourceType?: string;
  source?: string;
  allowTransform?: boolean;
}

export type DrawingResource = Record<string, { data: Record<string, DrawingItem>; order: string[] }>;

/** Graphique (spec) ou image (data URL) ancré dans la feuille. */
export interface SheetChart {
  spec?: ChartSpec;
  /** Segment d'Excel (cadre sle:slicer) : nom et sorte. */
  slicer?: { name: string; kind: 'table' | 'pivot' };
  shape?: ShapeSpec;
  image?: string;
  from: CellAnchor;
  to: CellAnchor;
}

// --- Géométrie ---

function colWidth(sheet: Partial<IWorksheetData>, c: number): number {
  const cd = sheet.columnData?.[c];
  if (cd?.hd) return 0;
  return cd?.w ?? sheet.defaultColumnWidth ?? 88;
}

function rowHeight(sheet: Partial<IWorksheetData>, r: number): number {
  const rd = sheet.rowData?.[r];
  if (rd?.hd) return 0;
  return rd?.h ?? sheet.defaultRowHeight ?? 24;
}

function anchorToPx(sheet: Partial<IWorksheetData>, a: CellAnchor): { x: number; y: number } {
  let x = ROW_HEADER;
  for (let c = 0; c < a.column; c++) x += colWidth(sheet, c);
  let y = COL_HEADER;
  for (let r = 0; r < a.row; r++) y += rowHeight(sheet, r);
  return { x: x + a.columnOffset, y: y + a.rowOffset };
}

// --- Lecture ---

type Xml = Record<string, unknown>;
const arr = <T>(v: T | T[] | undefined): T[] => (v === undefined ? [] : Array.isArray(v) ? v : [v]);

function refOf(node: unknown): string | undefined {
  const n = node as { strRef?: { f?: string }; numRef?: { f?: string } } | undefined;
  return n?.strRef?.f ?? n?.numRef?.f;
}

function literalOf(node: unknown): string | undefined {
  const n = node as { v?: string } | undefined;
  return typeof n?.v === 'string' ? n.v : undefined;
}

function richText(node: unknown): string | undefined {
  // <c:title><c:tx><c:rich><a:p><a:r><a:t>…</a:t></a:r></a:p></c:rich></c:tx></c:title>
  const paras = arr((node as { tx?: { rich?: { p?: unknown } } })?.tx?.rich?.p);
  const text = paras
    .map((p) => arr((p as { r?: unknown }).r).map((r) => String((r as { t?: unknown }).t ?? '')).join(''))
    .join('\n');
  return text || undefined;
}

function colorOf(spPr: unknown): string | undefined {
  const fill = (spPr as { solidFill?: { srgbClr?: { val?: string } } })?.solidFill?.srgbClr?.val;
  const line = (spPr as { ln?: { solidFill?: { srgbClr?: { val?: string } } } })?.ln?.solidFill?.srgbClr?.val;
  const v = fill ?? line;
  return v ? `#${v.toUpperCase()}` : undefined;
}

const GROUP_TAGS: Record<string, ChartType> = {
  barChart: 'column',
  bar3DChart: 'column',
  lineChart: 'line',
  line3DChart: 'line',
  areaChart: 'area',
  area3DChart: 'area',
  pieChart: 'pie',
  pie3DChart: 'pie',
  doughnutChart: 'doughnut',
  scatterChart: 'scatter',
  bubbleChart: 'bubble',
  radarChart: 'radar',
};

const val = (node: unknown): string | undefined => {
  const v = (node as { val?: unknown } | undefined)?.val;
  return v === undefined ? undefined : String(v);
};
const flag = (node: unknown): boolean => {
  const v = val(node);
  return v === '1' || v === 'true';
};
const numVal = (node: unknown): number | undefined => {
  const v = val(node);
  return v !== undefined && Number.isFinite(Number(v)) ? Number(v) : undefined;
};

const LABEL_POSITIONS: Record<string, LabelPosition> = { ctr: 'center', inEnd: 'insideEnd', inBase: 'insideBase', outEnd: 'outsideEnd', t: 'above', b: 'below', l: 'left', r: 'right' };

function parseDataLabels(node: unknown): DataLabelsSpec | undefined {
  const d = node as Xml | undefined;
  if (!d || flag(d.delete)) return undefined;
  const value = flag(d.showVal);
  const category = flag(d.showCatName);
  const series = flag(d.showSerName);
  const percent = flag(d.showPercent);
  if (!value && !category && !series && !percent) return undefined;
  const out: DataLabelsSpec = { show: true };
  if (value) out.value = true;
  if (category) out.category = true;
  if (series) out.series = true;
  if (percent) out.percent = true;
  const pos = LABEL_POSITIONS[val(d.dLblPos) ?? ''];
  if (pos) out.position = pos;
  const fmt = (d.numFmt as { formatCode?: string; sourceLinked?: unknown } | undefined);
  if (fmt?.formatCode && fmt.formatCode !== 'General' && String(fmt.sourceLinked) !== '1') out.numFmt = fmt.formatCode;
  if (typeof d.separator === 'string' && d.separator !== '; ') out.separator = d.separator;
  return out;
}

const TREND_TYPES: Record<string, TrendlineSpec['type']> = { linear: 'linear', exp: 'exponential', log: 'logarithmic', poly: 'polynomial', power: 'power', movingAvg: 'movingAverage' };

function parseTrendline(node: Xml): TrendlineSpec | null {
  const type = TREND_TYPES[val(node.trendlineType) ?? ''];
  if (!type) return null;
  const t: TrendlineSpec = { type };
  const order = numVal(node.order);
  const period = numVal(node.period);
  const forward = numVal(node.forward);
  const backward = numVal(node.backward);
  if (order !== undefined) t.order = order;
  if (period !== undefined) t.period = period;
  if (forward) t.forward = forward;
  if (backward) t.backward = backward;
  if (flag(node.dispEq)) t.showEquation = true;
  if (flag(node.dispRSqr)) t.showR2 = true;
  if (typeof node.name === 'string' && node.name) t.name = node.name;
  return t;
}

function parseErrorBars(node: Xml): ErrorBarsSpec | undefined {
  if (val(node.errDir) === 'x') return undefined;
  const kind = ({ fixedVal: 'fixed', percentage: 'percent', stdDev: 'stddev' } as const)[val(node.errValType) as 'fixedVal'];
  if (!kind) return undefined;
  const out: ErrorBarsSpec = { kind };
  const value = numVal(node.val);
  if (value !== undefined) out.value = value;
  const sides = val(node.errBarType);
  if (sides === 'plus' || sides === 'minus') out.sides = sides;
  if (flag(node.noEndCap)) out.noCap = true;
  return out;
}

const SYMBOLS: Record<string, MarkerShape> = { circle: 'circle', square: 'square', diamond: 'diamond', triangle: 'triangle', none: 'none' };

function parseSeries(ser: Xml, type: ChartType, palette: string[], idx: number, group: Xml): { spec: SeriesSpec; hasMarker: boolean } {
  const tx = ser.tx as Xml | undefined;
  const xy = type === 'scatter' || type === 'bubble';
  const s: SeriesSpec = {
    valuesRef: refOf(xy ? ser.yVal : ser.val) ?? '',
    categoriesRef: refOf(xy ? ser.xVal : ser.cat),
  };
  if (!s.categoriesRef) delete s.categoriesRef;
  const nameRef = refOf(tx);
  if (nameRef) s.nameRef = nameRef;
  else if (literalOf(tx)) s.name = literalOf(tx);
  if (type === 'bubble') {
    const sizes = refOf(ser.bubbleSize);
    if (sizes) s.sizesRef = sizes;
  }
  const spPr = ser.spPr as Xml | undefined;
  const color = colorOf(spPr) ?? colorOf((ser.marker as Xml | undefined)?.spPr);
  if (color) s.color = color;
  const ln = spPr?.ln as Xml | undefined;
  const w = Number((ln as { w?: unknown } | undefined)?.w);
  if (Number.isFinite(w) && w > 0 && ln?.noFill === undefined && (type === 'line' || type === 'scatter' || type === 'radar')) {
    const pt = Math.round((w / 12700) * 100) / 100;
    if (pt !== 2.25) s.lineWidth = pt;
  }
  const dash = val(ln?.prstDash);
  if (dash === 'dash' || dash === 'lgDash' || dash === 'sysDash') s.dash = 'dash';
  else if (dash === 'sysDot' || dash === 'dot') s.dash = 'dot';
  const marker = ser.marker as Xml | undefined;
  let hasMarker = false;
  if (marker) {
    const shape = SYMBOLS[val(marker.symbol) ?? ''];
    const size = numVal(marker.size);
    if (shape && shape !== 'none') hasMarker = true;
    const m: { shape?: MarkerShape; size?: number } = {};
    if (shape && shape !== 'circle') m.shape = shape;
    if (size !== undefined && size !== 5 && size !== 6 && shape !== 'none') m.size = size;
    if (Object.keys(m).length) s.marker = m;
  }
  if (flag(ser.smooth)) s.smooth = true;
  if (type === 'area' || type === 'radar') {
    const alpha = numVal(((spPr?.solidFill as Xml | undefined)?.srgbClr as Xml | undefined)?.alpha);
    if (alpha !== undefined && alpha < 100000) s.opacity = alpha / 100000;
  }
  const dl = parseDataLabels(ser.dLbls) ?? undefined;
  if (dl) s.dataLabels = dl;
  const trendlines = arr(ser.trendline as Xml | Xml[]).map(parseTrendline).filter((x): x is TrendlineSpec => !!x);
  if (trendlines.length) s.trendlines = trendlines;
  const eb = arr(ser.errBars as Xml | Xml[]).map(parseErrorBars).find((x) => !!x);
  if (eb) s.errorBars = eb;
  const points: Record<number, string> = {};
  for (const dPt of arr(ser.dPt as Xml | Xml[])) {
    const index = numVal(dPt.idx);
    const c = colorOf(dPt.spPr);
    if (index !== undefined && c && c !== palette[index % palette.length]!.toUpperCase()) points[index] = c;
  }
  if (Object.keys(points).length) s.pointColors = points;
  void idx;
  void group;
  return { spec: s, hasMarker };
}

function parseAxis(ax: Xml | undefined, kind: 'value' | 'category', defaultGrid = kind === 'value'): AxisSpec | undefined {
  if (!ax) return undefined;
  const a: AxisSpec = {};
  const scaling = ax.scaling as Xml | undefined;
  if (kind === 'value') {
    const min = numVal(scaling?.min);
    const max = numVal(scaling?.max);
    if (min !== undefined) a.min = min;
    if (max !== undefined) a.max = max;
    const major = numVal(ax.majorUnit);
    if (major !== undefined) a.majorUnit = major;
    if (scaling?.logBase !== undefined) a.log = true;
  }
  if (val(scaling?.orientation) === 'maxMin') a.reverse = true;
  const fmt = ax.numFmt as { formatCode?: string; sourceLinked?: unknown } | undefined;
  if (fmt?.formatCode && fmt.formatCode !== 'General' && String(fmt.sourceLinked) !== '1') a.numFmt = fmt.formatCode;
  if (val(ax.tickLblPos) === 'none') a.labels = 'none';
  const tick = val(ax.majorTickMark);
  if (tick === 'none' || tick === 'in' || tick === 'cross') a.ticks = tick;
  const minor = val(ax.minorTickMark);
  if (minor && minor !== 'none') a.minorTicks = true;
  const hasMajor = ax.majorGridlines !== undefined;
  if (hasMajor !== defaultGrid) a.gridMajor = hasMajor;
  if (ax.minorGridlines !== undefined) a.gridMinor = true;
  if (flag(ax.delete)) a.hidden = true;
  const rot = Number(((ax.txPr as Xml | undefined)?.bodyPr as { rot?: unknown } | undefined)?.rot);
  if (Number.isFinite(rot) && rot !== 0) a.labelAngle = Math.round(rot / 60000);
  return Object.keys(a).length ? a : undefined;
}

/** Convertit une partie xl/charts/chartN.xml en ChartSpec (null si rien d'exploitable). */
export function parseChartXml(xml: string): ChartSpec | null {
  const doc = parseXml<{ chartSpace?: Xml }>(xml);
  const space = doc.chartSpace;
  const chart = space?.chart as Xml | undefined;
  const plot = chart?.plotArea as Xml | undefined;
  if (!plot) return null;
  // Axes : on repère l'axe de valeurs secondaire (celui qui croise au maximum / à droite).
  const valAxes = arr(plot.valAx as Xml | Xml[]);
  const secondaryAxIds = new Set(
    valAxes.filter((ax) => (ax.crosses as Xml | undefined)?.val === 'max' || ['r', 't'].includes(String((ax.axPos as Xml | undefined)?.val))).map((ax) => String((ax.axId as Xml).val)),
  );
  const series: SeriesSpec[] = [];
  let mainType: ChartType | null = null;
  let mainGroup: Xml | null = null;
  let grouping: Grouping = 'standard';
  let anyMarker = false;
  const palette = PALETTES.office!.colors;
  for (const [tag, baseType] of Object.entries(GROUP_TAGS)) {
    for (const group of arr(plot[tag] as Xml | Xml[])) {
      let type = baseType;
      if (tag.startsWith('bar') && (group.barDir as Xml | undefined)?.val === 'bar') type = 'bar';
      const g = String((group.grouping as Xml | undefined)?.val ?? '');
      const axIds = arr(group.axId as Xml | Xml[]).map((a) => String(a.val));
      const secondary = axIds.some((id) => secondaryAxIds.has(id)) && type !== 'scatter' && type !== 'bubble';
      if (!secondary && mainType === null) {
        mainType = type;
        mainGroup = group;
        if (g === 'stacked') grouping = 'stacked';
        else if (g === 'percentStacked') grouping = 'percent';
      }
      arr(group.ser as Xml | Xml[]).forEach((ser, i) => {
        const { spec: s, hasMarker } = parseSeries(ser, type, palette, i, group);
        s.type = type;
        if (secondary) s.secondaryAxis = true;
        if (hasMarker) anyMarker = true;
        if (s.valuesRef) series.push(s);
      });
      // Étiquettes définies au niveau du groupe : reprises par les séries qui n'ont pas les leurs.
      const groupLabels = parseDataLabels(group.dLbls);
      if (groupLabels) for (const s of series.slice(-arr(group.ser as Xml | Xml[]).length)) s.dataLabels ??= groupLabels;
    }
  }
  if (!series.length) return null;
  const type = mainType ?? series[0]!.type ?? 'column';
  // Le type n'est gardé par série que s'il diffère du type principal (graphique combiné).
  for (const s of series) if (s.type === type) delete s.type;
  const spec: ChartSpec = { version: 1, type, series };
  const title = richText(chart?.title);
  if (title) spec.title = title;
  if (flag((chart?.title as Xml | undefined)?.overlay)) spec.titleOverlay = true;
  if (grouping !== 'standard') spec.grouping = grouping;
  const legend = chart?.legend as Xml | undefined;
  const legendPos = val(legend?.legendPos);
  spec.legend = legend === undefined ? 'none' : legendPos === 't' ? 'top' : legendPos === 'r' ? 'right' : legendPos === 'l' ? 'left' : 'bottom';
  if (legend && flag(legend.overlay)) spec.legendOverlay = true;

  const axisTitle = (ax: Xml | undefined) => (ax ? richText(ax.title) : undefined);
  const xy = type === 'scatter' || type === 'bubble';
  const catAx = arr(plot.catAx as Xml | Xml[])[0] ?? (xy ? valAxes[0] : undefined);
  const primaryVal = valAxes.find((ax) => !secondaryAxIds.has(String((ax.axId as Xml).val)) && ax !== catAx);
  const secondaryVal = valAxes.find((ax) => secondaryAxIds.has(String((ax.axId as Xml).val)));
  const x = axisTitle(catAx);
  const y = axisTitle(primaryVal);
  const y2 = axisTitle(secondaryVal);
  if (x) spec.xAxisTitle = x;
  if (y) spec.yAxisTitle = y;
  if (y2) spec.y2AxisTitle = y2;
  const axes: ChartSpec['axes'] = {};
  const ax = parseAxis(catAx, xy ? 'value' : 'category');
  const ay = parseAxis(primaryVal, 'value');
  const ay2 = parseAxis(secondaryVal, 'value', false);
  if (ax) axes.x = ax;
  if (ay) axes.y = ay;
  if (ay2) axes.y2 = ay2;
  if (Object.keys(axes).length) spec.axes = axes;

  // Options propres au type.
  const g = mainGroup;
  if (g) {
    const gap = numVal(g.gapWidth);
    if (gap !== undefined && gap !== 150) spec.gapWidth = gap;
    const overlap = numVal(g.overlap);
    if (overlap !== undefined && overlap !== (grouping === 'standard' ? 0 : 100)) spec.overlap = overlap;
    const hole = numVal(g.holeSize);
    if (hole !== undefined && hole !== 50) spec.holeSize = hole;
    const first = numVal(g.firstSliceAng);
    if (first) spec.firstSliceAngle = first;
    const bubble = numVal(g.bubbleScale);
    if (bubble !== undefined && bubble !== 100) spec.bubbleScale = bubble;
    if (val(g.radarStyle) === 'filled') spec.radarFilled = true;
    if (flag(g.varyColors) && series.length === 1 && type !== 'pie' && type !== 'doughnut') spec.varyColors = true;
    if (type === 'scatter') {
      const style = val(g.scatterStyle);
      const smooth = series.some((s) => s.smooth);
      const lines = arr(g.ser as Xml | Xml[]).some((s) => ((s.spPr as Xml | undefined)?.ln as Xml | undefined)?.noFill === undefined);
      if (lines) spec.scatterStyle = smooth || style === 'smoothMarker' ? (anyMarker ? 'smoothMarkers' : 'smooth') : anyMarker ? 'straightMarkers' : 'straight';
      for (const s of series) delete s.smooth;
    }
  }
  if ((type === 'line' || type === 'radar') && anyMarker) spec.markers = true;
  for (const s of series) {
    if (s.marker?.shape === 'none' && !spec.markers) {
      delete s.marker.shape;
      if (!Object.keys(s.marker).length) delete s.marker;
    }
  }
  if (type === 'line' && series.length && series.every((s) => s.smooth)) {
    spec.smooth = true;
    for (const s of series) delete s.smooth;
  }

  // Étiquettes identiques sur toutes les séries : réglage du graphique.
  const first = series[0]?.dataLabels;
  if (first && series.every((s) => JSON.stringify(s.dataLabels) === JSON.stringify(first))) {
    spec.dataLabels = first;
    for (const s of series) delete s.dataLabels;
  }
  if (plot.dTable !== undefined) spec.dataTable = true;
  const blanks = val(chart?.dispBlanksAs);
  if (blanks === 'zero') spec.emptyCells = 'zero';
  else if (blanks === 'span') spec.emptyCells = 'connect';
  if (val(chart?.plotVisOnly) === '0') spec.plotVisibleOnly = false;

  // Mise en forme de la zone de graphique.
  const sp = space?.spPr as Xml | undefined;
  const bg = (sp?.solidFill as { srgbClr?: { val?: string } } | undefined)?.srgbClr?.val;
  if (bg && bg.toUpperCase() !== 'FFFFFF') spec.background = `#${bg.toUpperCase()}`;
  if ((sp?.ln as Xml | undefined)?.noFill !== undefined) spec.border = false;
  const sz = Number((((space?.txPr as Xml | undefined)?.p as Xml | undefined)?.pPr as Xml | undefined)?.defRPr && ((((space?.txPr as Xml).p as Xml).pPr as Xml).defRPr as { sz?: unknown }).sz);
  if (Number.isFinite(sz) && sz > 0 && sz !== 1200) spec.fontSize = Math.round(sz / 100);
  return spec;
}

/** Lit les graphiques d'une feuille (via ses relations de dessin). */
export async function readSheetCharts(zip: JSZip, sheetPath: string): Promise<SheetChart[]> {
  const out: SheetChart[] = [];
  for (const rel of (await readRelationships(zip, sheetPath)).filter((r) => r.Type === REL_DRAWING)) {
    const drawingPath = resolveRelTarget(sheetPath, rel.Target);
    const file = zip.file(drawingPath);
    if (!file) continue;
    const xml = await file.async('string');
    const drawingRels = await readRelationships(zip, drawingPath);
    const relById = new Map(drawingRels.map((r) => [r.Id, r]));
    for (const anchor of [...extractAll(xml, 'twoCellAnchor'), ...extractAll(xml, 'oneCellAnchor')]) {
      let spec: ChartSpec | null = null;
      let image: string | undefined;
      let shape: ShapeSpec | null = null;
      const chartRef = /<c:chart\b[^>]*\br:id="([^"]+)"/.exec(anchor.xml)?.[1];
      const chartExRef = /<cx:chart\b[^>]*\br:id="([^"]+)"/.exec(anchor.xml)?.[1];
      const blipRef = /<a:blip\b[^>]*\br:embed="([^"]+)"/.exec(anchor.xml)?.[1];
      if (chartExRef) {
        const rel = relById.get(chartExRef);
        const file = rel && zip.file(resolveRelTarget(drawingPath, rel.Target));
        if (file) spec = parseChartExXml(await file.async('string'));
      } else if (chartRef) {
        const chartRel = relById.get(chartRef);
        const chartFile = chartRel && zip.file(resolveRelTarget(drawingPath, chartRel.Target));
        if (chartFile) spec = parseChartXml(await chartFile.async('string'));
      } else if (blipRef) {
        const rel = relById.get(blipRef);
        const target = rel ? resolveRelTarget(drawingPath, rel.Target) : undefined;
        const media = target ? zip.file(target) : null;
        if (target && media) image = `data:${mimeOf(target)};base64,${await media.async('base64')}`;
      }
      if (!spec && !image) shape = parseShapeAnchor(anchor.xml);
      if (!spec && !image && !shape) continue;
      const pos = (tag: string): CellAnchor | null => {
        const el = extractElement(anchor.xml, tag);
        if (!el) return null;
        const n = (name: string) => Number(new RegExp(`<xdr:${name}>(-?\\d+)</xdr:${name}>`).exec(el.xml)?.[1] ?? 0);
        return { column: n('col'), columnOffset: Math.round(n('colOff') / EMU_PER_PX), row: n('row'), rowOffset: Math.round(n('rowOff') / EMU_PER_PX) };
      };
      const from = pos('from');
      let to = pos('to');
      if (!from) continue;
      if (!to) {
        // oneCellAnchor : taille donnée par <xdr:ext cx cy>.
        const ext = /<xdr:ext\b[^>]*cx="(\d+)"[^>]*cy="(\d+)"/.exec(anchor.xml);
        const w = ext ? Number(ext[1]) / EMU_PER_PX : 480;
        const h = ext ? Number(ext[2]) / EMU_PER_PX : 288;
        to = { column: from.column + Math.max(1, Math.round(w / 64)), columnOffset: 0, row: from.row + Math.max(1, Math.round(h / 20)), rowOffset: 0 };
      }
      out.push(spec ? { spec, from, to } : shape ? { shape, from, to } : { image, from, to });
    }
  }
  return out;
}

/** Graphiques lus → entrées de la ressource de dessins Univer. */
export function chartsToDrawingResource(unitId: string, sheetId: string, sheet: Partial<IWorksheetData>, charts: SheetChart[]): DrawingResource[string] {
  const data: Record<string, DrawingItem> = {};
  const order: string[] = [];
  charts.forEach((c, i) => {
    const id = `chart${Date.now().toString(36)}${i}`;
    const a = anchorToPx(sheet, c.from);
    const b = anchorToPx(sheet, c.to);
    if (c.shape) {
      data[id] = {
        unitId,
        subUnitId: sheetId,
        drawingId: id,
        drawingType: DRAWING_DOM_TYPE,
        componentKey: SHAPE_COMPONENT_KEY,
        sheetTransform: { from: c.from, to: c.to },
        axisAlignSheetTransform: { from: c.from, to: c.to },
        transform: { left: a.x, top: a.y, width: Math.max(8, b.x - a.x), height: Math.max(8, b.y - a.y) },
        data: { shapeId: `shape-${id}`, spec: c.shape } as never,
        allowTransform: true,
      };
      order.push(id);
      return;
    }
    if (c.image) {
      data[id] = {
        unitId,
        subUnitId: sheetId,
        drawingId: id,
        drawingType: DRAWING_IMAGE_TYPE,
        imageSourceType: 'URL',
        source: c.image,
        sheetTransform: { from: c.from, to: c.to },
        axisAlignSheetTransform: { from: c.from, to: c.to },
        transform: { left: a.x, top: a.y, width: Math.max(1, b.x - a.x), height: Math.max(1, b.y - a.y) },
      } as DrawingItem;
      order.push(id);
      return;
    }
    data[id] = {
      unitId,
      subUnitId: sheetId,
      drawingId: id,
      drawingType: DRAWING_DOM_TYPE,
      componentKey: CHART_COMPONENT_KEY,
      sheetTransform: { from: c.from, to: c.to },
      axisAlignSheetTransform: { from: c.from, to: c.to },
      transform: { left: a.x, top: a.y, width: Math.max(40, b.x - a.x), height: Math.max(40, b.y - a.y) },
      data: { chartId: `chart-${id}`, spec: c.spec },
      allowTransform: true,
    };
    order.push(id);
  });
  return { data, order };
}

/** Segment ou chronologie lu → entrée (composant flottant de Grillon) de la ressource de dessins. */
export function addSlicerToResource(
  resource: DrawingResource[string] | undefined,
  unitId: string,
  sheetId: string,
  sheet: Partial<IWorksheetData>,
  slicer: { spec: unknown; from: CellAnchor; to: CellAnchor },
): DrawingResource[string] {
  const out = resource ?? { data: {}, order: [] };
  const id = `slicer${Date.now().toString(36)}${out.order.length}`;
  const a = anchorToPx(sheet, slicer.from);
  const b = anchorToPx(sheet, slicer.to);
  out.data[id] = {
    unitId,
    subUnitId: sheetId,
    drawingId: id,
    drawingType: DRAWING_DOM_TYPE,
    componentKey: SLICER_COMPONENT_KEY,
    sheetTransform: { from: slicer.from, to: slicer.to },
    axisAlignSheetTransform: { from: slicer.from, to: slicer.to },
    transform: { left: a.x, top: a.y, width: Math.max(40, b.x - a.x), height: Math.max(40, b.y - a.y) },
    data: slicer.spec as never,
    allowTransform: true,
  };
  out.order.push(id);
  return out;
}

// --- Écriture ---

/** Valeurs d'une référence, pour les caches du graphique (Excel les affiche avant recalcul). */
function readValues(wb: IWorkbookData, ref: string | undefined, defaultSheet: string): (string | number | null)[] {
  if (!ref) return [];
  const p = parseSheetRef(ref);
  if (!p) return [];
  const sheet = Object.values(wb.sheets).find((s) => s.name === (p.sheet ?? defaultSheet));
  const out: (string | number | null)[] = [];
  for (let r = p.startRow; r <= p.endRow; r++) {
    for (let c = p.startColumn; c <= p.endColumn; c++) {
      const cell = sheet?.cellData?.[r]?.[c] as ICellData | undefined;
      const v = cell?.v;
      out.push(v === undefined || v === null || v === '' ? null : typeof v === 'number' ? v : cell?.t === 2 && Number.isFinite(Number(v)) ? Number(v) : String(v));
    }
  }
  return out;
}

function numCache(values: (string | number | null)[]): string {
  const pts = values.map((v, i) => (typeof v === 'number' ? `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>` : '')).join('');
  return `<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>${pts}</c:numCache>`;
}

function strCache(values: (string | number | null)[]): string {
  const pts = values.map((v, i) => (v === null ? '' : `<c:pt idx="${i}"><c:v>${escapeXml(String(v))}</c:v></c:pt>`)).join('');
  return `<c:strCache><c:ptCount val="${values.length}"/>${pts}</c:strCache>`;
}

function fRef(ref: string): string {
  return escapeXml(ref.replace(/^=/, ''));
}

function richTitle(text: string, size = 1400, overlay = false): string {
  return `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="${size}" b="1"/></a:pPr><a:r><a:rPr lang="fr-FR" sz="${size}" b="1"/><a:t>${escapeXml(text)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="${overlay ? 1 : 0}"/></c:title>`;
}

function solidFill(color: string, alpha?: number): string {
  const hex = color.replace('#', '').toUpperCase();
  return alpha !== undefined && alpha < 1 ? `<a:solidFill><a:srgbClr val="${hex}"><a:alpha val="${Math.round(alpha * 100000)}"/></a:srgbClr></a:solidFill>` : `<a:solidFill><a:srgbClr val="${hex}"/></a:solidFill>`;
}

const DLBL_POS: Record<string, string> = { center: 'ctr', insideEnd: 'inEnd', insideBase: 'inBase', outsideEnd: 'outEnd', above: 't', below: 'b', left: 'l', right: 'r' };

/** Positions d'étiquettes acceptées par Excel selon le type de graphique (sinon le fichier est jugé corrompu). */
function allowedLabelPositions(type: ChartType, grouping: Grouping): string[] {
  switch (type) {
    case 'column':
    case 'bar':
      return grouping === 'standard' ? ['ctr', 'inEnd', 'inBase', 'outEnd'] : ['ctr', 'inEnd', 'inBase'];
    case 'line':
    case 'scatter':
    case 'bubble':
      return ['ctr', 't', 'b', 'l', 'r'];
    case 'pie':
      return ['ctr', 'inEnd', 'outEnd'];
    default:
      return [];
  }
}

function dLblsXml(dl: DataLabelsSpec | undefined, type: ChartType, grouping: Grouping): string {
  if (!dl?.show) return '';
  const value = dl.value ?? (!dl.category && !dl.series && !dl.percent);
  const pos = DLBL_POS[dl.position ?? ''];
  const posXml = pos && allowedLabelPositions(type, grouping).includes(pos) ? `<c:dLblPos val="${pos}"/>` : '';
  const fmt = dl.numFmt ? `<c:numFmt formatCode="${escapeXml(dl.numFmt)}" sourceLinked="0"/>` : '';
  const sep = dl.separator && dl.separator !== '; ' ? `<c:separator>${escapeXml(dl.separator)}</c:separator>` : '';
  const b = (v: boolean | undefined) => (v ? 1 : 0);
  return `<c:dLbls>${fmt}<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>${posXml}<c:showLegendKey val="0"/><c:showVal val="${b(value)}"/><c:showCatName val="${b(dl.category)}"/><c:showSerName val="${b(dl.series)}"/><c:showPercent val="${b(dl.percent && (type === 'pie' || type === 'doughnut'))}"/><c:showBubbleSize val="0"/>${sep}</c:dLbls>`;
}

const TREND_XML: Record<TrendlineSpec['type'], string> = { linear: 'linear', exponential: 'exp', logarithmic: 'log', polynomial: 'poly', power: 'power', movingAverage: 'movingAvg' };

function trendlineXml(t: TrendlineSpec, color: string): string {
  const order = t.type === 'polynomial' ? `<c:order val="${Math.min(6, Math.max(2, Math.round(t.order ?? 2)))}"/>` : '';
  const period = t.type === 'movingAverage' ? `<c:period val="${Math.max(2, Math.round(t.period ?? 2))}"/>` : '';
  const fwd = t.forward && t.type !== 'movingAverage' ? `<c:forward val="${t.forward}"/>` : '';
  const bwd = t.backward && t.type !== 'movingAverage' ? `<c:backward val="${t.backward}"/>` : '';
  const name = t.name ? `<c:name>${escapeXml(t.name)}</c:name>` : '';
  return `<c:trendline>${name}<c:spPr><a:ln w="19050" cap="rnd">${solidFill(t.color ?? color)}<a:prstDash val="sysDash"/></a:ln></c:spPr><c:trendlineType val="${TREND_XML[t.type]}"/>${order}${period}${fwd}${bwd}<c:dispRSqr val="${t.showR2 ? 1 : 0}"/><c:dispEq val="${t.showEquation ? 1 : 0}"/></c:trendline>`;
}

function errBarsXml(e: ErrorBarsSpec, color: string): string {
  const kind = { fixed: 'fixedVal', percent: 'percentage', stddev: 'stdDev' }[e.kind];
  const value = e.value ?? (e.kind === 'percent' ? 5 : 1);
  return `<c:errBars><c:errDir val="y"/><c:errBarType val="${e.sides ?? 'both'}"/><c:errValType val="${kind}"/><c:noEndCap val="${e.noCap ? 1 : 0}"/><c:val val="${value}"/><c:spPr><a:ln w="12700">${solidFill(color)}</a:ln></c:spPr></c:errBars>`;
}

const MARKER_XML: Record<MarkerShape, string> = { circle: 'circle', square: 'square', diamond: 'diamond', triangle: 'triangle', none: 'none' };

/** ChartSpec → XML de partie graphique. */
export function buildChartXml(wb: IWorkbookData, spec: ChartSpec, sheetName: string): string {
  const palette = PALETTES[spec.palette ?? 'office']?.colors ?? DEFAULT_COLORS;
  const grouping = groupingOf(spec);
  const xy = spec.type === 'scatter' || spec.type === 'bubble';
  const single = spec.type === 'pie' || spec.type === 'doughnut' || spec.type === 'radar' || xy;
  const groups = new Map<string, { type: ChartType; secondary: boolean; series: { s: SeriesSpec; idx: number }[] }>();
  spec.series.forEach((s, idx) => {
    const type = single ? spec.type : (s.type ?? spec.type);
    const secondary = !!s.secondaryAxis && !single;
    const key = `${type === 'bar' ? 'column' : type}|${secondary}`;
    if (!groups.has(key)) groups.set(key, { type, secondary, series: [] });
    groups.get(key)!.series.push({ s, idx });
  });
  const hasSecondary = [...groups.values()].some((g) => g.secondary);
  const horizontal = spec.type === 'bar';
  const AX = { cat: 500000001, val: 500000002, cat2: 500000003, val2: 500000004 };

  const serXml = (s: SeriesSpec, idx: number, type: ChartType) => {
    const color = s.color ?? palette[idx % palette.length]!;
    const dlSpec = s.dataLabels ?? spec.dataLabels;
    const tx = s.nameRef
      ? `<c:tx><c:strRef><c:f>${fRef(s.nameRef)}</c:f>${strCache(readValues(wb, s.nameRef, sheetName))}</c:strRef></c:tx>`
      : `<c:tx><c:v>${escapeXml(s.name ?? `Série ${idx + 1}`)}</c:v></c:tx>`;
    const lineW = Math.round((s.lineWidth ?? 2.25) * 12700);
    const dash = s.dash === 'dash' ? '<a:prstDash val="dash"/>' : s.dash === 'dot' ? '<a:prstDash val="sysDot"/>' : '';
    const scatterLines = type === 'scatter' && (spec.scatterStyle ?? 'markers') !== 'markers';
    const lineType = type === 'line' || type === 'radar' || scatterLines;
    const opacity = type === 'area' || (type === 'radar' && spec.radarFilled) ? s.opacity : undefined;
    let spPr: string;
    if (type === 'pie' || type === 'doughnut') spPr = '';
    else if (type === 'scatter' && !scatterLines) spPr = '<c:spPr><a:ln w="19050"><a:noFill/></a:ln></c:spPr>';
    else if (lineType) spPr = `<c:spPr><a:ln w="${lineW}" cap="rnd">${solidFill(color)}${dash}<a:round/></a:ln></c:spPr>`;
    else spPr = `<c:spPr>${solidFill(color, opacity)}</c:spPr>`;
    const cats = readValues(wb, s.categoriesRef, sheetName);
    const catNumeric = cats.length > 0 && cats.every((v) => typeof v === 'number' || v === null);
    const cat = s.categoriesRef
      ? catNumeric
        ? `<c:numRef><c:f>${fRef(s.categoriesRef)}</c:f>${numCache(cats)}</c:numRef>`
        : `<c:strRef><c:f>${fRef(s.categoriesRef)}</c:f>${strCache(cats)}</c:strRef>`
      : '';
    const val = `<c:numRef><c:f>${fRef(s.valuesRef)}</c:f>${numCache(readValues(wb, s.valuesRef, sheetName))}</c:numRef>`;
    const shape = s.marker?.shape ?? (type === 'scatter' ? (spec.scatterStyle === 'straight' || spec.scatterStyle === 'smooth' ? 'none' : 'circle') : spec.markers ? 'circle' : 'none');
    const markerSize = Math.min(72, Math.max(2, Math.round(s.marker?.size ?? (type === 'scatter' ? 6 : 5))));
    const marker =
      type === 'line' || type === 'scatter' || type === 'radar'
        ? `<c:marker><c:symbol val="${MARKER_XML[shape]}"/>${shape === 'none' ? '' : `<c:size val="${markerSize}"/><c:spPr>${solidFill(color)}</c:spPr>`}</c:marker>`
        : '';
    const point = (i: number, c: string) => `<c:dPt><c:idx val="${i}"/>${type === 'column' || type === 'bar' ? '<c:invertIfNegative val="0"/>' : ''}<c:bubble3D val="0"/><c:spPr>${solidFill(c)}</c:spPr></c:dPt>`;
    const paletteByIndex = (): Record<string, string> => Object.fromEntries(cats.map((_, i) => [String(i), palette[i % palette.length]!]));
    const pointColors: Record<string, string> = spec.varyColors && spec.series.length === 1 && type !== "pie" && type !== "doughnut" ? paletteByIndex() : {};
    const colors: Record<string, string> = { ...(type === 'pie' || type === 'doughnut' ? paletteByIndex() : {}), ...pointColors, ...(s.pointColors ?? {}) };
    const dPts = Object.entries(colors).map(([i, c]) => point(Number(i), c)).join('');
    const dLbls = dLblsXml(dlSpec, type, grouping);
    const trend = (s.trendlines ?? []).filter(() => type !== 'pie' && type !== 'doughnut' && type !== 'radar').map((t) => trendlineXml(t, color)).join('');
    const err = s.errorBars && type !== 'pie' && type !== 'doughnut' && type !== 'radar' ? errBarsXml(s.errorBars, color) : '';
    const head = `<c:ser><c:idx val="${idx}"/><c:order val="${idx}"/>${tx}${spPr}`;
    const explosion = (type === 'pie' || type === 'doughnut') && spec.explosion ? `<c:explosion val="${Math.round(spec.explosion)}"/>` : '';
    const smooth = (spec.smooth || s.smooth) && (type === 'line') ? '<c:smooth val="1"/>' : type === 'line' ? '<c:smooth val="0"/>' : '';
    switch (type) {
      case 'scatter': {
        const smoothed = (spec.scatterStyle ?? 'markers').startsWith('smooth');
        return `${head}${marker}${dPts}${dLbls}${trend}${err}${cat ? `<c:xVal>${cat}</c:xVal>` : ''}<c:yVal>${val}</c:yVal><c:smooth val="${smoothed ? 1 : 0}"/></c:ser>`;
      }
      case 'bubble': {
        const sizes = `<c:numRef><c:f>${fRef(s.sizesRef ?? s.valuesRef)}</c:f>${numCache(readValues(wb, s.sizesRef ?? s.valuesRef, sheetName))}</c:numRef>`;
        return `${head}<c:invertIfNegative val="0"/>${dPts}${dLbls}${trend}${err}${cat ? `<c:xVal>${cat}</c:xVal>` : ''}<c:yVal>${val}</c:yVal><c:bubbleSize>${sizes}</c:bubbleSize><c:bubble3D val="0"/></c:ser>`;
      }
      case 'column':
      case 'bar':
        return `${head}<c:invertIfNegative val="0"/>${dPts}${dLbls}${trend}${err}${cat ? `<c:cat>${cat}</c:cat>` : ''}<c:val>${val}</c:val></c:ser>`;
      case 'line':
        return `${head}${marker}${dPts}${dLbls}${trend}${err}${cat ? `<c:cat>${cat}</c:cat>` : ''}<c:val>${val}</c:val>${smooth}</c:ser>`;
      case 'area':
        return `${head}${dPts}${dLbls}${trend}${err}${cat ? `<c:cat>${cat}</c:cat>` : ''}<c:val>${val}</c:val></c:ser>`;
      case 'radar':
        return `${head}${marker}${dPts}${dLbls}${cat ? `<c:cat>${cat}</c:cat>` : ''}<c:val>${val}</c:val></c:ser>`;
      default:
        return `${head}${explosion}${dPts}${dLbls}${cat ? `<c:cat>${cat}</c:cat>` : ''}<c:val>${val}</c:val></c:ser>`;
    }
  };

  const groupingXml = (g: Grouping) => (g === 'percent' ? 'percentStacked' : g);
  const vary = spec.varyColors && spec.series.length === 1 ? 1 : 0;
  const groupXml = [...groups.values()]
    .map((g) => {
      const series = g.series.map(({ s, idx }) => serXml(s, idx, g.type)).join('');
      const axIds = g.secondary ? `<c:axId val="${AX.cat2}"/><c:axId val="${AX.val2}"/>` : `<c:axId val="${AX.cat}"/><c:axId val="${AX.val}"/>`;
      // L'empilement ne s'applique qu'aux séries du type principal ; les autres restent « standard ».
      const stack = g.type === spec.type || (g.type === 'column' && spec.type === 'bar') ? groupingXml(grouping) : 'standard';
      switch (g.type) {
        case 'column':
        case 'bar': {
          const overlap = spec.overlap ?? (stack !== 'standard' ? 100 : undefined);
          return `<c:barChart><c:barDir val="${horizontal ? 'bar' : 'col'}"/><c:grouping val="${stack === 'standard' ? 'clustered' : stack}"/><c:varyColors val="${vary}"/>${series}<c:gapWidth val="${Math.round(spec.gapWidth ?? 150)}"/>${overlap !== undefined ? `<c:overlap val="${Math.round(overlap)}"/>` : ''}${axIds}</c:barChart>`;
        }
        case 'line':
          return `<c:lineChart><c:grouping val="${stack}"/><c:varyColors val="0"/>${series}<c:marker val="1"/>${axIds}</c:lineChart>`;
        case 'area':
          return `<c:areaChart><c:grouping val="${stack}"/><c:varyColors val="0"/>${series}${axIds}</c:areaChart>`;
        case 'pie':
          return `<c:pieChart><c:varyColors val="1"/>${series}<c:firstSliceAng val="${Math.round(spec.firstSliceAngle ?? 0)}"/></c:pieChart>`;
        case 'doughnut':
          return `<c:doughnutChart><c:varyColors val="1"/>${series}<c:firstSliceAng val="${Math.round(spec.firstSliceAngle ?? 0)}"/><c:holeSize val="${Math.round(Math.min(90, Math.max(10, spec.holeSize ?? 50)))}"/></c:doughnutChart>`;
        case 'scatter': {
          const style = (spec.scatterStyle ?? 'markers').startsWith('smooth') ? 'smoothMarker' : 'lineMarker';
          return `<c:scatterChart><c:scatterStyle val="${style}"/><c:varyColors val="0"/>${series}${axIds}</c:scatterChart>`;
        }
        case 'bubble':
          return `<c:bubbleChart><c:varyColors val="0"/>${series}<c:bubbleScale val="${Math.round(spec.bubbleScale ?? 100)}"/><c:showNegBubbles val="0"/>${axIds}</c:bubbleChart>`;
        case 'radar':
          return `<c:radarChart><c:radarStyle val="${spec.radarFilled ? 'filled' : spec.markers ? 'marker' : 'standard'}"/><c:varyColors val="0"/>${series}${axIds}</c:radarChart>`;
        default:
          return '';
      }
    })
    .join('');

  const axTitle = (text?: string) => (text ? richTitle(text, 1000) : '');
  const gridXml = (on: boolean) => (on ? `<c:majorGridlines><c:spPr><a:ln w="9525">${solidFill('#D9D9D9')}</a:ln></c:spPr></c:majorGridlines>` : '');
  const minorGridXml = (on: boolean | undefined) => (on ? `<c:minorGridlines><c:spPr><a:ln w="6350">${solidFill('#EDEDED')}</a:ln></c:spPr></c:minorGridlines>` : '');
  const scalingXml = (a: AxisSpec | undefined, value: boolean) =>
    `<c:scaling>${value && a?.log ? '<c:logBase val="10"/>' : ''}<c:orientation val="${a?.reverse ? 'maxMin' : 'minMax'}"/>${value && a?.max !== undefined ? `<c:max val="${a.max}"/>` : ''}${value && a?.min !== undefined ? `<c:min val="${a.min}"/>` : ''}</c:scaling>`;
  const numFmtXml = (a: AxisSpec | undefined) => (a?.numFmt ? `<c:numFmt formatCode="${escapeXml(a.numFmt)}" sourceLinked="0"/>` : '<c:numFmt formatCode="General" sourceLinked="1"/>');
  const tickXml = (a: AxisSpec | undefined) => `<c:majorTickMark val="${a?.ticks ?? 'out'}"/><c:minorTickMark val="${a?.minorTicks ? 'out' : 'none'}"/><c:tickLblPos val="${a?.labels === 'none' ? 'none' : 'nextTo'}"/>`;
  const txPrXml = (a: AxisSpec | undefined) =>
    a?.labelAngle ? `<c:txPr><a:bodyPr rot="${Math.round(a.labelAngle * 60000)}" vert="horz"/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:endParaRPr lang="fr-FR"/></a:p></c:txPr>` : '';
  const unitXml = (a: AxisSpec | undefined) => (a?.majorUnit ? `<c:majorUnit val="${a.majorUnit}"/>` : '');

  const catAxisXml = (id: number, cross: number, pos: string, a: AxisSpec | undefined, title: string | undefined, deleted: boolean) =>
    `<c:catAx><c:axId val="${id}"/>${scalingXml(a, false)}<c:delete val="${deleted || a?.hidden ? 1 : 0}"/><c:axPos val="${pos}"/>${gridXml(a?.gridMajor === true)}${minorGridXml(a?.gridMinor)}${axTitle(title)}${numFmtXml(a)}${tickXml(a)}${txPrXml(a)}<c:crossAx val="${cross}"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>`;
  const valAxisXml = (id: number, cross: number, pos: string, a: AxisSpec | undefined, title: string | undefined, o: { grid: boolean; crosses?: string; between?: string }) =>
    `<c:valAx><c:axId val="${id}"/>${scalingXml(a, true)}<c:delete val="${a?.hidden ? 1 : 0}"/><c:axPos val="${pos}"/>${gridXml(a?.gridMajor ?? o.grid)}${minorGridXml(a?.gridMinor)}${axTitle(title)}${numFmtXml(a)}${tickXml(a)}${txPrXml(a)}<c:crossAx val="${cross}"/><c:crosses val="${o.crosses ?? 'autoZero'}"/><c:crossBetween val="${o.between ?? 'between'}"/>${unitXml(a)}</c:valAx>`;

  let axes = '';
  if (spec.type !== 'pie' && spec.type !== 'doughnut') {
    const catPos = horizontal ? 'l' : 'b';
    const valPos = horizontal ? 'b' : 'l';
    if (xy) {
      axes = valAxisXml(AX.cat, AX.val, 'b', spec.axes?.x, spec.xAxisTitle, { grid: spec.axes?.x?.gridMajor ?? true, between: 'midCat' }) + valAxisXml(AX.val, AX.cat, 'l', spec.axes?.y, spec.yAxisTitle, { grid: true, between: 'midCat' });
    } else {
      axes = catAxisXml(AX.cat, AX.val, catPos, spec.axes?.x, spec.xAxisTitle, false) + valAxisXml(AX.val, AX.cat, valPos, spec.axes?.y, spec.yAxisTitle, { grid: true, between: spec.type === 'radar' ? 'between' : 'between' });
      if (hasSecondary) {
        axes += catAxisXml(AX.cat2, AX.val2, catPos, undefined, undefined, true) + valAxisXml(AX.val2, AX.cat2, horizontal ? 't' : 'r', spec.axes?.y2, spec.y2AxisTitle, { grid: false, crosses: 'max' });
      }
    }
  }
  const dataTable = spec.dataTable && spec.type !== 'pie' && spec.type !== 'doughnut' && !xy && spec.type !== 'radar' ? '<c:dTable><c:showHorzBorder val="1"/><c:showVertBorder val="1"/><c:showOutline val="1"/><c:showKeys val="1"/></c:dTable>' : '';
  const legendPos = { bottom: 'b', top: 't', right: 'r', left: 'l' } as const;
  const legend = spec.legend === 'none' ? '' : `<c:legend><c:legendPos val="${legendPos[spec.legend ?? 'bottom'] ?? 'b'}"/><c:overlay val="${spec.legendOverlay ? 1 : 0}"/></c:legend>`;
  const blanks = spec.emptyCells === 'zero' ? 'zero' : spec.emptyCells === 'connect' ? 'span' : 'gap';
  const font = spec.fontSize && spec.fontSize !== 12 ? `<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="${Math.round(spec.fontSize * 100)}"/></a:pPr><a:endParaRPr lang="fr-FR"/></a:p></c:txPr>` : '';
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    '<c:lang val="fr-FR"/><c:roundedCorners val="0"/>' +
    '<c:chart>' +
    (spec.title ? richTitle(spec.title, 1400, spec.titleOverlay) + '<c:autoTitleDeleted val="0"/>' : '<c:autoTitleDeleted val="1"/>') +
    `<c:plotArea><c:layout/>${groupXml}${axes}${dataTable}</c:plotArea>` +
    legend +
    `<c:plotVisOnly val="${spec.plotVisibleOnly === false ? 0 : 1}"/><c:dispBlanksAs val="${blanks}"/></c:chart>` +
    `<c:spPr>${solidFill(spec.background ?? '#FFFFFF')}<a:ln w="9525">${spec.border === false ? '<a:noFill/>' : solidFill('#D9D9D9')}</a:ln></c:spPr>` +
    font +
    '</c:chartSpace>'
  );
}

function anchorXml(tag: 'from' | 'to', a: CellAnchor): string {
  return `<xdr:${tag}><xdr:col>${a.column}</xdr:col><xdr:colOff>${Math.round(a.columnOffset * EMU_PER_PX)}</xdr:colOff><xdr:row>${a.row}</xdr:row><xdr:rowOff>${Math.round(a.rowOffset * EMU_PER_PX)}</xdr:rowOff></xdr:${tag}>`;
}

/** Graphiques d'une feuille (depuis la ressource de dessins). */
export function chartsOfSheet(resource: DrawingResource | undefined, sheetId: string): SheetChart[] {
  const sheet = resource?.[sheetId];
  if (!sheet) return [];
  return (sheet.order ?? Object.keys(sheet.data ?? {}))
    .map((id) => sheet.data?.[id])
    .filter(
      (d): d is DrawingItem =>
        !!d && !!d.sheetTransform && ((d.componentKey === CHART_COMPONENT_KEY && !!d.data?.spec) || (d.componentKey === SHAPE_COMPONENT_KEY && !!d.data?.spec) || (d.drawingType === DRAWING_IMAGE_TYPE && !!d.source?.startsWith('data:image/'))),
    )
    .map((d) =>
      d.drawingType === DRAWING_IMAGE_TYPE
        ? { image: d.source, from: d.sheetTransform.from, to: d.sheetTransform.to }
        : d.componentKey === SHAPE_COMPONENT_KEY
          ? { shape: d.data!.spec as unknown as ShapeSpec, from: d.sheetTransform.from, to: d.sheetTransform.to }
          : { spec: d.data!.spec!, from: d.sheetTransform.from, to: d.sheetTransform.to },
    );
}

const EXTENSIONS: Record<string, string> = { 'image/png': 'png', 'image/jpeg': 'jpeg', 'image/gif': 'gif', 'image/bmp': 'bmp', 'image/webp': 'webp', 'image/svg+xml': 'svg' };

async function addDefaultContentType(zip: JSZip, ext: string, contentType: string): Promise<void> {
  const file = zip.file('[Content_Types].xml');
  if (!file) return;
  let xml = await file.async('string');
  if (new RegExp(`Extension="${ext}"`, 'i').test(xml)) return;
  xml = xml.replace(/<Types\b[^>]*>/, (m) => `${m}<Default Extension="${ext}" ContentType="${contentType}"/>`);
  zip.file('[Content_Types].xml', xml);
}

/**
 * Écrit les graphiques d'une feuille : partie dessin, parties graphiques, relations,
 * types de contenu, et <drawing r:id> dans la feuille. Renvoie le XML de feuille modifié.
 */
export async function writeSheetCharts(
  zip: JSZip,
  wb: IWorkbookData,
  sheetName: string,
  sheetPath: string,
  sheetXml: string,
  charts: SheetChart[],
  counters: { drawing: number; chart: number; image?: number; chartEx?: number },
): Promise<string> {
  if (!charts.length) return sheetXml;
  const drawingIndex = ++counters.drawing;
  const drawingPath = `xl/drawings/drawing${drawingIndex}.xml`;
  const anchors: string[] = [];
  for (const [i, c] of charts.entries()) {
    if (c.image) {
      // Image : fichier xl/media/imageN.ext + <xdr:pic>.
      const m = /^data:([^;]+);base64,(.*)$/.exec(c.image);
      if (!m) continue;
      const ext = EXTENSIONS[m[1]] ?? 'png';
      const n = (counters.image = (counters.image ?? 0) + 1);
      zip.file(`xl/media/image${n}.${ext}`, m[2], { base64: true });
      await addDefaultContentType(zip, ext, m[1]);
      const rid = await addRelationship(zip, drawingPath, REL_IMAGE, `../media/image${n}.${ext}`);
      anchors.push(
        `<xdr:twoCellAnchor editAs="oneCell">${anchorXml('from', c.from)}${anchorXml('to', c.to)}` +
          `<xdr:pic><xdr:nvPicPr><xdr:cNvPr${attrs({ id: i + 2, name: `Image ${i + 1}` })}/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>` +
          `<xdr:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>` +
          '<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:twoCellAnchor>',
      );
      continue;
    }
    if (c.slicer) {
      anchors.push(`<xdr:twoCellAnchor editAs="oneCell">${anchorXml('from', c.from)}${anchorXml('to', c.to)}${slicerFrameXml(c.slicer.name, i + 2, c.slicer.kind)}<xdr:clientData/></xdr:twoCellAnchor>`);
      continue;
    }
    if (c.shape) {
      anchors.push(`<xdr:twoCellAnchor editAs="oneCell">${anchorXml('from', c.from)}${anchorXml('to', c.to)}${shapeXml(c.shape, i + 2, `Forme ${i + 1}`)}<xdr:clientData/></xdr:twoCellAnchor>`);
      continue;
    }
    if (!c.spec) continue;
    if (isStatChart(c.spec.type)) {
      // Cascade, entonnoir, histogramme, boîte à moustaches : partie chartEx d'Excel 2016.
      const n = (counters.chartEx = (counters.chartEx ?? 0) + 1);
      const path = `xl/charts/chartEx${n}.xml`;
      zip.file(path, buildChartExXml(c.spec as never, (ref) => readValues(wb, ref, sheetName)));
      await addContentTypeOverride(zip, `/${path}`, CT_CHARTEX);
      // Style et palette : sans eux, Excel supprime tout le dessin à l'ouverture.
      zip.file(`xl/charts/style${n}.xml`, chartStyleXml());
      zip.file(`xl/charts/colors${n}.xml`, chartColorsXml());
      await addContentTypeOverride(zip, `/xl/charts/style${n}.xml`, CT_CHART_STYLE);
      await addContentTypeOverride(zip, `/xl/charts/colors${n}.xml`, CT_CHART_COLORS);
      await addRelationship(zip, path, REL_CHART_STYLE, `style${n}.xml`);
      await addRelationship(zip, path, REL_CHART_COLORS, `colors${n}.xml`);
      const rid = await addRelationship(zip, drawingPath, REL_CHARTEX, `../charts/chartEx${n}.xml`);
      anchors.push(`<xdr:twoCellAnchor editAs="oneCell">${anchorXml('from', c.from)}${anchorXml('to', c.to)}${chartExFrameXml(c.spec.type as never, i + 2, `Graphique ${i + 1}`, rid)}<xdr:clientData/></xdr:twoCellAnchor>`);
      continue;
    }
    const chartIndex = ++counters.chart;
    const chartPath = `xl/charts/chart${chartIndex}.xml`;
    zip.file(chartPath, buildChartXml(wb, c.spec, sheetName));
    await addContentTypeOverride(zip, `/${chartPath}`, CT_CHART);
    const rid = await addRelationship(zip, drawingPath, REL_CHART, `../charts/chart${chartIndex}.xml`);
    anchors.push(
      `<xdr:twoCellAnchor editAs="oneCell">${anchorXml('from', c.from)}${anchorXml('to', c.to)}` +
        `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr${attrs({ id: i + 2, name: `Graphique ${i + 1}` })}/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
        '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>' +
        `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="${rid}"/></a:graphicData></a:graphic>` +
        '</xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>',
    );
  }
  zip.file(
    drawingPath,
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      anchors.join('') +
      '</xdr:wsDr>',
  );
  await addContentTypeOverride(zip, `/${drawingPath}`, CT_DRAWING);
  const sheetRid = await addRelationship(zip, sheetPath, REL_DRAWING, `../drawings/drawing${drawingIndex}.xml`);
  let out = sheetXml;
  if (!/xmlns:r=/.test(out.slice(0, 2000))) out = out.replace(/<worksheet\b/, '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"');
  return insertWorksheetElement(out, 'drawing', `<drawing r:id="${sheetRid}"/>`);
}
