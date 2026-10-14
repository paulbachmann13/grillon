/**
 * Graphiques : DrawingML (xl/drawings + xl/charts) ⇄ graphiques Grillon (DOM flottants Univer,
 * ressource SHEET_DRAWING_PLUGIN, données { chartId, spec }).
 *
 * Pris en charge : histogramme, barres, courbes, aires (groupés ou empilés), secteurs,
 * nuage de points, graphiques combinés et axe secondaire, titres, légende, couleurs.
 */
import type JSZip from 'jszip';
import type { ChartSpec, ChartType, SeriesSpec } from '../../charts/spec';
import { DEFAULT_COLORS, parseSheetRef } from '../../charts/spec';
import type { ICellData, IWorkbookData, IWorksheetData } from '../../model/univer-types';
import { addContentTypeOverride, addRelationship, attrs, escapeXml, extractAll, extractElement, insertWorksheetElement, parseXml, readRelationships, resolveRelTarget } from './ooxml';

export const DRAWING_RESOURCE = 'SHEET_DRAWING_PLUGIN';
export const CHART_COMPONENT_KEY = 'GrillonChart';
const DRAWING_DOM_TYPE = 8;
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
  allowTransform?: boolean;
}

export type DrawingResource = Record<string, { data: Record<string, DrawingItem>; order: string[] }>;

export interface SheetChart {
  spec: ChartSpec;
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
  doughnutChart: 'pie',
  scatterChart: 'scatter',
};

/** Convertit une partie xl/charts/chartN.xml en ChartSpec (null si rien d'exploitable). */
export function parseChartXml(xml: string): ChartSpec | null {
  const doc = parseXml<{ chartSpace?: { chart?: Xml } }>(xml);
  const chart = doc.chartSpace?.chart;
  const plot = chart?.plotArea as Xml | undefined;
  if (!plot) return null;
  // Axes : on repère l'axe de valeurs secondaire (celui qui croise au maximum / à droite).
  const valAxes = arr(plot.valAx as Xml | Xml[]);
  const secondaryAxIds = new Set(
    valAxes.filter((ax) => (ax.crosses as Xml | undefined)?.val === 'max' || ['r', 't'].includes(String((ax.axPos as Xml | undefined)?.val))).map((ax) => String((ax.axId as Xml).val)),
  );
  const series: SeriesSpec[] = [];
  let mainType: ChartType | null = null;
  let stacked = false;
  for (const [tag, baseType] of Object.entries(GROUP_TAGS)) {
    for (const group of arr(plot[tag] as Xml | Xml[])) {
      let type = baseType;
      if (tag.startsWith('bar') && (group.barDir as Xml | undefined)?.val === 'bar') type = 'bar';
      const grouping = String((group.grouping as Xml | undefined)?.val ?? '');
      if (grouping === 'stacked' || grouping === 'percentStacked') stacked = true;
      const axIds = arr(group.axId as Xml | Xml[]).map((a) => String(a.val));
      const secondary = axIds.some((id) => secondaryAxIds.has(id)) && type !== 'scatter';
      mainType ??= secondary ? null : type;
      for (const ser of arr(group.ser as Xml | Xml[])) {
        const tx = ser.tx as Xml | undefined;
        const s: SeriesSpec = {
          valuesRef: refOf(type === 'scatter' ? ser.yVal : ser.val) ?? '',
          categoriesRef: refOf(type === 'scatter' ? ser.xVal : ser.cat),
        };
        const nameRef = refOf(tx);
        if (nameRef) s.nameRef = nameRef;
        else if (literalOf(tx)) s.name = literalOf(tx);
        const color = colorOf(ser.spPr) ?? colorOf((ser.marker as Xml | undefined)?.spPr);
        if (color) s.color = color;
        s.type = type;
        if (secondary) s.secondaryAxis = true;
        if (s.valuesRef) series.push(s);
      }
    }
  }
  if (!series.length) return null;
  const type = mainType ?? series[0].type ?? 'column';
  // Le type n'est gardé par série que s'il diffère du type principal (graphique combiné).
  for (const s of series) if (s.type === type) delete s.type;
  const spec: ChartSpec = { version: 1, type, series };
  const title = richText(chart?.title);
  if (title) spec.title = title;
  if (stacked) spec.stacked = true;
  const legendPos = ((chart?.legend as Xml | undefined)?.legendPos as Xml | undefined)?.val;
  spec.legend = chart?.legend === undefined ? 'none' : legendPos === 't' ? 'top' : legendPos === 'r' ? 'right' : legendPos === 'l' ? 'left' : 'bottom';
  const axisTitle = (ax: Xml | undefined) => (ax ? richText(ax.title) : undefined);
  const catAx = arr(plot.catAx as Xml | Xml[])[0] ?? (type === 'scatter' ? valAxes[0] : undefined);
  const primaryVal = valAxes.find((ax) => !secondaryAxIds.has(String((ax.axId as Xml).val)) && ax !== catAx);
  const secondaryVal = valAxes.find((ax) => secondaryAxIds.has(String((ax.axId as Xml).val)));
  const x = axisTitle(catAx);
  const y = axisTitle(primaryVal);
  const y2 = axisTitle(secondaryVal);
  if (x) spec.xAxisTitle = x;
  if (y) spec.yAxisTitle = y;
  if (y2) spec.y2AxisTitle = y2;
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
      const chartRef = /<c:chart\b[^>]*\br:id="([^"]+)"/.exec(anchor.xml)?.[1];
      if (!chartRef) continue;
      const chartRel = relById.get(chartRef);
      const chartFile = chartRel && zip.file(resolveRelTarget(drawingPath, chartRel.Target));
      if (!chartFile) continue;
      const spec = parseChartXml(await chartFile.async('string'));
      if (!spec) continue;
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
      out.push({ spec, from, to });
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

function richTitle(text: string, size = 1400): string {
  return `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="${size}" b="1"/></a:pPr><a:r><a:rPr lang="fr-FR" sz="${size}" b="1"/><a:t>${escapeXml(text)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>`;
}

function solidFill(color: string): string {
  return `<a:solidFill><a:srgbClr val="${color.replace('#', '').toUpperCase()}"/></a:solidFill>`;
}

/** ChartSpec → XML de partie graphique. */
export function buildChartXml(wb: IWorkbookData, spec: ChartSpec, sheetName: string): string {
  const groups = new Map<string, { type: ChartType; secondary: boolean; series: { s: SeriesSpec; idx: number }[] }>();
  spec.series.forEach((s, idx) => {
    const type = spec.type === 'pie' || spec.type === 'scatter' ? spec.type : (s.type ?? spec.type);
    const secondary = !!s.secondaryAxis && type !== 'pie' && type !== 'scatter';
    const key = `${type === 'bar' ? 'column' : type}|${secondary}`;
    if (!groups.has(key)) groups.set(key, { type, secondary, series: [] });
    groups.get(key)!.series.push({ s, idx });
  });
  const hasSecondary = [...groups.values()].some((g) => g.secondary);
  const horizontal = spec.type === 'bar';
  const AX = { cat: 500000001, val: 500000002, cat2: 500000003, val2: 500000004 };

  const serXml = (s: SeriesSpec, idx: number, type: ChartType) => {
    const color = s.color ?? DEFAULT_COLORS[idx % DEFAULT_COLORS.length];
    const tx = s.nameRef
      ? `<c:tx><c:strRef><c:f>${fRef(s.nameRef)}</c:f>${strCache(readValues(wb, s.nameRef, sheetName))}</c:strRef></c:tx>`
      : `<c:tx><c:v>${escapeXml(s.name ?? `Série ${idx + 1}`)}</c:v></c:tx>`;
    const line = type === 'line' || type === 'scatter';
    const spPr = type === 'pie' ? '' : line ? `<c:spPr><a:ln w="28575" cap="rnd">${solidFill(color)}<a:round/></a:ln></c:spPr>` : `<c:spPr>${solidFill(color)}</c:spPr>`;
    const cats = readValues(wb, s.categoriesRef, sheetName);
    const catNumeric = cats.length > 0 && cats.every((v) => typeof v === 'number' || v === null);
    const cat = s.categoriesRef
      ? catNumeric
        ? `<c:numRef><c:f>${fRef(s.categoriesRef)}</c:f>${numCache(cats)}</c:numRef>`
        : `<c:strRef><c:f>${fRef(s.categoriesRef)}</c:f>${strCache(cats)}</c:strRef>`
      : '';
    const val = `<c:numRef><c:f>${fRef(s.valuesRef)}</c:f>${numCache(readValues(wb, s.valuesRef, sheetName))}</c:numRef>`;
    const marker = type === 'line' ? '<c:marker><c:symbol val="circle"/><c:size val="5"/></c:marker>' : type === 'scatter' ? `<c:marker><c:symbol val="circle"/><c:size val="6"/><c:spPr>${solidFill(color)}</c:spPr></c:marker>` : '';
    if (type === 'scatter') {
      return `<c:ser><c:idx val="${idx}"/><c:order val="${idx}"/>${tx}<c:spPr><a:ln w="19050"><a:noFill/></a:ln></c:spPr>${marker}${cat ? `<c:xVal>${cat}</c:xVal>` : ''}<c:yVal>${val}</c:yVal><c:smooth val="0"/></c:ser>`;
    }
    return `<c:ser><c:idx val="${idx}"/><c:order val="${idx}"/>${tx}${spPr}${type === 'column' || type === 'bar' ? '<c:invertIfNegative val="0"/>' : ''}${marker}${cat ? `<c:cat>${cat}</c:cat>` : ''}<c:val>${val}</c:val>${type === 'line' ? '<c:smooth val="0"/>' : ''}</c:ser>`;
  };

  const groupXml = [...groups.values()]
    .map((g) => {
      const series = g.series.map(({ s, idx }) => serXml(s, idx, g.type)).join('');
      const axIds = g.secondary ? `<c:axId val="${AX.cat2}"/><c:axId val="${AX.val2}"/>` : `<c:axId val="${AX.cat}"/><c:axId val="${AX.val}"/>`;
      const grouping = spec.stacked ? 'stacked' : g.type === 'line' ? 'standard' : 'clustered';
      switch (g.type) {
        case 'column':
        case 'bar':
          return `<c:barChart><c:barDir val="${horizontal ? 'bar' : 'col'}"/><c:grouping val="${spec.stacked ? 'stacked' : 'clustered'}"/><c:varyColors val="0"/>${series}<c:gapWidth val="150"/>${spec.stacked ? '<c:overlap val="100"/>' : ''}${axIds}</c:barChart>`;
        case 'line':
          return `<c:lineChart><c:grouping val="${spec.stacked ? 'stacked' : 'standard'}"/><c:varyColors val="0"/>${series}<c:marker val="1"/>${axIds}</c:lineChart>`;
        case 'area':
          return `<c:areaChart><c:grouping val="${grouping === 'clustered' ? 'standard' : grouping}"/><c:varyColors val="0"/>${series}${axIds}</c:areaChart>`;
        case 'pie':
          return `<c:pieChart><c:varyColors val="1"/>${series}<c:firstSliceAng val="0"/></c:pieChart>`;
        case 'scatter':
          return `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${series}${axIds}</c:scatterChart>`;
        default:
          return '';
      }
    })
    .join('');

  const axTitle = (text?: string) => (text ? richTitle(text, 1000) : '');
  const gridlines = '<c:majorGridlines><c:spPr><a:ln w="9525">' + solidFill('#D9D9D9') + '</a:ln></c:spPr></c:majorGridlines>';
  let axes = '';
  if (spec.type !== 'pie') {
    const catPos = horizontal ? 'l' : 'b';
    const valPos = horizontal ? 'b' : 'l';
    const catAx =
      spec.type === 'scatter'
        ? `<c:valAx><c:axId val="${AX.cat}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/>${axTitle(spec.xAxisTitle)}<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="${AX.val}"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx>`
        : `<c:catAx><c:axId val="${AX.cat}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${catPos}"/>${axTitle(spec.xAxisTitle)}<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="${AX.val}"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>`;
    const valAx = `<c:valAx><c:axId val="${AX.val}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${valPos}"/>${gridlines}${axTitle(spec.yAxisTitle)}<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="${AX.cat}"/><c:crosses val="autoZero"/><c:crossBetween val="${spec.type === 'scatter' ? 'midCat' : 'between'}"/></c:valAx>`;
    axes = catAx + valAx;
    if (hasSecondary) {
      axes +=
        `<c:catAx><c:axId val="${AX.cat2}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="1"/><c:axPos val="${catPos}"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="${AX.val2}"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>` +
        `<c:valAx><c:axId val="${AX.val2}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${horizontal ? 't' : 'r'}"/>${axTitle(spec.y2AxisTitle)}<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="${AX.cat2}"/><c:crosses val="max"/><c:crossBetween val="between"/></c:valAx>`;
    }
  }
  const legendPos = { bottom: 'b', top: 't', right: 'r', left: 'l' } as const;
  const legend = spec.legend === 'none' ? '' : `<c:legend><c:legendPos val="${legendPos[spec.legend ?? 'bottom'] ?? 'b'}"/><c:overlay val="0"/></c:legend>`;
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    '<c:lang val="fr-FR"/><c:roundedCorners val="0"/>' +
    '<c:chart>' +
    (spec.title ? richTitle(spec.title) + '<c:autoTitleDeleted val="0"/>' : '<c:autoTitleDeleted val="1"/>') +
    `<c:plotArea><c:layout/>${groupXml}${axes}</c:plotArea>` +
    legend +
    '<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>' +
    `<c:spPr>${solidFill('#FFFFFF')}<a:ln w="9525">${solidFill('#D9D9D9')}</a:ln></c:spPr>` +
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
    .filter((d): d is DrawingItem => !!d && d.componentKey === CHART_COMPONENT_KEY && !!d.data?.spec)
    .map((d) => ({ spec: d.data!.spec!, from: d.sheetTransform.from, to: d.sheetTransform.to }));
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
  counters: { drawing: number; chart: number },
): Promise<string> {
  if (!charts.length) return sheetXml;
  const drawingIndex = ++counters.drawing;
  const drawingPath = `xl/drawings/drawing${drawingIndex}.xml`;
  const anchors: string[] = [];
  for (const [i, c] of charts.entries()) {
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
