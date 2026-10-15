/**
 * Modèle de graphique Grillon (indépendant d'ECharts et d'Excel) et conversions :
 * - plage sélectionnée → séries (comme « Insérer un graphique » d'Excel) ;
 * - spécification + valeurs → option ECharts.
 *
 * Les options reprennent celles d'Excel : types et sous-types, éléments du graphique (titres,
 * légende, étiquettes de données, quadrillage, tableau de données), axes (min, max, unité, log,
 * ordre inversé, format de nombre, graduations), séries (couleur, ligne, marqueurs, lissage,
 * courbes de tendance, barres d'erreur, couleurs de points) et mise en forme générale.
 */
import { columnToLetters, rangeFromA1 } from '../io/xlsx/units';
import { formatGeneral, formatNumberPattern } from './format';
import { errorAmounts, trendline, type ErrorBarsSpec, type TrendlineSpec } from './trendline';

export type { ErrorBarsSpec, TrendlineSpec, TrendlineType } from './trendline';

export type ChartType = 'column' | 'bar' | 'line' | 'area' | 'pie' | 'doughnut' | 'scatter' | 'bubble' | 'radar';
export type Grouping = 'standard' | 'stacked' | 'percent';
export type MarkerShape = 'circle' | 'square' | 'diamond' | 'triangle' | 'none';
export type LabelPosition = 'auto' | 'center' | 'insideEnd' | 'insideBase' | 'outsideEnd' | 'above' | 'below' | 'left' | 'right';
export type ScatterStyle = 'markers' | 'straight' | 'straightMarkers' | 'smooth' | 'smoothMarkers';
export type LegendPosition = 'bottom' | 'top' | 'right' | 'left' | 'none';

export interface DataLabelsSpec {
  show?: boolean;
  value?: boolean;
  category?: boolean;
  series?: boolean;
  percent?: boolean;
  position?: LabelPosition;
  numFmt?: string;
  separator?: string;
}

export interface AxisSpec {
  min?: number;
  max?: number;
  majorUnit?: number;
  log?: boolean;
  /** Valeurs en ordre inversé. */
  reverse?: boolean;
  numFmt?: string;
  /** Étiquettes : à côté de l'axe (défaut) ou masquées. */
  labels?: 'nextTo' | 'none';
  ticks?: 'none' | 'in' | 'out' | 'cross';
  minorTicks?: boolean;
  gridMajor?: boolean;
  gridMinor?: boolean;
  /** Axe supprimé. */
  hidden?: boolean;
  labelAngle?: number;
}

export interface SeriesSpec {
  /** Nom : texte fixe, ou référence (« Feuil1!$B$1 ») lue dans la feuille. */
  name?: string;
  nameRef?: string;
  /** Catégories (axe X) ou valeurs X (nuage de points, bulles). */
  categoriesRef?: string;
  valuesRef: string;
  /** Tailles des bulles. */
  sizesRef?: string;
  /** Type propre à la série (graphique combiné). */
  type?: ChartType;
  secondaryAxis?: boolean;
  color?: string;
  lineWidth?: number;
  dash?: 'solid' | 'dash' | 'dot';
  marker?: { shape?: MarkerShape; size?: number };
  smooth?: boolean;
  /** Opacité du remplissage (aires, radar), de 0 à 1. */
  opacity?: number;
  dataLabels?: DataLabelsSpec;
  trendlines?: TrendlineSpec[];
  errorBars?: ErrorBarsSpec;
  /** Couleur d'un point précis (index → couleur). */
  pointColors?: Record<number, string>;
}

export interface ChartSpec {
  version: 1;
  type: ChartType;
  title?: string;
  titleOverlay?: boolean;
  series: SeriesSpec[];
  /** Ancien indicateur d'empilement (remplacé par `grouping`). */
  stacked?: boolean;
  grouping?: Grouping;
  legend?: LegendPosition;
  legendOverlay?: boolean;
  xAxisTitle?: string;
  yAxisTitle?: string;
  y2AxisTitle?: string;
  axes?: { x?: AxisSpec; y?: AxisSpec; y2?: AxisSpec };
  dataLabels?: DataLabelsSpec;
  dataTable?: boolean;
  /** Courbes : marqueurs affichés (défaut : non). */
  markers?: boolean;
  smooth?: boolean;
  scatterStyle?: ScatterStyle;
  palette?: string;
  /** Une couleur par point (une seule série). */
  varyColors?: boolean;
  background?: string;
  border?: boolean;
  fontSize?: number;
  /** Histogramme / barres : espacement entre groupes (%) et chevauchement (-100 à 100). */
  gapWidth?: number;
  overlap?: number;
  /** Anneau : taille du trou (%). Secteurs : angle de départ (°) et éclatement (%). */
  holeSize?: number;
  firstSliceAngle?: number;
  explosion?: number;
  bubbleScale?: number;
  radarFilled?: boolean;
  emptyCells?: 'gap' | 'zero' | 'connect';
  /** N'afficher que les cellules visibles (défaut : oui). */
  plotVisibleOnly?: boolean;
}

export const PALETTES: Record<string, { label: string; colors: string[] }> = {
  office: { label: 'Office', colors: ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47', '#264478', '#9E480E', '#636363', '#997300'] },
  colorful: { label: 'Coloré', colors: ['#E8453C', '#F9BB2D', '#4EB35D', '#2B7BE6', '#8E44AD', '#16A5A5', '#E67E22', '#D81B60', '#7F8C8D', '#546E7A'] },
  blue: { label: 'Camaïeu de bleus', colors: ['#08306B', '#08519C', '#2171B5', '#4292C6', '#6BAED6', '#9ECAE1', '#C6DBEF', '#1B3F73', '#3D7DBF', '#85B7DD'] },
  green: { label: 'Camaïeu de verts', colors: ['#00441B', '#006D2C', '#238B45', '#41AB5D', '#74C476', '#A1D99B', '#C7E9C0', '#1B5E20', '#4C9A53', '#8BC48F'] },
  gray: { label: 'Nuances de gris', colors: ['#222222', '#4D4D4D', '#777777', '#999999', '#B3B3B3', '#CCCCCC', '#333333', '#666666', '#8C8C8C', '#A6A6A6'] },
  pastel: { label: 'Pastel', colors: ['#8DA0CB', '#FC8D62', '#66C2A5', '#E78AC3', '#A6D854', '#FFD92F', '#E5C494', '#B3B3B3', '#B2DF8A', '#CAB2D6'] },
};

export const DEFAULT_COLORS = PALETTES.office!.colors;

export type CellScalar = string | number | boolean | null;

/** Référence absolue « 'Feuille'!$A$1:$A$5 ». */
export function sheetRef(sheet: string, r1: number, c1: number, r2: number, c2: number): string {
  const quoted = /^[\p{L}_][\p{L}\p{N}_.]*$/u.test(sheet) ? sheet : `'${sheet.replace(/'/g, "''")}'`;
  const a = `$${columnToLetters(c1)}$${r1 + 1}`;
  const b = `$${columnToLetters(c2)}$${r2 + 1}`;
  return `${quoted}!${a === b ? a : `${a}:${b}`}`;
}

/** « 'Feuille'!$A$1:$B$2 » → feuille et plage (indices 0). */
export function parseSheetRef(ref: string): { sheet?: string; startRow: number; endRow: number; startColumn: number; endColumn: number } | null {
  const m = /^(?:(?:'((?:[^']|'')+)'|([^!]+))!)?(.+)$/.exec(ref.trim().replace(/^=/, ''));
  if (!m) return null;
  try {
    const range = rangeFromA1(m[3].replace(/\$/g, ''));
    return { sheet: m[1] !== undefined ? m[1].replace(/''/g, "'") : m[2], ...range };
  } catch {
    return null;
  }
}

export const groupingOf = (spec: ChartSpec): Grouping => spec.grouping ?? (spec.stacked ? 'stacked' : 'standard');

/**
 * Séries à partir d'une plage, comme Excel : la première ligne (si texte) donne les noms,
 * la première colonne (si texte ou dates) les catégories ; une série par colonne, ou par
 * ligne si la plage est plus large que haute.
 */
export function seriesFromRange(
  sheet: string,
  grid: CellScalar[][],
  origin: { row: number; col: number },
  type: ChartType,
  options: { byRows?: boolean } = {},
): SeriesSpec[] {
  const rows = grid.length;
  const cols = Math.max(0, ...grid.map((r) => r.length));
  if (!rows || !cols) return [];
  const isText = (v: CellScalar) => typeof v === 'string' && v.trim() !== '';
  const headerRow = cols > 1 && grid[0].slice(1).some(isText) && !grid[0].slice(1).some((v) => typeof v === 'number');
  const firstDataRow = headerRow ? 1 : 0;
  const headerCol = cols > 1 && (grid.slice(firstDataRow).some((r) => isText(r[0])) || (headerRow && !isText(grid[0][0]) && grid[0][0] === null));
  const xy = type === 'scatter' || type === 'bubble';
  const single = type === 'pie' || type === 'doughnut';
  const auto = !headerCol && !headerRow ? false : cols - (headerCol ? 1 : 0) > rows - (headerRow ? 1 : 0) && !xy && !single;
  const byRows = options.byRows ?? auto;
  const r0 = origin.row;
  const c0 = origin.col;
  const out: SeriesSpec[] = [];
  if (xy) {
    // Nuage / bulles : la première colonne donne X, les suivantes Y (et la taille des bulles).
    const step = type === 'bubble' ? 2 : 1;
    for (let c = 1; c < cols; c += step) {
      if (type === 'bubble' && c + 1 >= cols) break;
      out.push({
        nameRef: headerRow ? sheetRef(sheet, r0, c0 + c, r0, c0 + c) : undefined,
        name: headerRow ? undefined : `Série ${out.length + 1}`,
        categoriesRef: sheetRef(sheet, r0 + firstDataRow, c0, r0 + rows - 1, c0),
        valuesRef: sheetRef(sheet, r0 + firstDataRow, c0 + c, r0 + rows - 1, c0 + c),
        sizesRef: type === 'bubble' ? sheetRef(sheet, r0 + firstDataRow, c0 + c + 1, r0 + rows - 1, c0 + c + 1) : undefined,
      });
    }
    return out;
  }
  if (!byRows) {
    const firstCol = headerCol ? 1 : 0;
    for (let c = firstCol; c < cols; c++) {
      out.push({
        nameRef: headerRow ? sheetRef(sheet, r0, c0 + c, r0, c0 + c) : undefined,
        name: headerRow ? undefined : `Série ${c - firstCol + 1}`,
        categoriesRef: headerCol ? sheetRef(sheet, r0 + firstDataRow, c0, r0 + rows - 1, c0) : undefined,
        valuesRef: sheetRef(sheet, r0 + firstDataRow, c0 + c, r0 + rows - 1, c0 + c),
      });
    }
  } else {
    const firstCol = headerCol ? 1 : 0;
    for (let r = firstDataRow; r < rows; r++) {
      out.push({
        nameRef: headerCol ? sheetRef(sheet, r0 + r, c0, r0 + r, c0) : undefined,
        name: headerCol ? undefined : `Série ${r - firstDataRow + 1}`,
        categoriesRef: headerRow ? sheetRef(sheet, r0, c0 + firstCol, r0, c0 + cols - 1) : undefined,
        valuesRef: sheetRef(sheet, r0 + r, c0 + firstCol, r0 + r, c0 + cols - 1),
      });
    }
  }
  return single ? out.slice(0, 1) : out;
}

/** Valeurs résolues d'une série (par le composant, à partir de la feuille). */
export interface ResolvedSeries {
  spec: SeriesSpec;
  name: string;
  categories: CellScalar[];
  values: (number | null)[];
  sizes?: (number | null)[];
}

const toNumber = (v: CellScalar): number | null => {
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (typeof v === 'string' && v.trim() !== '' && Number.isFinite(Number(v.replace(',', '.')))) return Number(v.replace(',', '.'));
  return null;
};

export function resolveSeries(spec: ChartSpec, read: (ref: string, asDisplay?: boolean) => CellScalar[]): ResolvedSeries[] {
  return spec.series.map((s, i) => {
    const values = read(s.valuesRef).map(toNumber);
    const xy = (s.type ?? spec.type) === 'scatter' || (s.type ?? spec.type) === 'bubble';
    // Catégories : texte affiché (dates, formats…) ; abscisses d'un nuage : valeurs numériques.
    const categories = s.categoriesRef ? read(s.categoriesRef, !xy) : values.map((_, k) => k + 1);
    const name = s.name ?? (s.nameRef ? String(read(s.nameRef, true)[0] ?? '') : '') ?? `Série ${i + 1}`;
    const sizes = s.sizesRef ? read(s.sizesRef).map(toNumber) : undefined;
    return { spec: s, name: name || `Série ${i + 1}`, categories, values, sizes };
  });
}

const formatCategory = (v: CellScalar) => (v === null || v === undefined ? '' : typeof v === 'number' ? formatGeneral(v) : String(v));
const fmt = (v: number, pattern?: string) => formatNumberPattern(v, pattern);

const MARKER_SYMBOLS: Record<MarkerShape, string> = { circle: 'circle', square: 'rect', diamond: 'diamond', triangle: 'triangle', none: 'none' };

/** Contenu du tableau de données (sous le graphique). */
export function dataTableOf(series: ResolvedSeries[]): { header: string[]; rows: { name: string; color?: string; cells: string[] }[] } {
  const categories = (series[0]?.categories ?? []).map(formatCategory);
  return {
    header: categories,
    rows: series.map((s, i) => ({ name: s.name, color: s.spec.color ?? DEFAULT_COLORS[i % DEFAULT_COLORS.length], cells: s.values.map((v) => (v === null ? '' : fmt(v))) })),
  };
}

interface LabelCtx {
  horizontal: boolean;
  kind: 'bar' | 'line' | 'pie' | 'scatter';
  stackedFill: boolean;
  text: string;
  fontSize: number;
}

function labelOption(dl: DataLabelsSpec | undefined, ctx: LabelCtx): Record<string, unknown> {
  if (!dl?.show) return { show: false };
  const wantValue = dl.value ?? (!dl.category && !dl.series && !dl.percent);
  const sep = dl.separator ?? '; ';
  const pos: Record<string, string> = ctx.kind === 'bar'
    ? ctx.horizontal
      ? { center: 'inside', insideEnd: 'insideRight', insideBase: 'insideLeft', outsideEnd: 'right', auto: 'right' }
      : { center: 'inside', insideEnd: 'insideTop', insideBase: 'insideBottom', outsideEnd: 'top', auto: ctx.stackedFill ? 'inside' : 'top' }
    : ctx.kind === 'pie'
      ? { center: 'inside', insideEnd: 'inside', insideBase: 'inside', outsideEnd: 'outside', auto: 'outside' }
      : { center: 'top', above: 'top', below: 'bottom', left: 'left', right: 'right', auto: 'top' };
  return {
    show: true,
    position: pos[dl.position ?? 'auto'] ?? pos.auto,
    color: ctx.stackedFill && (dl.position === 'center' || dl.position === 'insideEnd' || dl.position === 'insideBase') ? '#fff' : ctx.text,
    fontSize: Math.max(8, ctx.fontSize - 1),
    formatter: (p: { value: unknown; data?: { raw?: number; value?: unknown }; name?: string; seriesName?: string; percent?: number }) => {
      const parts: string[] = [];
      if (dl.series) parts.push(p.seriesName ?? '');
      if (dl.category) parts.push(p.name ?? '');
      if (wantValue) {
        const raw = p.data?.raw ?? (Array.isArray(p.value) ? (p.value[ctx.horizontal ? 0 : 1] ?? p.value[1]) : p.value);
        if (typeof raw === 'number') parts.push(fmt(raw, dl.numFmt));
      }
      if (dl.percent && p.percent !== undefined) parts.push(`${formatGeneral(Math.round(p.percent * 10) / 10)} %`);
      return parts.filter((x) => x !== '').join(sep);
    },
  };
}

function axisOption(a: AxisSpec | undefined, o: { title?: string; kind: 'value' | 'category'; text: string; grid: string; fontSize: number; position?: string; defaultGrid: boolean; percentFormat?: boolean; data?: string[] }): Record<string, unknown> {
  const isValue = o.kind === 'value';
  const gridMajor = a?.gridMajor ?? o.defaultGrid;
  const ticks = a?.ticks ?? 'out';
  const minor = !!a?.minorTicks || !!a?.gridMinor;
  return {
    type: isValue && a?.log ? 'log' : o.kind,
    data: o.data,
    name: o.title,
    nameLocation: 'middle',
    nameGap: isValue ? 44 : 30,
    nameTextStyle: { color: o.text, fontSize: o.fontSize },
    position: o.position,
    show: !a?.hidden,
    inverse: a?.reverse || undefined,
    min: isValue ? a?.min : undefined,
    max: isValue ? a?.max : undefined,
    interval: isValue ? a?.majorUnit : undefined,
    boundaryGap: o.kind === 'category' ? true : undefined,
    axisLabel: {
      show: a?.labels !== 'none',
      color: o.text,
      fontSize: o.fontSize,
      rotate: a?.labelAngle,
      hideOverlap: true,
      formatter: isValue ? (v: number) => (o.percentFormat && !a?.numFmt ? `${formatGeneral(Math.round(v * 1000) / 10)} %` : fmt(v, a?.numFmt)) : undefined,
    },
    axisTick: { show: ticks !== 'none', inside: ticks === 'in', alignWithLabel: true, lineStyle: { color: o.grid } },
    minorTick: { show: isValue && minor, splitNumber: 5 },
    axisLine: { lineStyle: { color: o.grid } },
    splitLine: { show: gridMajor, lineStyle: { color: o.grid } },
    minorSplitLine: { show: isValue && !!a?.gridMinor, lineStyle: { color: o.grid, opacity: 0.5 } },
  };
}

/** Option ECharts pour un graphique (type principal + éventuelles séries d'autres types). */
export function buildEChartsOption(spec: ChartSpec, series: ResolvedSeries[], dark = false): Record<string, unknown> {
  const text = dark ? '#e6e6e6' : '#404040';
  const grid = dark ? '#3a3a3a' : '#d9d9d9';
  const fontSize = spec.fontSize ?? 12;
  const palette = PALETTES[spec.palette ?? 'office']?.colors ?? DEFAULT_COLORS;
  const legendPos = spec.legend ?? 'bottom';
  const grouping = groupingOf(spec);
  const title = spec.title ? { text: spec.title, left: 'center', top: 6, textStyle: { color: text, fontSize: fontSize + 3, fontWeight: 600 } } : undefined;
  const colorOf = (s: ResolvedSeries, i: number) => s.spec.color ?? palette[i % palette.length]!;
  const background = spec.background ?? (dark ? 'transparent' : '#ffffff');
  const emptyAsZero = spec.emptyCells === 'zero';
  const val = (v: number | null) => (v === null && emptyAsZero ? 0 : v);

  const legendNames = series.map((s) => s.name);
  const legendBase = (extraNames: string[]) =>
    legendPos === 'none'
      ? { show: false }
      : {
          show: true,
          type: 'scroll',
          data: [...legendNames, ...extraNames],
          textStyle: { color: text, fontSize },
          ...(legendPos === 'right' || legendPos === 'left' ? { orient: 'vertical', [legendPos]: 4, top: 'middle' } : { [legendPos]: 4 }),
        };
  const top = spec.title && !spec.titleOverlay ? fontSize + 34 : 16;
  const gridBox = (secondary: boolean) => ({
    left: 16 + (legendPos === 'left' && !spec.legendOverlay ? 90 : 0),
    right: (secondary ? 24 : 16) + (legendPos === 'right' && !spec.legendOverlay ? 90 : 0),
    top: top + (legendPos === 'top' && !spec.legendOverlay ? 28 : 0),
    bottom: 12 + (legendPos === 'bottom' && !spec.legendOverlay ? 30 : 0),
    containLabel: true,
  });

  // --- Secteurs et anneaux ---
  if (spec.type === 'pie' || spec.type === 'doughnut') {
    const s = series[0];
    const hole = spec.type === 'doughnut' ? spec.holeSize ?? 50 : 0;
    const dl = s?.spec.dataLabels ?? spec.dataLabels;
    return {
      backgroundColor: background,
      title,
      color: palette,
      legend: { ...legendBase([]), data: (s?.categories ?? []).map(formatCategory) },
      tooltip: { trigger: 'item', valueFormatter: (v: number) => fmt(v) },
      series: s
        ? [
            {
              type: 'pie',
              name: s.name,
              radius: [`${(hole * 0.62).toFixed(0)}%`, '62%'],
              center: ['50%', spec.title ? '56%' : '52%'],
              startAngle: 90 - (spec.firstSliceAngle ?? 0),
              selectedOffset: (spec.explosion ?? 0) * 0.6,
              data: s.values.map((v, i) => ({
                name: formatCategory(s.categories[i]),
                value: v ?? 0,
                itemStyle: { color: s.spec.pointColors?.[i] ?? palette[i % palette.length] },
                selected: (spec.explosion ?? 0) > 0,
              })),
              label: labelOption(dl, { horizontal: false, kind: 'pie', stackedFill: false, text, fontSize }),
              labelLine: { show: !!dl?.show },
            },
          ]
        : [],
    };
  }

  // --- Radar ---
  if (spec.type === 'radar') {
    const categories = (series[0]?.categories ?? []).map(formatCategory);
    const all = series.flatMap((s) => s.values.filter((v): v is number => v !== null));
    const max = spec.axes?.y?.max ?? (all.length ? Math.max(...all) * 1.1 : 1);
    const min = spec.axes?.y?.min ?? Math.min(0, ...all);
    return {
      backgroundColor: background,
      title,
      color: series.map(colorOf),
      legend: legendBase([]),
      tooltip: { trigger: 'item' },
      radar: { indicator: categories.map((name) => ({ name, max, min })), center: ['50%', '56%'], radius: '62%', axisName: { color: text, fontSize }, splitLine: { lineStyle: { color: grid } }, axisLine: { lineStyle: { color: grid } } },
      series: [
        {
          type: 'radar',
          data: series.map((s, i) => ({ name: s.name, value: s.values.map((v) => v ?? 0), lineStyle: { width: s.spec.lineWidth ?? 2, color: colorOf(s, i) }, itemStyle: { color: colorOf(s, i) }, areaStyle: spec.radarFilled ? { opacity: s.spec.opacity ?? 0.35 } : undefined, symbol: spec.markers === false ? 'none' : 'circle' })),
          label: labelOption(spec.dataLabels, { horizontal: false, kind: 'line', stackedFill: false, text, fontSize }),
        },
      ],
    };
  }

  // --- Graphiques cartésiens ---
  const horizontal = spec.type === 'bar';
  const isScatter = spec.type === 'scatter' || spec.type === 'bubble';
  const hasSecondary = series.some((s) => s.spec.secondaryAxis);
  const categories = (series[0]?.categories ?? []).map(formatCategory);
  const percent = grouping === 'percent';
  // Total par catégorie (graphiques 100 % empilés).
  const totals = categories.map((_, k) => series.reduce((sum, s) => sum + Math.abs(s.values[k] ?? 0), 0));
  const valueAxes = [
    axisOption(spec.axes?.y, { title: spec.yAxisTitle, kind: 'value', text, grid, fontSize, position: horizontal ? 'bottom' : 'left', defaultGrid: true, percentFormat: percent }),
    ...(hasSecondary ? [{ ...axisOption(spec.axes?.y2, { title: spec.y2AxisTitle, kind: 'value', text, grid, fontSize, position: horizontal ? 'top' : 'right', defaultGrid: false }), alignTicks: true }] : []),
  ];
  const categoryAxis = isScatter
    ? axisOption(spec.axes?.x, { title: spec.xAxisTitle, kind: 'value', text, grid, fontSize, position: 'bottom', defaultGrid: true })
    : axisOption(spec.axes?.x, { title: spec.xAxisTitle, kind: 'category', text, grid, fontSize, position: horizontal ? 'left' : 'bottom', defaultGrid: false, data: categories });
  if (horizontal) (categoryAxis as { inverse?: boolean }).inverse = !spec.axes?.x?.reverse; // 1re catégorie en haut, comme Excel

  const gap = spec.gapWidth ?? 150;
  const overlap = spec.overlap ?? (grouping !== 'standard' ? 100 : 0);
  const outSeries: Record<string, unknown>[] = [];
  const extraLegend: string[] = [];
  const axisIndexKey = horizontal ? 'xAxisIndex' : 'yAxisIndex';

  series.forEach((s, i) => {
    const type = s.spec.type ?? spec.type;
    const color = colorOf(s, i);
    const axisIndex = s.spec.secondaryAxis && hasSecondary ? 1 : 0;
    const dl = s.spec.dataLabels ?? spec.dataLabels;
    const base = { name: s.name, [axisIndexKey]: axisIndex, emphasis: { focus: 'series' }, connectNulls: spec.emptyCells === 'connect' };
    const lineStyle = { width: s.spec.lineWidth ?? 2.25, color, type: s.spec.dash === 'dash' ? 'dashed' : s.spec.dash === 'dot' ? 'dotted' : 'solid' };
    const point = (v: number | null, k: number) => {
      const color2 = s.spec.pointColors?.[k] ?? (spec.varyColors && series.length === 1 ? palette[k % palette.length] : undefined);
      const value = percent && totals[k] ? (v ?? 0) / totals[k]! : val(v);
      const item: Record<string, unknown> = { value: percent ? value : value, raw: v };
      if (color2) item.itemStyle = { color: color2 };
      return item;
    };
    if (type === 'scatter' || type === 'bubble') {
      const style = spec.scatterStyle ?? 'markers';
      const withLine = type === 'scatter' && style !== 'markers';
      const withMarkers = style === 'markers' || style === 'straightMarkers' || style === 'smoothMarkers';
      const sizes = s.sizes ?? [];
      const maxSize = Math.max(1e-9, ...sizes.filter((v): v is number => typeof v === 'number'));
      const scale = (spec.bubbleScale ?? 100) / 100;
      outSeries.push({
        ...base,
        type: withLine ? 'line' : 'scatter',
        smooth: withLine ? style.startsWith('smooth') : undefined,
        data: s.values.map((v, k) => [toNumber(s.categories[k] as CellScalar) ?? k + 1, val(v), sizes[k] ?? 0]),
        symbol: withMarkers || !withLine ? MARKER_SYMBOLS[s.spec.marker?.shape ?? 'circle'] : 'none',
        symbolSize: type === 'bubble' ? (d: number[]) => Math.max(4, Math.sqrt((d[2] ?? 0) / maxSize) * 48 * scale) : (s.spec.marker?.size ?? 7),
        itemStyle: { color, opacity: type === 'bubble' ? 0.7 : 1 },
        lineStyle: withLine ? lineStyle : undefined,
        label: labelOption(dl, { horizontal: false, kind: 'scatter', stackedFill: false, text, fontSize }),
      });
    } else if (type === 'line' || type === 'area') {
      const marker = s.spec.marker?.shape ?? (spec.markers ? 'circle' : 'none');
      outSeries.push({
        ...base,
        type: 'line',
        data: s.values.map(point),
        smooth: s.spec.smooth ?? spec.smooth ?? false,
        symbol: MARKER_SYMBOLS[marker],
        symbolSize: s.spec.marker?.size ?? 6,
        lineStyle,
        itemStyle: { color },
        areaStyle: type === 'area' ? { color, opacity: s.spec.opacity ?? (grouping === 'standard' ? 0.5 : 0.85) } : undefined,
        stack: grouping !== 'standard' && type === spec.type ? 'total' : undefined,
        label: labelOption(dl, { horizontal: false, kind: 'line', stackedFill: false, text, fontSize }),
      });
    } else {
      outSeries.push({
        ...base,
        type: 'bar',
        data: s.values.map(point),
        itemStyle: { color, borderRadius: 0 },
        stack: grouping !== 'standard' ? 'total' : undefined,
        barGap: `${-overlap}%`,
        barCategoryGap: `${((gap / (100 + gap)) * 100).toFixed(0)}%`,
        barMaxWidth: 80,
        label: labelOption(dl, { horizontal, kind: 'bar', stackedFill: grouping !== 'standard', text, fontSize }),
      });
    }

    // Courbes de tendance (pas sur les graphiques en barres horizontales ni 100 % empilés).
    if (!horizontal && !percent) {
      const xs = s.values.map((_, k) => (isScatter ? toNumber(s.categories[k] as CellScalar) : k + 1));
      (s.spec.trendlines ?? []).forEach((t) => {
        const result = trendline(xs, s.values, t);
        if (!result) return;
        const label = t.name ?? `${{ linear: 'Linéaire', exponential: 'Exponentielle', logarithmic: 'Logarithmique', polynomial: 'Polynomiale', power: 'Puissance', movingAverage: 'Moyenne mobile' }[t.type]} (${s.name})`;
        extraLegend.push(label);
        const last = result.points[result.points.length - 1];
        const caption = [t.showEquation ? result.equation : undefined, t.showR2 && result.r2 !== undefined ? `R² = ${formatGeneral(Number(result.r2.toFixed(4)))}` : undefined].filter(Boolean).join('\n');
        outSeries.push({
          name: label,
          type: 'line',
          [axisIndexKey]: axisIndex,
          data: result.points.map(([x, y]) => [isScatter ? x : x - 1, y]),
          symbol: 'none',
          silent: true,
          smooth: false,
          lineStyle: { width: 1.75, type: 'dashed', color: t.color ?? color },
          itemStyle: { color: t.color ?? color },
          markPoint: caption && last ? { symbol: 'none', data: [{ coord: [isScatter ? last[0] : last[0] - 1, last[1]], value: caption }], label: { show: true, formatter: caption, position: 'top', color: t.color ?? color, fontSize: fontSize - 1 } } : undefined,
          z: 5,
        });
      });
    }

    // Barres d'erreur verticales.
    if (s.spec.errorBars && !horizontal) {
      const eb = s.spec.errorBars;
      const amounts = errorAmounts(s.values, eb);
      const data = s.values.flatMap((v, k) => {
        const a = amounts[k];
        if (v === null || a === null || a === undefined) return [];
        return [[isScatter ? (toNumber(s.categories[k] as CellScalar) ?? k + 1) : k, v, eb.sides === 'plus' ? v : v - a, eb.sides === 'minus' ? v : v + a]];
      });
      const cap = eb.noCap ? 0 : 5;
      outSeries.push({
        type: 'custom',
        name: `${s.name} (erreur)`,
        [axisIndexKey]: axisIndex,
        silent: true,
        tooltip: { show: false },
        z: 10,
        data,
        encode: { x: 0, y: [1, 2, 3] },
        renderItem: (_: unknown, api: { value(i: number): number; coord(p: number[]): number[]; style(o: object): object }) => {
          const x = api.value(0);
          const low = api.coord([x, api.value(2)]);
          const high = api.coord([x, api.value(3)]);
          const style = api.style({ stroke: color, lineWidth: 1.25, fill: undefined });
          return {
            type: 'group',
            children: [
              { type: 'line', shape: { x1: high[0], y1: high[1], x2: low[0], y2: low[1] }, style },
              ...(cap ? [{ type: 'line', shape: { x1: high[0]! - cap, y1: high[1], x2: high[0]! + cap, y2: high[1] }, style }, { type: 'line', shape: { x1: low[0]! - cap, y1: low[1], x2: low[0]! + cap, y2: low[1] }, style }] : []),
            ],
          };
        },
      });
    }
  });

  return {
    backgroundColor: background,
    title,
    color: series.map(colorOf),
    legend: legendBase(extraLegend),
    textStyle: { fontSize },
    tooltip: { trigger: isScatter ? 'item' : 'axis', valueFormatter: (v: number) => (v === null || v === undefined ? '' : fmt(v)) },
    grid: gridBox(hasSecondary),
    xAxis: horizontal ? valueAxes : categoryAxis,
    yAxis: horizontal ? categoryAxis : valueAxes,
    series: outSeries,
  };
}
