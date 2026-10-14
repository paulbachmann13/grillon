/**
 * Modèle de graphique Grillon (indépendant d'ECharts et d'Excel) et conversions :
 * - plage sélectionnée → séries (comme « Insérer un graphique » d'Excel) ;
 * - spécification + valeurs → option ECharts.
 */
import { columnToLetters, rangeFromA1 } from '../io/xlsx/units';

export type ChartType = 'column' | 'bar' | 'line' | 'area' | 'pie' | 'scatter';

export interface SeriesSpec {
  /** Nom : texte fixe, ou référence (« Feuil1!$B$1 ») lue dans la feuille. */
  name?: string;
  nameRef?: string;
  /** Catégories (axe X) ou valeurs X (nuage de points). */
  categoriesRef?: string;
  valuesRef: string;
  /** Type propre à la série (graphique combiné). */
  type?: ChartType;
  secondaryAxis?: boolean;
  color?: string;
}

export interface ChartSpec {
  version: 1;
  type: ChartType;
  title?: string;
  series: SeriesSpec[];
  stacked?: boolean;
  legend?: 'bottom' | 'top' | 'right' | 'left' | 'none';
  xAxisTitle?: string;
  yAxisTitle?: string;
  y2AxisTitle?: string;
}

export const DEFAULT_COLORS = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47', '#264478', '#9E480E', '#636363', '#997300'];

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
): SeriesSpec[] {
  const rows = grid.length;
  const cols = Math.max(0, ...grid.map((r) => r.length));
  if (!rows || !cols) return [];
  const isText = (v: CellScalar) => typeof v === 'string' && v.trim() !== '';
  const headerRow = cols > 1 && grid[0].slice(1).some(isText) && !grid[0].slice(1).some((v) => typeof v === 'number');
  const firstDataRow = headerRow ? 1 : 0;
  const headerCol = cols > 1 && (grid.slice(firstDataRow).some((r) => isText(r[0])) || (headerRow && !isText(grid[0][0]) && grid[0][0] === null));
  const byRows = !headerCol && !headerRow ? false : cols - (headerCol ? 1 : 0) > rows - (headerRow ? 1 : 0) && type !== 'scatter' && type !== 'pie';
  const r0 = origin.row;
  const c0 = origin.col;
  const out: SeriesSpec[] = [];
  if (type === 'scatter') {
    // Nuage : la première colonne numérique donne X, les suivantes Y.
    const xCol = headerCol ? 0 : 0;
    for (let c = xCol + 1; c < cols; c++) {
      out.push({
        nameRef: headerRow ? sheetRef(sheet, r0, c0 + c, r0, c0 + c) : undefined,
        name: headerRow ? undefined : `Série ${c}`,
        categoriesRef: sheetRef(sheet, r0 + firstDataRow, c0 + xCol, r0 + rows - 1, c0 + xCol),
        valuesRef: sheetRef(sheet, r0 + firstDataRow, c0 + c, r0 + rows - 1, c0 + c),
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
  return type === 'pie' ? out.slice(0, 1) : out;
}

/** Valeurs résolues d'une série (par le composant, à partir de la feuille). */
export interface ResolvedSeries {
  spec: SeriesSpec;
  name: string;
  categories: CellScalar[];
  values: (number | null)[];
}

const toNumber = (v: CellScalar): number | null => {
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (typeof v === 'string' && v.trim() !== '' && Number.isFinite(Number(v.replace(',', '.')))) return Number(v.replace(',', '.'));
  return null;
};

export function resolveSeries(spec: ChartSpec, read: (ref: string) => CellScalar[]): ResolvedSeries[] {
  return spec.series.map((s, i) => {
    const values = read(s.valuesRef).map(toNumber);
    const categories = s.categoriesRef ? read(s.categoriesRef) : values.map((_, k) => k + 1);
    const name = s.name ?? (s.nameRef ? String(read(s.nameRef)[0] ?? '') : '') ?? `Série ${i + 1}`;
    return { spec: s, name: name || `Série ${i + 1}`, categories, values };
  });
}

const formatCategory = (v: CellScalar) => (v === null || v === undefined ? '' : typeof v === 'number' ? String(v).replace('.', ',') : String(v));
const formatNumber = (v: number) => (Math.abs(v) >= 1000 ? v.toLocaleString('fr-FR') : String(Number(v.toPrecision(10))).replace('.', ','));

/** Option ECharts pour un graphique (type principal + éventuelles séries d'autres types). */
export function buildEChartsOption(spec: ChartSpec, series: ResolvedSeries[], dark = false): Record<string, unknown> {
  const text = dark ? '#e6e6e6' : '#333';
  const grid = dark ? '#3a3a3a' : '#e5e5e5';
  const legendPos = spec.legend ?? 'bottom';
  const legend =
    legendPos === 'none'
      ? { show: false }
      : { show: true, type: 'scroll', textStyle: { color: text }, ...(legendPos === 'right' || legendPos === 'left' ? { orient: 'vertical', [legendPos]: 4, top: 'middle' } : { [legendPos]: 4 }) };
  const title = spec.title ? { text: spec.title, left: 'center', top: 6, textStyle: { color: text, fontSize: 14, fontWeight: 600 } } : undefined;
  const color = series.map((s, i) => s.spec.color ?? DEFAULT_COLORS[i % DEFAULT_COLORS.length]);

  if (spec.type === 'pie') {
    const s = series[0];
    return {
      title,
      color: DEFAULT_COLORS,
      legend,
      tooltip: { trigger: 'item', valueFormatter: (v: number) => formatNumber(v) },
      series: s
        ? [{ type: 'pie', name: s.name, radius: ['0%', '62%'], center: ['50%', '52%'], data: s.values.map((v, i) => ({ name: formatCategory(s.categories[i]), value: v ?? 0 })), label: { color: text } }]
        : [],
    };
  }

  const horizontal = spec.type === 'bar';
  const hasSecondary = series.some((s) => s.spec.secondaryAxis);
  const categories = (series[0]?.categories ?? []).map(formatCategory);
  const valueAxis = (name?: string, secondary = false) => ({
    type: 'value',
    name,
    nameLocation: 'middle',
    nameGap: 42,
    position: secondary ? (horizontal ? 'top' : 'right') : undefined,
    axisLabel: { color: text, formatter: (v: number) => formatNumber(v) },
    nameTextStyle: { color: text },
    splitLine: { show: !secondary, lineStyle: { color: grid } },
  });
  const isScatter = spec.type === 'scatter';
  const categoryAxis = isScatter
    ? { type: 'value', name: spec.xAxisTitle, nameLocation: 'middle', nameGap: 28, axisLabel: { color: text, formatter: (v: number) => formatNumber(v) }, splitLine: { lineStyle: { color: grid } } }
    : { type: 'category', data: categories, name: spec.xAxisTitle, nameLocation: 'middle', nameGap: 28, axisLabel: { color: text }, nameTextStyle: { color: text } };
  const valueAxes = [valueAxis(spec.yAxisTitle), ...(hasSecondary ? [valueAxis(spec.y2AxisTitle, true)] : [])];

  return {
    title,
    color,
    legend,
    tooltip: { trigger: isScatter ? 'item' : 'axis', valueFormatter: (v: number) => (v === null || v === undefined ? '' : formatNumber(v)) },
    grid: { left: 16, right: hasSecondary ? 24 : 16, top: spec.title ? 40 : 16, bottom: legendPos === 'bottom' ? 36 : 16, containLabel: true },
    xAxis: horizontal ? valueAxes : categoryAxis,
    yAxis: horizontal ? categoryAxis : valueAxes,
    series: series.map((s) => {
      const type = s.spec.type ?? spec.type;
      const base = {
        name: s.name,
        [horizontal ? 'xAxisIndex' : 'yAxisIndex']: s.spec.secondaryAxis && hasSecondary ? 1 : 0,
        emphasis: { focus: 'series' },
      };
      if (type === 'scatter' || isScatter) {
        return { ...base, type: 'scatter', data: s.values.map((v, i) => [toNumber(s.categories[i]) ?? i + 1, v]) };
      }
      if (type === 'line' || type === 'area') {
        return { ...base, type: 'line', data: s.values, smooth: false, symbolSize: 5, areaStyle: type === 'area' ? { opacity: 0.35 } : undefined, stack: spec.stacked && type === 'area' ? 'total' : undefined };
      }
      return { ...base, type: 'bar', data: s.values, stack: spec.stacked ? 'total' : undefined, barMaxWidth: 48 };
    }),
  };
}
