/**
 * Mini-graphiques (sparklines) : modèle, répartition des cellules de destination et option ECharts.
 */
export type SparklineType = 'line' | 'column' | 'winloss';

export interface SparklineSpec {
  type: SparklineType;
  /** Données : plage d'une ligne ou d'une colonne (« Feuil1!$B$2:$F$2 »). */
  dataRef: string;
  color?: string;
  /** Couleur des points mis en évidence (premier, dernier, haut, bas, négatifs). */
  highlight?: string;
  lineWidth?: number;
  markers?: boolean;
  high?: boolean;
  low?: boolean;
  first?: boolean;
  last?: boolean;
  negative?: boolean;
  emptyCells?: 'gap' | 'zero' | 'connect';
  min?: number;
  max?: number;
}

export const SPARK_COLOR = '#376092';
export const SPARK_HIGHLIGHT = '#D00000';

/** Cellule de destination et plage de données d'un mini-graphique. */
export interface SparkTarget {
  row: number;
  col: number;
  dataRef: string;
}

export interface Rect {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

/**
 * Répartit les lignes (ou colonnes) de la plage de données sur les cellules de l'emplacement,
 * comme Excel : emplacement en colonne → une ligne de données par cellule ; en ligne → une colonne.
 */
export function sparkTargets(data: Rect, location: Rect, ref: (r: Rect) => string): SparkTarget[] | null {
  const locRows = location.r2 - location.r1 + 1;
  const locCols = location.c2 - location.c1 + 1;
  const dataRows = data.r2 - data.r1 + 1;
  const dataCols = data.c2 - data.c1 + 1;
  const out: SparkTarget[] = [];
  if (locCols === 1 && locRows === dataRows) {
    for (let i = 0; i < dataRows; i++) out.push({ row: location.r1 + i, col: location.c1, dataRef: ref({ r1: data.r1 + i, c1: data.c1, r2: data.r1 + i, c2: data.c2 }) });
    return out;
  }
  if (locRows === 1 && locCols === dataCols) {
    for (let j = 0; j < dataCols; j++) out.push({ row: location.r1, col: location.c1 + j, dataRef: ref({ r1: data.r1, c1: data.c1 + j, r2: data.r2, c2: data.c1 + j }) });
    return out;
  }
  // Un seul mini-graphique : une plage d'une ligne ou d'une colonne dans une cellule.
  if (locRows === 1 && locCols === 1 && (dataRows === 1 || dataCols === 1)) return [{ row: location.r1, col: location.c1, dataRef: ref(data) }];
  return null;
}

export function sparklineOption(spec: SparklineSpec, values: (number | null)[], dark = false): Record<string, unknown> {
  const color = spec.color ?? (dark ? '#8fb4e3' : SPARK_COLOR);
  const hl = spec.highlight ?? SPARK_HIGHLIGHT;
  const nums = values.map((v) => (v === null && spec.emptyCells === 'zero' ? 0 : v));
  const defined = nums.filter((v): v is number => v !== null);
  const max = defined.length ? Math.max(...defined) : 0;
  const min = defined.length ? Math.min(...defined) : 0;
  const lastIndex = nums.length - 1;
  const highlighted = (i: number, v: number | null): boolean =>
    v !== null && ((!!spec.high && v === max) || (!!spec.low && v === min) || (!!spec.first && i === 0) || (!!spec.last && i === lastIndex) || (!!spec.negative && v < 0));
  const base = {
    animation: false,
    grid: { left: 3, right: 3, top: 4, bottom: 4 },
    xAxis: { type: 'category', show: false, data: nums.map((_, i) => i), boundaryGap: spec.type !== 'line' },
    tooltip: { show: false },
  };
  if (spec.type === 'line') {
    return {
      ...base,
      yAxis: { type: 'value', show: false, min: spec.min ?? 'dataMin', max: spec.max ?? 'dataMax' },
      series: [
        {
          type: 'line',
          data: nums.map((v, i) => (highlighted(i, v) ? { value: v, symbol: 'circle', symbolSize: 6, itemStyle: { color: hl } } : v)),
          symbol: spec.markers ? 'circle' : 'none',
          symbolSize: 4,
          lineStyle: { width: spec.lineWidth ?? 1.5, color },
          itemStyle: { color },
          connectNulls: spec.emptyCells === 'connect',
          silent: true,
        },
      ],
    };
  }
  if (spec.type === 'winloss') {
    return {
      ...base,
      yAxis: { type: 'value', show: false, min: -1.2, max: 1.2 },
      series: [
        {
          type: 'bar',
          silent: true,
          barCategoryGap: '25%',
          data: nums.map((v, i) => (v === null || v === 0 ? null : { value: v > 0 ? 1 : -1, itemStyle: { color: highlighted(i, v) || v < 0 ? hl : color } })),
        },
      ],
    };
  }
  return {
    ...base,
    yAxis: { type: 'value', show: false, min: spec.min, max: spec.max },
    series: [
      {
        type: 'bar',
        silent: true,
        barCategoryGap: '20%',
        data: nums.map((v, i) => (v === null ? null : { value: v, itemStyle: { color: highlighted(i, v) ? hl : color } })),
      },
    ],
  };
}
