/**
 * Calculs des graphiques statistiques d'Excel : histogramme (classes) et boîte à moustaches
 * (quartiles, moustaches, valeurs aberrantes), et cumul de la cascade.
 */

export interface Bin {
  from: number;
  to: number;
  count: number;
  label: string;
}

export interface BinOptions {
  /** Largeur des classes (sinon automatique, règle de Scott comme Excel). */
  width?: number;
  /** Nombre de classes (prioritaire sur la largeur). */
  count?: number;
  /** Classe de dépassement : « > valeur ». */
  overflow?: number;
  /** Classe de soupassement : « ≤ valeur ». */
  underflow?: number;
}

const fmt = (n: number) => String(Number(n.toPrecision(6))).replace('.', ',');

/** Classes de l'histogramme : [min, min+w], ]min+w, min+2w]… comme Excel. */
export function histogramBins(values: number[], opts: BinOptions = {}): Bin[] {
  const data = values.filter((v) => Number.isFinite(v));
  if (!data.length) return [];
  const sorted = [...data].sort((a, b) => a - b);
  const lo = opts.underflow ?? sorted[0];
  const hi = opts.overflow ?? sorted[sorted.length - 1];
  const inside = sorted.filter((v) => (opts.underflow === undefined || v > opts.underflow) && (opts.overflow === undefined || v <= opts.overflow));
  let width: number;
  if (opts.count && opts.count > 0) width = (hi - lo) / opts.count || 1;
  else if (opts.width && opts.width > 0) width = opts.width;
  else {
    const n = data.length;
    const mean = data.reduce((s, v) => s + v, 0) / n;
    const sd = Math.sqrt(data.reduce((s, v) => s + (v - mean) ** 2, 0) / Math.max(1, n - 1));
    width = sd > 0 ? (3.5 * sd) / Math.cbrt(n) : 1;
  }
  const bins: Bin[] = [];
  if (opts.underflow !== undefined) bins.push({ from: -Infinity, to: opts.underflow, count: sorted.filter((v) => v <= opts.underflow!).length, label: `≤${fmt(opts.underflow)}` });
  const count = Math.max(1, opts.count ?? Math.ceil((hi - lo) / width - 1e-9));
  for (let i = 0; i < count; i++) {
    const from = lo + i * width;
    const to = i === count - 1 ? Math.max(hi, from + width) : from + width;
    const first = i === 0 && opts.underflow === undefined;
    const n = inside.filter((v) => (first ? v >= from : v > from) && v <= to).length;
    bins.push({ from, to, count: n, label: `${first ? '[' : ']'}${fmt(from)}, ${fmt(to)}]` });
  }
  if (opts.overflow !== undefined) bins.push({ from: opts.overflow, to: Infinity, count: sorted.filter((v) => v > opts.overflow!).length, label: `>${fmt(opts.overflow)}` });
  return bins;
}

export interface BoxStats {
  min: number;
  q1: number;
  median: number;
  q3: number;
  max: number;
  mean: number;
  outliers: number[];
}

/** Quantile « exclusif » (QUARTILE.EXCLURE, méthode par défaut d'Excel pour les boîtes). */
function quantileExc(sorted: number[], p: number): number {
  const n = sorted.length;
  if (n === 1) return sorted[0];
  const h = (n + 1) * p;
  if (h <= 1) return sorted[0];
  if (h >= n) return sorted[n - 1];
  const i = Math.floor(h);
  return sorted[i - 1] + (h - i) * (sorted[i] - sorted[i - 1]);
}

/** Boîte à moustaches : moustaches jusqu'aux dernières valeurs à 1,5 écart interquartile. */
export function boxStats(values: number[]): BoxStats | null {
  const sorted = values.filter((v) => Number.isFinite(v)).sort((a, b) => a - b);
  if (!sorted.length) return null;
  const q1 = quantileExc(sorted, 0.25);
  const median = quantileExc(sorted, 0.5);
  const q3 = quantileExc(sorted, 0.75);
  const iqr = q3 - q1;
  const inside = sorted.filter((v) => v >= q1 - 1.5 * iqr && v <= q3 + 1.5 * iqr);
  return {
    min: inside[0] ?? sorted[0],
    q1,
    median,
    q3,
    max: inside[inside.length - 1] ?? sorted[sorted.length - 1],
    mean: sorted.reduce((s, v) => s + v, 0) / sorted.length,
    outliers: sorted.filter((v) => v < q1 - 1.5 * iqr || v > q3 + 1.5 * iqr),
  };
}

export interface WaterfallBar {
  base: number;
  value: number;
  kind: 'increase' | 'decrease' | 'total';
}

/** Cascade : chaque barre part du cumul précédent ; une barre « total » part de zéro. */
export function waterfallBars(values: (number | null)[], totals: number[] = []): WaterfallBar[] {
  let running = 0;
  return values.map((raw, i) => {
    const v = raw ?? 0;
    if (totals.includes(i)) {
      running = v;
      return { base: Math.min(0, v), value: Math.abs(v), kind: 'total' };
    }
    const start = running;
    running += v;
    return { base: Math.min(start, running), value: Math.abs(v), kind: v >= 0 ? 'increase' : 'decrease' };
  });
}
