/**
 * Courbes de tendance et barres d'erreur (calculs purs, comme Excel).
 */
import { formatGeneral } from './format';

export type TrendlineType = 'linear' | 'exponential' | 'logarithmic' | 'polynomial' | 'power' | 'movingAverage';

export interface TrendlineSpec {
  type: TrendlineType;
  /** Degré du polynôme (2 à 6). */
  order?: number;
  /** Période de la moyenne mobile. */
  period?: number;
  /** Prolongement en avant / en arrière, en unités de X. */
  forward?: number;
  backward?: number;
  showEquation?: boolean;
  showR2?: boolean;
  name?: string;
  color?: string;
}

export interface TrendResult {
  points: [number, number][];
  equation?: string;
  r2?: number;
}

/** Résout un système linéaire (élimination de Gauss avec pivot partiel). */
function solve(a: number[][], b: number[]): number[] | null {
  const n = b.length;
  const m = a.map((row, i) => [...row, b[i]!]);
  for (let col = 0; col < n; col++) {
    let pivot = col;
    for (let r = col + 1; r < n; r++) if (Math.abs(m[r]![col]!) > Math.abs(m[pivot]![col]!)) pivot = r;
    if (Math.abs(m[pivot]![col]!) < 1e-12) return null;
    [m[col], m[pivot]] = [m[pivot]!, m[col]!];
    for (let r = col + 1; r < n; r++) {
      const f = m[r]![col]! / m[col]![col]!;
      for (let c = col; c <= n; c++) m[r]![c]! -= f * m[col]![c]!;
    }
  }
  const x = new Array<number>(n).fill(0);
  for (let r = n - 1; r >= 0; r--) {
    let sum = m[r]![n]!;
    for (let c = r + 1; c < n; c++) sum -= m[r]![c]! * x[c]!;
    x[r] = sum / m[r]![r]!;
  }
  return x;
}

/** Ajuste y = c0 + c1·x + … + cn·xⁿ par moindres carrés. */
export function polyFit(xs: number[], ys: number[], order: number): number[] | null {
  const n = order + 1;
  if (xs.length < n) return null;
  const a: number[][] = Array.from({ length: n }, () => new Array<number>(n).fill(0));
  const b = new Array<number>(n).fill(0);
  for (let k = 0; k < xs.length; k++) {
    const pow = [1];
    for (let i = 1; i < 2 * n; i++) pow.push(pow[i - 1]! * xs[k]!);
    for (let i = 0; i < n; i++) {
      for (let j = 0; j < n; j++) a[i]![j]! += pow[i + j]!;
      b[i]! += pow[i]! * ys[k]!;
    }
  }
  return solve(a, b);
}

const evalPoly = (c: number[], x: number) => c.reduce((s, coef, i) => s + coef * x ** i, 0);

function r2Of(ys: number[], predicted: number[]): number {
  const mean = ys.reduce((s, v) => s + v, 0) / ys.length;
  const total = ys.reduce((s, v) => s + (v - mean) ** 2, 0);
  const residual = ys.reduce((s, v, i) => s + (v - predicted[i]!) ** 2, 0);
  return total === 0 ? 1 : 1 - residual / total;
}

const num = (v: number) => {
  const text = formatGeneral(Number(v.toPrecision(5)));
  return text.replace(/ /g, '');
};
const term = (coef: number, power: string, first: boolean) => {
  const sign = coef < 0 ? '-' : first ? '' : '+';
  const abs = Math.abs(coef);
  return `${first ? sign : ` ${sign} `}${num(abs)}${power}`;
};

/** Courbe de tendance de (xs, ys) ; seuls les points où X et Y sont des nombres comptent. */
export function trendline(rawX: (number | null)[], rawY: (number | null)[], spec: TrendlineSpec): TrendResult | null {
  const xs: number[] = [];
  const ys: number[] = [];
  rawX.forEach((x, i) => {
    const y = rawY[i];
    if (typeof x === 'number' && typeof y === 'number' && Number.isFinite(x) && Number.isFinite(y)) {
      xs.push(x);
      ys.push(y);
    }
  });
  if (xs.length < 2) return null;
  const xMin = Math.min(...xs) - (spec.backward ?? 0);
  const xMax = Math.max(...xs) + (spec.forward ?? 0);
  const sample = (f: (x: number) => number): [number, number][] => {
    const steps = 60;
    const out: [number, number][] = [];
    for (let i = 0; i <= steps; i++) {
      const x = xMin + ((xMax - xMin) * i) / steps;
      const y = f(x);
      if (Number.isFinite(y)) out.push([x, y]);
    }
    return out;
  };

  if (spec.type === 'movingAverage') {
    const period = Math.max(2, Math.round(spec.period ?? 2));
    const points: [number, number][] = [];
    for (let i = period - 1; i < xs.length; i++) {
      let sum = 0;
      for (let k = i - period + 1; k <= i; k++) sum += ys[k]!;
      points.push([xs[i]!, sum / period]);
    }
    return points.length ? { points } : null;
  }

  if (spec.type === 'linear' || spec.type === 'polynomial') {
    const order = spec.type === 'linear' ? 1 : Math.min(6, Math.max(2, Math.round(spec.order ?? 2)));
    const c = polyFit(xs, ys, order);
    if (!c) return null;
    const f = (x: number) => evalPoly(c, x);
    let equation = 'y = ';
    for (let p = order; p >= 0; p--) equation += term(c[p]!, p === 0 ? '' : p === 1 ? 'x' : `x^${p}`, p === order);
    return { points: sample(f), equation, r2: r2Of(ys, xs.map(f)) };
  }

  if (spec.type === 'exponential') {
    if (ys.some((y) => y <= 0)) return null;
    const c = polyFit(xs, ys.map(Math.log), 1);
    if (!c) return null;
    const a = Math.exp(c[0]!);
    const f = (x: number) => a * Math.exp(c[1]! * x);
    return { points: sample(f), equation: `y = ${num(a)}e^${num(c[1]!)}x`, r2: r2Of(ys, xs.map(f)) };
  }

  if (spec.type === 'logarithmic') {
    if (xs.some((x) => x <= 0)) return null;
    const c = polyFit(xs.map(Math.log), ys, 1);
    if (!c) return null;
    const f = (x: number) => c[1]! * Math.log(x) + c[0]!;
    return { points: sample((x) => (x > 0 ? f(x) : NaN)), equation: `y = ${term(c[1]!, 'ln(x)', true)}${c[0]! < 0 ? ' - ' : ' + '}${num(Math.abs(c[0]!))}`, r2: r2Of(ys, xs.map(f)) };
  }

  // Puissance : y = a·x^b
  if (xs.some((x) => x <= 0) || ys.some((y) => y <= 0)) return null;
  const c = polyFit(xs.map(Math.log), ys.map(Math.log), 1);
  if (!c) return null;
  const a = Math.exp(c[0]!);
  const f = (x: number) => a * x ** c[1]!;
  return { points: sample((x) => (x > 0 ? f(x) : NaN)), equation: `y = ${num(a)}x^${num(c[1]!)}`, r2: r2Of(ys, xs.map(f)) };
}

export type ErrorBarKind = 'fixed' | 'percent' | 'stddev';

export interface ErrorBarsSpec {
  kind: ErrorBarKind;
  /** Valeur fixe, pourcentage, ou nombre d'écarts types. */
  value?: number;
  sides?: 'both' | 'plus' | 'minus';
  noCap?: boolean;
}

/** Amplitude de l'erreur pour chaque valeur. */
export function errorAmounts(values: (number | null)[], spec: ErrorBarsSpec): (number | null)[] {
  const nums = values.filter((v): v is number => typeof v === 'number');
  const mean = nums.reduce((s, v) => s + v, 0) / (nums.length || 1);
  const sd = Math.sqrt(nums.reduce((s, v) => s + (v - mean) ** 2, 0) / Math.max(1, nums.length - 1));
  const k = spec.value ?? (spec.kind === 'percent' ? 5 : spec.kind === 'stddev' ? 1 : 1);
  return values.map((v) => {
    if (typeof v !== 'number') return null;
    if (spec.kind === 'fixed') return k;
    if (spec.kind === 'percent') return (Math.abs(v) * k) / 100;
    return sd * k;
  });
}
