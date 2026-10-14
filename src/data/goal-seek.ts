/**
 * Valeur cible : trouve x tel que f(x) = cible (méthode de la sécante, avec repli par
 * dichotomie dès qu'un encadrement est trouvé). f est asynchrone : chaque évaluation
 * modifie une cellule et attend le recalcul.
 */

export interface GoalSeekResult {
  found: boolean;
  x: number;
  value: number;
  iterations: number;
}

export async function goalSeek(
  f: (x: number) => Promise<number> | number,
  target: number,
  x0: number,
  { maxIterations = 100, tolerance = 1e-7 } = {},
): Promise<GoalSeekResult> {
  const g = async (x: number) => (await f(x)) - target;
  let a = x0;
  let fa = await g(a);
  if (!Number.isFinite(fa)) return { found: false, x: a, value: fa + target, iterations: 0 };
  if (Math.abs(fa) <= tolerance) return { found: true, x: a, value: fa + target, iterations: 0 };
  let b = x0 === 0 ? 1 : x0 * 1.01 + 0.01;
  let fb = await g(b);
  // Encadrement [lo, hi] dès que les signes diffèrent.
  let lo: number | null = null;
  let hi: number | null = null;
  const bracket = (x: number, fx: number, y: number, fy: number) => {
    if (Number.isFinite(fx) && Number.isFinite(fy) && Math.sign(fx) !== Math.sign(fy)) {
      lo = x;
      hi = y;
    }
  };
  bracket(a, fa, b, fb);
  for (let i = 1; i <= maxIterations; i++) {
    if (Number.isFinite(fb) && Math.abs(fb) <= tolerance) return { found: true, x: b, value: fb + target, iterations: i };
    let next: number;
    if (lo !== null && hi !== null) {
      // Sécante bornée, sinon milieu.
      const flo = await g(lo);
      const fhi = await g(hi);
      next = lo - (flo * (hi - lo)) / (fhi - flo);
      if (!Number.isFinite(next) || next <= Math.min(lo, hi) || next >= Math.max(lo, hi)) next = (lo + hi) / 2;
      const fn = await g(next);
      if (Math.abs(fn) <= tolerance) return { found: true, x: next, value: fn + target, iterations: i };
      if (Math.sign(fn) === Math.sign(flo)) lo = next;
      else hi = next;
      a = b;
      fa = fb;
      b = next;
      fb = fn;
      if (Math.abs(hi - lo) < 1e-12 * Math.max(1, Math.abs(next))) return { found: Math.abs(fn) < 1e-6, x: next, value: fn + target, iterations: i };
      continue;
    }
    const denom = fb - fa;
    next = denom === 0 || !Number.isFinite(denom) ? b + (b - a || 1) * 2 : b - (fb * (b - a)) / denom;
    if (!Number.isFinite(next)) next = b * 2 + 1;
    const fn = await g(next);
    bracket(b, fb, next, fn);
    a = b;
    fa = fb;
    b = next;
    fb = fn;
  }
  return { found: Number.isFinite(fb) && Math.abs(fb) <= 1e-6, x: b, value: fb + target, iterations: maxIterations };
}
