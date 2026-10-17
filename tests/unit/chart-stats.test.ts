import { describe, expect, it } from 'vitest';
import { boxStats, histogramBins, waterfallBars } from '../../src/charts/stats';

describe('graphiques statistiques', () => {
  it('histogramme : classes de largeur fixe, à la manière d’Excel', () => {
    const bins = histogramBins([1, 2, 2, 3, 5, 6, 9, 10], { width: 3 });
    expect(bins.map((b) => b.label)).toEqual(['[1, 4]', ']4, 7]', ']7, 10]']);
    expect(bins.map((b) => b.count)).toEqual([4, 2, 2]);
  });

  it('histogramme : dépassement et soupassement', () => {
    const bins = histogramBins([1, 2, 3, 4, 5, 6, 7, 8, 9, 10], { count: 2, underflow: 2, overflow: 8 });
    expect(bins.map((b) => b.label)).toEqual(['≤2', ']2, 5]', ']5, 8]', '>8']);
    expect(bins.map((b) => b.count)).toEqual([2, 3, 3, 2]);
  });

  it('histogramme automatique : toutes les valeurs comptées', () => {
    const values = Array.from({ length: 50 }, (_, i) => (i * 37) % 23);
    expect(histogramBins(values).reduce((s, b) => s + b.count, 0)).toBe(50);
  });

  it('boîte à moustaches : quartiles exclusifs et valeur aberrante', () => {
    const s = boxStats([1, 2, 3, 4, 5, 6, 7, 8, 9, 40])!;
    expect(s.median).toBe(5.5);
    expect(s.q1).toBe(2.75);
    expect(s.q3).toBe(8.25);
    expect(s.max).toBe(9);
    expect(s.outliers).toEqual([40]);
    expect(s.mean).toBe(8.5);
  });

  it('cascade : cumul et total', () => {
    expect(waterfallBars([100, -30, 20, null], [3]).map((b) => [b.base, b.value, b.kind])).toEqual([
      [0, 100, 'increase'],
      [70, 30, 'decrease'],
      [70, 20, 'increase'],
      [0, 0, 'total'],
    ]);
  });
});
