import { describe, expect, it } from 'vitest';
import { cellText, type Grid } from '../../src/data/cells';
import { computePivot, distinctValues } from '../../src/data/pivot';

const g = (rows: (string | number | null)[][]): Grid => rows.map((r) => r.map((v) => (v === null ? null : { v, t: typeof v === 'number' ? 2 : 1 })));
const texts = (grid: Grid) => grid.map((r) => r.map((c) => cellText(c)));

const data = g([
  ['Ville', 'Produit', 'Montant', 'Qté'],
  ['Lyon', 'Café', 10, 1],
  ['Paris', 'Thé', 5, 2],
  ['Lyon', 'Thé', 7, 1],
  ['Lyon', 'Café', 3, 3],
  ['Nîmes', 'Café', 8, 1],
]);

describe('tableau croisé dynamique', () => {
  it('lignes + somme, avec total général', () => {
    const out = computePivot(data, { rows: [0], columns: [], values: [{ col: 2, fn: 'sum' }], filters: [] });
    expect(texts(out)).toEqual([
      ['Étiquettes de lignes', 'Somme de Montant'],
      ['Lyon', '20'],
      ['Nîmes', '8'],
      ['Paris', '5'],
      ['Total général', '33'],
    ]);
  });

  it('lignes × colonnes avec totaux', () => {
    const out = computePivot(data, { rows: [0], columns: [1], values: [{ col: 2, fn: 'sum' }], filters: [] });
    expect(texts(out)).toEqual([
      ['Somme de Montant', 'Café', 'Thé'],
      ['Étiquettes de lignes', '', '', 'Total général'],
      ['Lyon', '13', '7', '20'],
      ['Nîmes', '8', '', '8'],
      ['Paris', '', '5', '5'],
      ['Total général', '21', '12', '33'],
    ]);
  });

  it('plusieurs valeurs (moyenne, nombre, min, max) et filtre', () => {
    const out = computePivot(data, {
      rows: [1],
      columns: [],
      values: [
        { col: 2, fn: 'average' },
        { col: 3, fn: 'count' },
        { col: 2, fn: 'min' },
        { col: 2, fn: 'max' },
      ],
      filters: [{ col: 0, include: ['Lyon'] }],
    });
    expect(texts(out)).toEqual([
      ['Étiquettes de lignes', 'Moyenne de Montant', 'Nombre de Qté', 'Min de Montant', 'Max de Montant'],
      ['Café', '6,5'.replace(',', '.'), '2', '3', '10'],
      ['Thé', '7', '1', '7', '7'],
      ['Total général', '6.666666666666667', '3', '3', '10'],
    ]);
  });

  it('valeurs distinctes pour les filtres', () => {
    expect(distinctValues(data, 0)).toEqual(['Lyon', 'Nîmes', 'Paris']);
  });
});
