/** Classeur avec une source de données et un tableau croisé dynamique (tests unitaires et LibreOffice). */
import type { PivotDefinition } from '../../src/data/pivot';
import type { IWorkbookData } from '../../src/model/univer-types';

export const PIVOT_DEF: PivotDefinition = {
  sourceSheet: 'Ventes',
  sourceRange: 'A1:D7',
  config: {
    rows: [0],
    columns: [1],
    values: [{ col: 2, fn: 'sum' }, { col: 3, fn: 'average' }],
    filters: [],
  },
  outRows: 6,
  outCols: 6,
};

export function pivotWorkbook(def: PivotDefinition = PIVOT_DEF): IWorkbookData {
  const rows: (string | number)[][] = [
    ['Ville', 'Produit', 'Montant', 'Quantité'],
    ['Lyon', 'Café', 10, 1],
    ['Paris', 'Thé', 5, 2],
    ['Lyon', 'Thé', 7, 3],
    ['Lyon', 'Café', 3, 4],
    ['Nîmes', 'Café', 8, 5],
    ['Paris', 'Café', 12, 6],
  ];
  const cellData: Record<number, Record<number, object>> = {};
  rows.forEach((r, i) => (cellData[i] = Object.fromEntries(r.map((v, j) => [j, { v, t: typeof v === 'number' ? 2 : 1 }]))));
  return {
    id: 'w',
    name: 'w',
    appVersion: '1',
    locale: 'frFR' as IWorkbookData['locale'],
    styles: {},
    sheetOrder: ['s1', 's2'],
    sheets: {
      s1: { id: 's1', name: 'Ventes', cellData },
      s2: { id: 's2', name: 'TCD', cellData: { 2: { 0: { v: 'Étiquettes de lignes', t: 1 } } }, custom: { grillon: { pivot: def } } } as never,
    },
  };
}
