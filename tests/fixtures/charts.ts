/** Classeur Univer avec graphiques (tests unitaires et LibreOffice). */
import type { ChartSpec } from '../../src/charts/spec';
import type { IWorkbookData } from '../../src/model/univer-types';

export function chartWorkbook(specs: ChartSpec[]): IWorkbookData {
  const sheetId = 's1';
  const rows: (string | number)[][] = [
    ['Mois', 'Ventes', 'Coûts', 'Marge %'],
    ['Jan', 120, 80, 0.33],
    ['Fév', 150, 90, 0.4],
    ['Mar', 90, 70, 0.22],
  ];
  const cellData: Record<number, Record<number, object>> = {};
  rows.forEach((r, i) => (cellData[i] = Object.fromEntries(r.map((v, j) => [j, { v, t: typeof v === 'number' ? 2 : 1 }]))));
  const data: Record<string, object> = {};
  specs.forEach((spec, i) => {
    const id = `d${i}`;
    data[id] = {
      drawingId: id,
      drawingType: 8,
      componentKey: 'GrillonChart',
      sheetTransform: { from: { column: 5, columnOffset: 10, row: 1 + i * 16, rowOffset: 4 }, to: { column: 12, columnOffset: 0, row: 15 + i * 16, rowOffset: 0 } },
      transform: { left: 0, top: 0, width: 400, height: 280 },
      data: { chartId: `c${i}`, spec },
    };
  });
  return {
    id: 'w',
    name: 'w',
    appVersion: '1',
    locale: 'frFR' as IWorkbookData['locale'],
    styles: {},
    sheetOrder: [sheetId],
    sheets: { [sheetId]: { id: sheetId, name: 'Ventes 2026', cellData } },
    resources: [{ name: 'SHEET_DRAWING_PLUGIN', data: JSON.stringify({ [sheetId]: { data, order: Object.keys(data) } }) }],
  };
}

export const S = "'Ventes 2026'";
export const COLUMN: ChartSpec = {
  version: 1,
  type: 'column',
  title: 'Ventes et coûts',
  legend: 'bottom',
  xAxisTitle: 'Mois',
  yAxisTitle: 'Euros',
  series: [
    { nameRef: `${S}!$B$1`, categoriesRef: `${S}!$A$2:$A$4`, valuesRef: `${S}!$B$2:$B$4`, color: '#4472C4' },
    { nameRef: `${S}!$C$1`, categoriesRef: `${S}!$A$2:$A$4`, valuesRef: `${S}!$C$2:$C$4`, color: '#ED7D31' },
    { nameRef: `${S}!$D$1`, categoriesRef: `${S}!$A$2:$A$4`, valuesRef: `${S}!$D$2:$D$4`, color: '#70AD47', type: 'line', secondaryAxis: true },
  ],
};
export const PIE: ChartSpec = { version: 1, type: 'pie', legend: 'right', series: [{ name: 'Ventes', categoriesRef: `${S}!$A$2:$A$4`, valuesRef: `${S}!$B$2:$B$4` }] };
export const SCATTER: ChartSpec = { version: 1, type: 'scatter', legend: 'none', series: [{ nameRef: `${S}!$C$1`, categoriesRef: `${S}!$B$2:$B$4`, valuesRef: `${S}!$C$2:$C$4`, color: '#A5A5A5' }] };
export const STACKED_BAR: ChartSpec = { version: 1, type: 'bar', stacked: true, legend: 'top', series: [COLUMN.series[0], COLUMN.series[1]] };
