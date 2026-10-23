/** Classeur avec segments (tableau, tableau croisé) et chronologie (tests unitaires et Excel). */
import type { PivotDefinition } from '../../src/data/pivot';
import type { SlicerSpec } from '../../src/data/slicer';
import type { IWorkbookData } from '../../src/model/univer-types';
import { PIVOT_DEF, pivotWorkbook } from './pivot';

export const SLICER_PIVOT: PivotDefinition = { ...PIVOT_DEF, config: { ...PIVOT_DEF.config, columns: [], filters: [{ col: 1, include: ['Café'] }] }, slicerCols: [1] };

export const TABLE_SLICER: SlicerSpec = { id: 'sl-ville', type: 'slicer', title: 'Ville', target: { kind: 'table', sheet: 'Ventes', ref: 'A1:D7', col: 0 }, selected: ['Lyon', 'Paris'] };
export const PIVOT_SLICER: SlicerSpec = { id: 'sl-produit', type: 'slicer', title: 'Produit', target: { kind: 'pivot', sheet: 'TCD', col: 1 }, selected: ['Café'] };
export const TIMELINE: SlicerSpec = { id: 'tl', type: 'timeline', title: 'Montant', target: { kind: 'table', sheet: 'Ventes', ref: 'A1:D7', col: 2 }, selected: null, level: 'months', from: null, to: null };

const item = (id: string, spec: SlicerSpec, column: number) => ({
  drawingId: id,
  drawingType: 8,
  componentKey: 'GrillonSlicer',
  sheetTransform: { from: { column, columnOffset: 0, row: 1, rowOffset: 0 }, to: { column: column + 3, columnOffset: 0, row: 14, rowOffset: 0 } },
  transform: { left: 0, top: 0, width: 200, height: 260 },
  data: spec,
  allowTransform: true,
});

export function slicerWorkbook(specs: { sheet: 's1' | 's2'; spec: SlicerSpec }[] = [
  { sheet: 's1', spec: TABLE_SLICER },
  { sheet: 's1', spec: TIMELINE },
  { sheet: 's2', spec: PIVOT_SLICER },
]): IWorkbookData {
  const wb = pivotWorkbook(SLICER_PIVOT);
  // Lignes masquées par le segment « Ville » (Nîmes).
  wb.sheets.s1!.rowData = { 5: { hd: 1 } } as never;
  const drawings: Record<string, { data: Record<string, object>; order: string[] }> = {};
  specs.forEach(({ sheet, spec }, i) => {
    const d = (drawings[sheet] ??= { data: {}, order: [] });
    d.data[`d${i}`] = item(`d${i}`, spec, 6 + i * 4);
    d.order.push(`d${i}`);
  });
  return { ...wb, resources: [{ name: 'SHEET_DRAWING_PLUGIN', data: JSON.stringify(drawings) }] };
}
