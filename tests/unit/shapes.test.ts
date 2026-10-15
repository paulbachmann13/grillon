import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { resolveShape, type ShapeSpec } from '../../src/charts/shape';
import { exportXlsx } from '../../src/io/xlsx/export';
import { importXlsx } from '../../src/io/xlsx/import';
import { parseShapeAnchor, shapeXml } from '../../src/io/xlsx/shapes';
import { chartsOfSheet, type DrawingResource } from '../../src/io/xlsx/charts';
import { chartWorkbook } from '../fixtures/charts';

const SHAPES: ShapeSpec[] = [
  { kind: 'rect', text: 'Total\nannuel', fill: '#FFC000', stroke: '#7F6000', strokeWidth: 2.5, fontSize: 14, color: '#000000', bold: true, italic: true, align: 'left' },
  { kind: 'roundRect', text: 'Bravo' },
  { kind: 'ellipse', fill: 'none', stroke: 'none' },
  { kind: 'line', stroke: '#FF0000', strokeWidth: 3 },
  { kind: 'arrow' },
  { kind: 'textbox', text: 'Note : voir Feuil2', align: 'right', fontSize: 9 },
];

describe('formes et zones de texte', () => {
  it('valeurs par défaut selon le type', () => {
    expect(resolveShape({ kind: 'rect' })).toMatchObject({ fill: '#4472C4', color: '#FFFFFF', align: 'center' });
    expect(resolveShape({ kind: 'textbox' })).toMatchObject({ fill: 'none', color: '#000000', align: 'left' });
    expect(resolveShape({ kind: 'arrow' }).fill).toBe('none');
  });

  it('DrawingML : écriture puis lecture de chaque forme', () => {
    for (const spec of SHAPES) {
      const xml = `<xdr:twoCellAnchor xmlns:xdr="x" xmlns:a="y"><xdr:from/><xdr:to/>${shapeXml(spec, 2, 'F')}<xdr:clientData/></xdr:twoCellAnchor>`;
      expect(parseShapeAnchor(xml), spec.kind).toEqual(spec);
    }
  });

  it('round-trip .xlsx : les formes sont écrites dans le dessin puis relues', async () => {
    const wb = chartWorkbook([]);
    const data: Record<string, unknown> = {};
    SHAPES.forEach((spec, i) => {
      data[`s${i}`] = {
        drawingId: `s${i}`,
        drawingType: 8,
        componentKey: 'GrillonShape',
        sheetTransform: { from: { column: 6, columnOffset: 0, row: i * 4, rowOffset: 0 }, to: { column: 9, columnOffset: 0, row: i * 4 + 3, rowOffset: 0 } },
        transform: { left: 0, top: 0, width: 100, height: 60 },
        data: { shapeId: `s${i}`, spec },
      };
    });
    wb.resources = [{ name: 'SHEET_DRAWING_PLUGIN', data: JSON.stringify({ s1: { data, order: Object.keys(data) } }) }];
    const bytes = await exportXlsx(wb);
    const drawing = await (await JSZip.loadAsync(bytes)).file('xl/drawings/drawing1.xml')!.async('string');
    expect(drawing).toContain('<xdr:sp ');
    expect(drawing).toContain('<xdr:cxnSp ');
    expect(drawing).toContain('prst="roundRect"');
    expect(drawing).toContain('<a:tailEnd type="triangle"/>');
    const back = await importXlsx(bytes);
    const res = JSON.parse(back.resources!.find((r) => r.name === 'SHEET_DRAWING_PLUGIN')!.data) as DrawingResource;
    const found = chartsOfSheet(res, Object.keys(res)[0]!);
    expect(found.map((f) => f.shape)).toEqual(SHAPES);
  });
});
