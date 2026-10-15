import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { buildEChartsOption, resolveSeries, seriesFromRange, type ChartSpec } from '../../src/charts/spec';
import { exportXlsx } from '../../src/io/xlsx/export';
import { importXlsx } from '../../src/io/xlsx/import';
import type { IWorkbookData } from '../../src/model/univer-types';
import { BUBBLE, chartWorkbook, COLUMN, DOUGHNUT, LINES, PIE, RADAR, RICH, S, SCATTER, SCATTER_LINES, STACKED_BAR } from '../fixtures/charts';

function chartsOf(wb: IWorkbookData): ChartSpec[] {
  const res = JSON.parse(wb.resources!.find((r) => r.name === 'SHEET_DRAWING_PLUGIN')!.data);
  const sheet = Object.values(res)[0] as { data: Record<string, { data: { spec: ChartSpec } }>; order: string[] };
  return sheet.order.map((id) => sheet.data[id]).filter((d) => d.data?.spec).map((d) => d.data.spec);
}

describe('graphiques — séries depuis une plage', () => {
  it('reconnaît en-têtes et catégories, une série par colonne', () => {
    const series = seriesFromRange('Feuil1', [['Mois', 'A', 'B'], ['Jan', 1, 2], ['Fév', 3, 4]], { row: 0, col: 0 }, 'column');
    expect(series).toEqual([
      { nameRef: 'Feuil1!$B$1', name: undefined, categoriesRef: 'Feuil1!$A$2:$A$3', valuesRef: 'Feuil1!$B$2:$B$3' },
      { nameRef: 'Feuil1!$C$1', name: undefined, categoriesRef: 'Feuil1!$A$2:$A$3', valuesRef: 'Feuil1!$C$2:$C$3' },
    ]);
  });
  it('construit une option ECharts avec axe secondaire', () => {
    const values: Record<string, (string | number)[]> = {
      [`${S}!$B$1`]: ['Ventes'],
      [`${S}!$C$1`]: ['Coûts'],
      [`${S}!$D$1`]: ['Marge %'],
      [`${S}!$A$2:$A$4`]: ['Jan', 'Fév', 'Mar'],
      [`${S}!$B$2:$B$4`]: [120, 150, 90],
      [`${S}!$C$2:$C$4`]: [80, 90, 70],
      [`${S}!$D$2:$D$4`]: [0.33, 0.4, 0.22],
    };
    const opt = buildEChartsOption(COLUMN, resolveSeries(COLUMN, (r) => values[r] ?? [])) as { series: { type: string; yAxisIndex: number; name: string }[]; yAxis: unknown[] };
    expect(opt.series.map((s) => [s.name, s.type, s.yAxisIndex])).toEqual([
      ['Ventes', 'bar', 0],
      ['Coûts', 'bar', 0],
      ['Marge %', 'line', 1],
    ]);
    expect(opt.yAxis).toHaveLength(2);
  });
});

describe('graphiques — .xlsx', () => {
  it('round-trip : histogramme combiné avec axe secondaire, secteurs, nuage, barres empilées', async () => {
    const specs = [COLUMN, PIE, SCATTER, STACKED_BAR];
    const out = await exportXlsx(chartWorkbook(specs));
    const back = await importXlsx(out);
    expect(chartsOf(back)).toEqual(specs);
  });

  it('round-trip : toutes les options (axes, étiquettes, tendances, barres d’erreur, tableau de données…)', async () => {
    const specs = [RICH, BUBBLE, DOUGHNUT, RADAR, LINES, SCATTER_LINES];
    const back = await importXlsx(await exportXlsx(chartWorkbook(specs)));
    const got = chartsOf(back);
    specs.forEach((spec, i) => expect(got[i], spec.type).toEqual(spec));
  });

  it('écrit des parties DrawingML valides (relations, types de contenu, ancres)', async () => {
    const zip = await JSZip.loadAsync(await exportXlsx(chartWorkbook([COLUMN])));
    const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
    expect(sheet).toMatch(/<drawing r:id="rId\d+"\/>/);
    const drawing = await zip.file('xl/drawings/drawing1.xml')!.async('string');
    expect(drawing).toContain('<xdr:from><xdr:col>5</xdr:col><xdr:colOff>95250</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>38100</xdr:rowOff></xdr:from>');
    const chart = await zip.file('xl/charts/chart1.xml')!.async('string');
    expect(chart).toContain("<c:f>'Ventes 2026'!$B$2:$B$4</c:f>");
    expect(chart).toContain('<c:pt idx="1"><c:v>150</c:v></c:pt>');
    expect(chart).toContain('<c:barChart><c:barDir val="col"/>');
    expect(chart).toContain('<c:lineChart>');
    expect(chart).toContain('<c:crosses val="max"/>');
    const types = await zip.file('[Content_Types].xml')!.async('string');
    expect(types).toContain('/xl/charts/chart1.xml');
    expect(types).toContain('/xl/drawings/drawing1.xml');
    expect(await zip.file('xl/drawings/_rels/drawing1.xml.rels')!.async('string')).toContain('../charts/chart1.xml');
  });
});

describe('images et notes — .xlsx', () => {
  it('round-trip d’une image ancrée et d’une note de cellule', async () => {
    const png = 'iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAYAAABytg0kAAAAEklEQVR4nGP4z8DwHwqZYBwAX5wF+0zQ2ewAAAAASUVORK5CYII=';
    const wb = chartWorkbook([COLUMN]);
    const drawing = JSON.parse(wb.resources![0].data);
    drawing.s1.data.img = {
      drawingId: 'img',
      drawingType: 0,
      imageSourceType: 'URL',
      source: `data:image/png;base64,${png}`,
      sheetTransform: { from: { column: 1, columnOffset: 0, row: 8, rowOffset: 0 }, to: { column: 2, columnOffset: 10, row: 10, rowOffset: 5 } },
      transform: { left: 0, top: 0, width: 100, height: 50 },
    };
    drawing.s1.order.push('img');
    wb.resources = [
      { name: 'SHEET_DRAWING_PLUGIN', data: JSON.stringify(drawing) },
      { name: 'SHEET_NOTE_PLUGIN', data: JSON.stringify({ s1: { 1: { 2: { note: 'À vérifier', width: 200, height: 100 } } } }) },
    ];
    const back = await importXlsx(await exportXlsx(wb));
    const res = JSON.parse(back.resources!.find((r) => r.name === 'SHEET_DRAWING_PLUGIN')!.data);
    const sheet = Object.values(res)[0] as { data: Record<string, { drawingType: number; source?: string; sheetTransform: object }>; order: string[] };
    const image = sheet.order.map((id) => sheet.data[id]).find((d) => d.drawingType === 0)!;
    expect(image.source).toBe(`data:image/png;base64,${png}`);
    expect(image.sheetTransform).toEqual({ from: { column: 1, columnOffset: 0, row: 8, rowOffset: 0 }, to: { column: 2, columnOffset: 10, row: 10, rowOffset: 5 } });
    expect(chartsOf(back)).toHaveLength(1);
    const notes = JSON.parse(back.resources!.find((r) => r.name === 'SHEET_NOTE_PLUGIN')!.data);
    expect((Object.values(notes)[0] as Record<number, Record<number, { note: string }>>)[1][2].note).toBe('À vérifier');
  });
});
