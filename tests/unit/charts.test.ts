import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { buildEChartsOption, resolveSeries, seriesFromRange, type ChartSpec } from '../../src/charts/spec';
import { exportXlsx } from '../../src/io/xlsx/export';
import { importXlsx } from '../../src/io/xlsx/import';
import type { IWorkbookData } from '../../src/model/univer-types';
import { chartWorkbook, COLUMN, PIE, S, SCATTER, STACKED_BAR } from '../fixtures/charts';

function chartsOf(wb: IWorkbookData): ChartSpec[] {
  const res = JSON.parse(wb.resources!.find((r) => r.name === 'SHEET_DRAWING_PLUGIN')!.data);
  const sheet = Object.values(res)[0] as { data: Record<string, { data: { spec: ChartSpec } }>; order: string[] };
  return sheet.order.map((id) => sheet.data[id].data.spec);
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
