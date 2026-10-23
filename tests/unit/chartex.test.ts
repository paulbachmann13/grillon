import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import type { ChartSpec } from '../../src/charts/spec';
import { exportXlsx } from '../../src/io/xlsx/export';
import { importXlsx } from '../../src/io/xlsx/import';
import type { IWorkbookData } from '../../src/model/univer-types';
import { BOX, chartWorkbook, COLUMN, FUNNEL, HISTOGRAM, WATERFALL } from '../fixtures/charts';

function chartsOf(wb: IWorkbookData): ChartSpec[] {
  const res = JSON.parse(wb.resources!.find((r) => r.name === 'SHEET_DRAWING_PLUGIN')!.data);
  const sheet = Object.values(res)[0] as { data: Record<string, { data: { spec: ChartSpec } }>; order: string[] };
  return sheet.order.map((id) => sheet.data[id]).filter((d) => d.data?.spec).map((d) => d.data.spec);
}

describe('graphiques chartEx (cascade, entonnoir, histogramme, boîte à moustaches)', () => {
  it('écrit des parties chartEx et un dessin avec remplacement', async () => {
    const bytes = await exportXlsx(chartWorkbook([COLUMN, WATERFALL, FUNNEL, HISTOGRAM, BOX]));
    const zip = await JSZip.loadAsync(bytes);
    const types = await zip.file('[Content_Types].xml')!.async('string');
    expect(types).toContain('/xl/charts/chartEx1.xml" ContentType="application/vnd.ms-office.chartex+xml"');
    const rels = await zip.file('xl/drawings/_rels/drawing1.xml.rels')!.async('string');
    expect(rels).toContain('relationships/chartEx" Target="../charts/chartEx4.xml"');
    const drawing = await zip.file('xl/drawings/drawing1.xml')!.async('string');
    expect(drawing.match(/<mc:AlternateContent/g)).toHaveLength(4);
    expect(drawing).not.toContain('Requires="cx2"'); // entonnoir : Excel ne l'affiche qu'avec cx1
    const waterfall = await zip.file('xl/charts/chartEx1.xml')!.async('string');
    expect(waterfall).toContain('layoutId="waterfall"');
    expect(waterfall).toContain('<cx:subtotals><cx:idx val="2"/></cx:subtotals>');
    expect(waterfall).toContain('<cx:pt idx="1">150</cx:pt>');
    const histogram = await zip.file('xl/charts/chartEx3.xml')!.async('string');
    expect(histogram).toContain('<cx:binning intervalClosed="r" overflow="140"><cx:binSize val="20"/></cx:binning>');
  });

  it('relie à chaque graphique un style et une palette (sinon Excel supprime le dessin)', async () => {
    const zip = await JSZip.loadAsync(await exportXlsx(chartWorkbook([WATERFALL, FUNNEL])));
    const types = await zip.file('[Content_Types].xml')!.async('string');
    for (const n of [1, 2]) {
      expect(types).toContain(`/xl/charts/style${n}.xml" ContentType="application/vnd.ms-office.chartstyle+xml"`);
      expect(types).toContain(`/xl/charts/colors${n}.xml" ContentType="application/vnd.ms-office.chartcolorstyle+xml"`);
      const rels = await zip.file(`xl/charts/_rels/chartEx${n}.xml.rels`)!.async('string');
      expect(rels).toContain(`relationships/chartStyle" Target="style${n}.xml"`);
      expect(rels).toContain(`relationships/chartColorStyle" Target="colors${n}.xml"`);
      const style = await zip.file(`xl/charts/style${n}.xml`)!.async('string');
      expect(style).toMatch(/<cs:dataPoint>.*<a:solidFill><a:schemeClr val="phClr"\/><\/a:solidFill>/);
      expect(await zip.file(`xl/charts/colors${n}.xml`)!.async('string')).toContain('<a:schemeClr val="accent6"/>');
    }
  });

  it('relit les graphiques à l’identique', async () => {
    const back = chartsOf(await importXlsx(await exportXlsx(chartWorkbook([WATERFALL, FUNNEL, HISTOGRAM, BOX]))));
    expect(back.map((s) => s.type)).toEqual(['waterfall', 'funnel', 'histogram', 'boxWhisker']);
    expect(back[0]).toMatchObject({ title: 'Résultat', waterfallTotals: [2], series: [{ nameRef: WATERFALL.series[0].nameRef, categoriesRef: WATERFALL.series[0].categoriesRef, valuesRef: WATERFALL.series[0].valuesRef }] });
    expect(back[1].series[0].color).toBe('#70AD47');
    expect(back[2].bins).toEqual({ width: 20, overflow: 140 });
    expect(back[3]).toMatchObject({ boxShowOutliers: false, legend: 'top' });
    expect(back[3].series).toHaveLength(2);
  });
});
