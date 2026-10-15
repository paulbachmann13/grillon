import { describe, expect, it } from 'vitest';
import { exportXlsx } from '../../src/io/xlsx/export';
import { importXlsx } from '../../src/io/xlsx/import';
import { chartWorkbook } from '../fixtures/charts';
import {
  fromExcelPageSetup,
  hfFromExcel,
  hfToCssContent,
  hfToExcel,
  marginPresetOf,
  MARGIN_PRESETS,
  parseRowRange,
  printableSizePx,
  printZoom,
  toExcelPageSetup,
  type PageSetup,
} from '../../src/data/page-setup';

describe('mise en page', () => {
  it('en-têtes et pieds : aller-retour avec la syntaxe d’Excel', () => {
    const hf = { left: 'Rapport', center: 'Page &P / &N', right: '&D' };
    expect(hfToExcel(hf)).toBe('&LRapport&CPage &P / &N&R&D');
    expect(hfFromExcel(hfToExcel(hf))).toEqual(hf);
    expect(hfFromExcel('Titre seul')).toEqual({ center: 'Titre seul' });
    expect(hfToExcel({})).toBeUndefined();
  });

  it('codes d’en-tête → contenu CSS des marges de page', () => {
    const ctx = { file: 'budget.xlsx', sheet: 'Janvier', date: '15/10/2026', time: '10:30' };
    expect(hfToCssContent('Page &P / &N', ctx)).toBe('"Page " counter(page) " / " counter(pages)');
    expect(hfToCssContent('&F — &A', ctx)).toBe('"budget.xlsx — Janvier"');
    expect(hfToCssContent('&"Arial,Gras"&12Total && co', ctx)).toBe('"Total & co"');
    expect(hfToCssContent('', ctx)).toBe('""');
  });

  it('réglages ⇄ ExcelJS', () => {
    const page: PageSetup = {
      orientation: 'landscape',
      paper: 'A3',
      margins: MARGIN_PRESETS.narrow,
      fitWidth: 1,
      fitHeight: 0,
      centerH: true,
      printArea: 'A1:F40',
      titleRows: '1:2',
      gridlines: true,
      header: { center: 'Budget' },
      footer: { right: 'Page &P' },
    };
    const x = toExcelPageSetup(page);
    expect(x.pageSetup).toMatchObject({ paperSize: 8, orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0, printArea: 'A1:F40', printTitlesRow: '1:2', horizontalCentered: true, showGridLines: true });
    expect(x.oddHeader).toBe('&CBudget');
    expect(fromExcelPageSetup(x.pageSetup, x.oddHeader, x.oddFooter)).toEqual(page);
  });

  it('une mise en page par défaut ne laisse rien', () => {
    expect(fromExcelPageSetup({ paperSize: 9, orientation: 'portrait', margins: MARGIN_PRESETS.normal })).toBeUndefined();
    expect(fromExcelPageSetup({ scale: 80 })).toEqual({ scale: 80 });
  });

  it('marges prédéfinies, lignes à répéter, taille imprimable et échelle', () => {
    expect(marginPresetOf(MARGIN_PRESETS.wide)).toBe('wide');
    expect(marginPresetOf({ ...MARGIN_PRESETS.normal, top: 2 })).toBe('custom');
    expect(parseRowRange('$1:$3')).toEqual({ first: 0, last: 2 });
    expect(parseRowRange('3:1')).toBeNull();
    const a4 = printableSizePx({});
    expect(Math.round(a4.width)).toBe(Math.round((210 / 25.4 - 1.4) * 96));
    expect(printableSizePx({ orientation: 'landscape' }).width).toBeGreaterThan(a4.width);
    expect(printZoom({ scale: 50 }, 3000)).toBe(0.5);
    expect(printZoom({ fitWidth: 1 }, a4.width * 2)).toBeCloseTo(0.5, 5);
    expect(printZoom({ fitWidth: 1 }, 100)).toBe(1);
  });

  it("round-trip .xlsx : la mise en page est écrite dans le fichier puis relue", async () => {
    const page: PageSetup = { orientation: "landscape", paper: "A3", margins: MARGIN_PRESETS.wide, fitWidth: 1, fitHeight: 0, centerH: true, printArea: "A1:C4", titleRows: "1:1", gridlines: true, header: { left: "Budget", right: "&D" }, footer: { center: "Page &P / &N" } };
    const wb = chartWorkbook([]);
    (wb.sheets.s1 as { custom?: unknown }).custom = { grillon: { page } };
    const bytes = await exportXlsx(wb);
    const JSZip = (await import("jszip")).default;
    const zip = await JSZip.loadAsync(bytes);
    const xml = await zip.file("xl/worksheets/sheet1.xml")!.async("string");
    expect(xml).toMatch(/<pageSetup[^>]*paperSize="8"[^>]*orientation="landscape"/);
    expect(xml).toContain("<oddHeader>&amp;LBudget&amp;R&amp;D</oddHeader>");
    const back = await importXlsx(bytes);
    const got = (Object.values(back.sheets)[0] as { custom?: { grillon?: { page?: PageSetup } } }).custom?.grillon?.page;
    expect(got).toMatchObject({ orientation: "landscape", paper: "A3", fitWidth: 1, printArea: "A1:C4", titleRows: "1:1", centerH: true, gridlines: true });
    expect(got?.header).toEqual({ left: "Budget", right: "&D" });
    expect(got?.footer).toEqual({ center: "Page &P / &N" });
  });
});
