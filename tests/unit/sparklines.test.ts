import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { sparkTargets, sparklineOption, type SparklineSpec } from '../../src/charts/sparkline';
import { exportXlsx } from '../../src/io/xlsx/export';
import { importXlsx } from '../../src/io/xlsx/import';
import { SPARKLINE_COMPONENT_KEY, readSparklines, sparklinesOfSheet, writeSparklines } from '../../src/io/xlsx/sparklines';
import { chartWorkbook } from '../fixtures/charts';

const ref = (r: { r1: number; c1: number; r2: number; c2: number }) => `Feuil1!R${r.r1}C${r.c1}:R${r.r2}C${r.c2}`;

describe('mini-graphiques', () => {
  it('répartit les lignes de données sur les cellules d’un emplacement en colonne', () => {
    const t = sparkTargets({ r1: 1, c1: 1, r2: 3, c2: 5 }, { r1: 1, c1: 6, r2: 3, c2: 6 }, ref);
    expect(t).toEqual([
      { row: 1, col: 6, dataRef: 'Feuil1!R1C1:R1C5' },
      { row: 2, col: 6, dataRef: 'Feuil1!R2C1:R2C5' },
      { row: 3, col: 6, dataRef: 'Feuil1!R3C1:R3C5' },
    ]);
  });

  it('emplacement en ligne : une colonne de données par cellule ; sinon refus', () => {
    const t = sparkTargets({ r1: 1, c1: 1, r2: 4, c2: 2 }, { r1: 5, c1: 1, r2: 5, c2: 2 }, ref);
    expect(t?.map((x) => x.dataRef)).toEqual(['Feuil1!R1C1:R4C1', 'Feuil1!R1C2:R4C2']);
    expect(sparkTargets({ r1: 1, c1: 1, r2: 3, c2: 5 }, { r1: 1, c1: 6, r2: 2, c2: 6 }, ref)).toBeNull();
    expect(sparkTargets({ r1: 1, c1: 1, r2: 1, c2: 5 }, { r1: 1, c1: 6, r2: 1, c2: 6 }, ref)).toHaveLength(1);
  });

  it('options ECharts : courbe, histogramme, positif/négatif, points en évidence', () => {
    const values = [3, -2, 5, null, 1];
    const line = sparklineOption({ type: 'line', dataRef: 'x', high: true, low: true }, values) as { series: { type: string; data: unknown[] }[] };
    expect(line.series[0]!.type).toBe('line');
    expect(line.series[0]!.data.filter((d) => typeof d === 'object' && d !== null)).toHaveLength(2);
    const wl = sparklineOption({ type: 'winloss', dataRef: 'x' }, values) as { series: { data: ({ value: number } | null)[] }[] };
    expect(wl.series[0]!.data.map((d) => d?.value ?? null)).toEqual([1, -1, 1, null, 1]);
    const col = sparklineOption({ type: 'column', dataRef: 'x', negative: true, highlight: '#FF0000' }, values) as { series: { data: ({ itemStyle: { color: string } } | null)[] }[] };
    expect(col.series[0]!.data[1]!.itemStyle.color).toBe('#FF0000');
  });

  it('xlsx : écriture x14 puis relecture', () => {
    const base = '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>';
    const specs: SparklineSpec[] = [
      { type: 'line', dataRef: 'Feuil1!A2:E2', markers: true, high: true, color: '#00AA00' },
      { type: 'line', dataRef: 'Feuil1!A3:E3', markers: true, high: true, color: '#00AA00' },
      { type: 'winloss', dataRef: 'Feuil1!A4:E4', negative: true },
      { type: 'column', dataRef: 'Feuil1!A5:E5', min: 0, max: 10, emptyCells: 'zero' },
    ];
    const xml = writeSparklines(base, specs.map((spec, i) => ({ spec, row: 1 + i, col: 5 })));
    expect(xml).toContain('<x14:sparklineGroups');
    expect(xml.match(/<x14:sparklineGroup\b/g)).toHaveLength(3); // les deux courbes identiques sont groupées
    const back = readSparklines(xml);
    expect(back.map((b) => [b.row, b.col, b.spec.type, b.spec.dataRef])).toEqual([
      [1, 5, 'line', 'Feuil1!A2:E2'],
      [2, 5, 'line', 'Feuil1!A3:E3'],
      [3, 5, 'winloss', 'Feuil1!A4:E4'],
      [4, 5, 'column', 'Feuil1!A5:E5'],
    ]);
    expect(back[0]!.spec).toMatchObject({ markers: true, high: true, color: '#00AA00' });
    expect(back[3]!.spec).toMatchObject({ min: 0, max: 10, emptyCells: 'zero' });
  });

  it('round-trip .xlsx complet : le fichier contient l’extension et on retrouve les mini-graphiques', async () => {
    const wb = chartWorkbook([]);
    const item = (id: string, row: number, spec: SparklineSpec) => ({
      drawingId: id,
      drawingType: 8,
      componentKey: SPARKLINE_COMPONENT_KEY,
      sheetTransform: { from: { column: 6, columnOffset: 0, row, rowOffset: 0 }, to: { column: 7, columnOffset: 0, row: row + 1, rowOffset: 0 } },
      transform: { left: 0, top: 0, width: 88, height: 24 },
      data: { sparkId: id, spec },
    });
    const data = {
      a: item('a', 1, { type: 'column', dataRef: "'Ventes 2026'!$B$2:$D$2", color: '#2E75B6' }),
      b: item('b', 2, { type: 'line', dataRef: "'Ventes 2026'!$B$3:$D$3", markers: true, last: true }),
    };
    wb.resources = [{ name: 'SHEET_DRAWING_PLUGIN', data: JSON.stringify({ s1: { data, order: ['a', 'b'] } }) }];
    const bytes = await exportXlsx(wb);
    const sheet = await (await JSZip.loadAsync(bytes)).file('xl/worksheets/sheet1.xml')!.async('string');
    expect(sheet).toContain('<x14:sparklineGroups');
    const back = await importXlsx(bytes);
    const res = JSON.parse(back.resources!.find((r) => r.name === 'SHEET_DRAWING_PLUGIN')!.data);
    const found = sparklinesOfSheet(res, Object.keys(res)[0]!);
    expect(found.map((f) => [f.row, f.col, f.spec.type])).toEqual([
      [1, 6, 'column'],
      [2, 6, 'line'],
    ]);
    expect(found[0]!.spec.color).toBe('#2E75B6');
    expect(found[1]!.spec).toMatchObject({ markers: true, last: true });
  });
});
