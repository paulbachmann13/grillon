import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { exportXlsx } from '../../src/io/xlsx/export';
import { importXlsx } from '../../src/io/xlsx/import';

// Une formule recopiée vers le bas dans Univer : la première cellule garde « f », les suivantes seulement « si » (formule partagée).
const book = (cells: Record<number, Record<number, object>>) =>
  ({
    id: 'wb',
    name: 'T',
    appVersion: '1',
    locale: 'frFR',
    styles: {},
    sheetOrder: ['s'],
    sheets: { s: { id: 's', name: 'Feuil1', cellData: cells, rowCount: 50, columnCount: 10 } },
  }) as never;

const filled = book({
  0: { 0: { v: 1 }, 1: { f: '=A1*2', v: 2, t: 2 } },
  1: { 0: { v: 2 }, 1: { si: 'g1', f: '=A2*2', v: 4, t: 2 } },
  2: { 0: { v: 3 }, 1: { si: 'g1', v: 6, t: 2 } },
  3: { 0: { v: 4 }, 1: { si: 'g1', v: 8, t: 2 } },
  4: { 0: { v: 5 }, 1: { si: 'g1', v: 10, t: 2 } },
});

describe('formules recopiées (formules partagées d’Univer)', () => {
  it('chaque cellule garde sa formule dans le .xlsx, avec ses références décalées', async () => {
    const zip = await JSZip.loadAsync(await exportXlsx(filled));
    const xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
    const formulas = [...xml.matchAll(/<c r="(B\d)"[^>]*><f>([^<]*)<\/f><v>([^<]*)<\/v>/g)].map((m) => [m[1], m[2], m[3]]);
    expect(formulas).toEqual([
      ['B1', 'A1*2', '2'],
      ['B2', 'A2*2', '4'],
      ['B3', 'A3*2', '6'],
      ['B4', 'A4*2', '8'],
      ['B5', 'A5*2', '10'],
    ]);
  });

  it('aller-retour : les formules sont relues (pas des nombres)', async () => {
    const back = await importXlsx(await exportXlsx(filled));
    const cells = Object.values(back.sheets)[0]!.cellData as Record<number, Record<number, { f?: string; v?: unknown }>>;
    expect([0, 1, 2, 3, 4].map((r) => cells[r]?.[1]?.f)).toEqual(['=A1*2', '=A2*2', '=A3*2', '=A4*2', '=A5*2']);
    expect([0, 1, 2, 3, 4].map((r) => Number(cells[r]?.[1]?.v))).toEqual([2, 4, 6, 8, 10]);
  });

  it('remplissage vers la droite et en bloc, références absolues et autre feuille', async () => {
    const wb = book({
      0: { 0: { v: 10 }, 1: { v: 20 }, 2: { v: 30 }, 3: { v: 0.2 } },
      1: {
        0: { si: 'h', f: '=A1*$D$1+Feuil2!A1', v: 2 },
        1: { si: 'h', v: 4 },
        2: { si: 'h', v: 6 },
      },
      2: {
        0: { si: 'h', v: 1 },
        1: { si: 'h', v: 1 },
      },
    });
    const zip = await JSZip.loadAsync(await exportXlsx(wb));
    const xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
    const f = (ref: string) => new RegExp(`<c r="${ref}"[^>]*><f>([^<]*)</f>`).exec(xml)?.[1];
    expect(f('A2')).toBe('A1*$D$1+Feuil2!A1');
    expect(f('B2')).toBe('B1*$D$1+Feuil2!B1');
    expect(f('C2')).toBe('C1*$D$1+Feuil2!C1');
    expect(f('A3')).toBe('A2*$D$1+Feuil2!A2');
    expect(f('B3')).toBe('B2*$D$1+Feuil2!B2');
  });

  it('une cellule partagée dont la première est introuvable reste une valeur (rien n’est inventé)', async () => {
    const zip = await JSZip.loadAsync(await exportXlsx(book({ 0: { 0: { si: 'orpheline', v: 7 } } })));
    const xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
    expect(xml).toContain('<c r="A1"');
    expect(xml).not.toContain('<f>');
  });
});

describe('fichier écrit par Excel avec des formules partagées', () => {
  it('chaque cellule est relue avec sa propre formule', async () => {
    const zip = await JSZip.loadAsync(await exportXlsx(filled));
    let xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
    // Excel écrit une formule partagée ainsi : la première porte le texte et la plage, les autres seulement l'indice.
    xml = xml.replace(/<f>A2\*2<\/f>/, '<f t="shared" ref="B2:B5" si="0">A2*2</f>').replace(/<f>A[345]\*2<\/f>/g, '<f t="shared" si="0"/>');
    zip.file('xl/worksheets/sheet1.xml', xml);
    const back = await importXlsx(await zip.generateAsync({ type: 'uint8array' }));
    const cells = Object.values(back.sheets)[0]!.cellData as Record<number, Record<number, { f?: string }>>;
    expect([0, 1, 2, 3, 4].map((r) => cells[r]?.[1]?.f)).toEqual(['=A1*2', '=A2*2', '=A3*2', '=A4*2', '=A5*2']);
  });
});
