import JSZip from 'jszip';
import { exportXlsx } from '../../src/io/xlsx/export';
import { importXlsx } from '../../src/io/xlsx/import';
import { describe, expect, it } from 'vitest';
import { colBreaksXml, hasBreakAt, hasBreaks, insertBreaks, normalizeBreaks, pageGrid, paginate, readBreaks, removeBreaks, rowBreaksXml, withBreaks } from '../../src/data/page-breaks';

describe('sauts de page manuels', () => {
  it('normalise : trié, sans doublon ni valeur invalide', () => {
    expect(normalizeBreaks([20, 5, 5, 0, -3, 1.5, 10])).toEqual([5, 10, 20]);
    expect(normalizeBreaks(undefined)).toEqual([]);
  });

  it('insérer sur une cellule : saut au-dessus et à gauche (comme Excel)', () => {
    expect(insertBreaks({}, { startRow: 9, startCol: 3 })).toEqual({ rowBreaks: [9], colBreaks: [3] });
    expect(insertBreaks({ rowBreaks: [30] }, { startRow: 9, startCol: 0 })).toEqual({ rowBreaks: [9, 30], colBreaks: [] });
  });

  it('insérer sur une ligne entière ou une colonne entière', () => {
    expect(insertBreaks({}, { startRow: 9, startCol: 0, wholeRow: true })).toEqual({ rowBreaks: [9], colBreaks: [] });
    expect(insertBreaks({}, { startRow: 0, startCol: 4, wholeCol: true })).toEqual({ rowBreaks: [], colBreaks: [4] });
  });

  it('un saut avant la première ligne ou colonne est ignoré', () => {
    expect(insertBreaks({}, { startRow: 0, startCol: 0 })).toEqual({ rowBreaks: [], colBreaks: [] });
  });

  it('supprimer le saut de la cellule active', () => {
    const b = { rowBreaks: [9, 30], colBreaks: [3] };
    expect(removeBreaks(b, { startRow: 9, startCol: 3 })).toEqual({ rowBreaks: [30], colBreaks: [] });
    expect(removeBreaks(b, { startRow: 30, startCol: 0, wholeRow: true })).toEqual({ rowBreaks: [9], colBreaks: [3] });
    expect(hasBreakAt(b, { startRow: 9, startCol: 0 })).toBe(true);
    expect(hasBreakAt(b, { startRow: 8, startCol: 2 })).toBe(false);
    expect(hasBreaks(b)).toBe(true);
    expect(hasBreaks({ rowBreaks: [] })).toBe(false);
  });

  it('withBreaks garde la mise en page sans champs vides', () => {
    expect(withBreaks({ scale: 80, rowBreaks: [4] }, { rowBreaks: [], colBreaks: [2] })).toEqual({ scale: 80, colBreaks: [2] });
  });
});

describe('découpage en pages', () => {
  const fixed = (n: number) => () => n;

  it('saut automatique quand la page est pleine', () => {
    // 10 lignes de 20 px, pages de 100 px : 5 lignes par page.
    expect(paginate(fixed(20), 0, 9, 100)).toEqual([
      { start: 0, end: 4, manual: false },
      { start: 5, end: 9, manual: false },
    ]);
  });

  it('saut manuel : nouvelle page, puis le décompte repart de zéro', () => {
    expect(paginate(fixed(20), 0, 9, 100, [3])).toEqual([
      { start: 0, end: 2, manual: false },
      { start: 3, end: 7, manual: true },
      { start: 8, end: 9, manual: false },
    ]);
  });

  it('les lignes masquées (taille 0) ne comptent pas, un élément trop grand occupe sa page', () => {
    const sizes = [20, 0, 0, 20, 500, 20];
    expect(paginate((i) => sizes[i], 0, 5, 100)).toEqual([
      { start: 0, end: 3, manual: false },
      { start: 4, end: 4, manual: false },
      { start: 5, end: 5, manual: false },
    ]);
  });

  it('ignore les sauts hors de la zone', () => {
    expect(paginate(fixed(10), 5, 8, 1000, [2, 5, 20])).toEqual([{ start: 5, end: 8, manual: false }]);
    expect(paginate(fixed(10), 3, 2, 100)).toEqual([]);
  });

  it('numérote les pages de haut en bas puis de gauche à droite', () => {
    const rows = paginate(fixed(10), 0, 3, 20);
    const cols = paginate(fixed(10), 0, 3, 20);
    const grid = pageGrid(rows, cols);
    expect(grid).toHaveLength(4);
    expect(grid.map((p) => [p.n, p.rows.start, p.cols.start])).toEqual([
      [1, 0, 0],
      [2, 2, 0],
      [3, 0, 2],
      [4, 2, 2],
    ]);
  });
});

describe('<rowBreaks> et <colBreaks>', () => {
  it('écrit les sauts comme Excel', () => {
    expect(rowBreaksXml([9, 30])).toBe('<rowBreaks count="2" manualBreakCount="2"><brk id="9" max="16383" man="1"/><brk id="30" max="16383" man="1"/></rowBreaks>');
    expect(colBreaksXml([3])).toBe('<colBreaks count="1" manualBreakCount="1"><brk id="3" max="1048575" man="1"/></colBreaks>');
    expect(rowBreaksXml([])).toBe('');
  });

  it('relit les sauts manuels et ignore les sauts automatiques', () => {
    const xml = '<worksheet><rowBreaks count="3" manualBreakCount="2"><brk id="9" max="16383" man="1"/><brk id="20" max="16383"/><brk id="30" max="16383" man="1"/></rowBreaks><colBreaks count="1"><brk id="4" max="1048575" man="1"/></colBreaks></worksheet>';
    expect(readBreaks(xml, 'rowBreaks')).toEqual([9, 30]);
    expect(readBreaks(xml, 'colBreaks')).toEqual([4]);
    expect(readBreaks('<worksheet/>', 'rowBreaks')).toEqual([]);
  });
});

describe('sauts de page dans le .xlsx', () => {
  const book = (page: object) =>
    ({
      id: 'wb',
      name: 'T',
      appVersion: '1',
      locale: 'frFR',
      styles: {},
      sheetOrder: ['s'],
      sheets: { s: { id: 's', name: 'Feuil1', cellData: { 0: { 0: { v: 'x' } } }, rowCount: 50, columnCount: 10, custom: { grillon: { page } } } },
    }) as never;

  it('aller-retour : sauts de ligne et de colonne, avec le reste de la mise en page', async () => {
    const bytes = await exportXlsx(book({ orientation: 'landscape', rowBreaks: [9, 30], colBreaks: [3] }));
    const zip = await JSZip.loadAsync(bytes);
    const xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
    expect(xml).toContain('<rowBreaks count="2" manualBreakCount="2"><brk id="9" max="16383" man="1"/><brk id="30" max="16383" man="1"/></rowBreaks>');
    expect(xml).toContain('<colBreaks count="1" manualBreakCount="1"><brk id="3" max="1048575" man="1"/></colBreaks>');
    // Ordre du schéma : pageSetup, puis rowBreaks, puis colBreaks.
    expect(xml.indexOf('<pageSetup')).toBeLessThan(xml.indexOf('<rowBreaks'));
    expect(xml.indexOf('<rowBreaks')).toBeLessThan(xml.indexOf('<colBreaks'));
    const back = await importXlsx(bytes);
    const page = (Object.values(back.sheets)[0]!.custom as { grillon: { page: Record<string, unknown> } }).grillon.page;
    expect(page).toMatchObject({ orientation: 'landscape', rowBreaks: [9, 30], colBreaks: [3] });
  });

  it('sans saut, rien n’est écrit', async () => {
    const zip = await JSZip.loadAsync(await exportXlsx(book({ orientation: 'landscape' })));
    const xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
    expect(xml).not.toContain('Breaks');
  });
});

import { computeLayout } from '../../src/data/page-layout';
import { printableSizePx } from '../../src/data/page-setup';

describe('mise en page des pages', () => {
  const rowH = 20;
  const colW = 64;
  const base = { rows: { from: 0, to: 199 }, cols: { from: 0, to: 9 }, rowSize: () => rowH, colSize: () => colW };

  it('pages automatiques selon le papier et les marges', () => {
    const { height, width } = printableSizePx({});
    const perPage = Math.floor(height / rowH);
    const layout = computeLayout({ ...base, page: {} });
    expect(layout.rowPages[0]).toEqual({ start: 0, end: perPage - 1, manual: false });
    expect(layout.colPages).toHaveLength(Math.ceil((10 * colW) / width));
    expect(layout.pages.map((p) => p.n)).toEqual(layout.pages.map((_, i) => i + 1));
  });

  it('un saut manuel démarre une nouvelle page, et le paysage change la découpe', () => {
    const portrait = computeLayout({ ...base, page: { rowBreaks: [10] } });
    expect(portrait.rowPages[1]).toMatchObject({ start: 10, manual: true });
    const landscape = computeLayout({ ...base, page: { orientation: 'landscape' } });
    expect(landscape.rowPages.length).toBeGreaterThan(portrait.rowPages.length);
  });

  it('ajusté à 1 page de large : sauts manuels ignorés, échelle réduite', () => {
    const wide = { ...base, cols: { from: 0, to: 39 } };
    const layout = computeLayout({ ...wide, page: { fitWidth: 1, colBreaks: [5], rowBreaks: [3] } });
    expect(layout.manualIgnored).toBe(true);
    expect(layout.zoom).toBeLessThan(1);
    expect(layout.colPages).toHaveLength(1);
    expect(layout.rowPages.every((p) => !p.manual)).toBe(true);
  });

  it('les lignes de titre répétées réduisent la place de chaque page', () => {
    const without = computeLayout({ ...base, page: {} });
    const withTitles = computeLayout({ ...base, page: {}, repeatedHeight: 100 });
    expect(withTitles.rowPages[0].end).toBeLessThan(without.rowPages[0].end);
  });
});
