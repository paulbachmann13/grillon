import { describe, expect, it } from 'vitest';
import { exportXlsx } from '../../src/io/xlsx/export';
import { importXlsx } from '../../src/io/xlsx/import';

const book = (sheet: object, styles: object = {}) =>
  ({
    id: 'wb',
    name: 'T',
    appVersion: '1',
    locale: 'frFR',
    styles,
    sheetOrder: ['s'],
    sheets: { s: { id: 's', name: 'Feuil1', rowCount: 50, columnCount: 10, cellData: { 0: { 0: { v: 'titre' } }, 3: { 0: { v: 'données' } } }, ...sheet } },
  }) as never;
const roundTrip = async (wb: never) => Object.values((await importXlsx(await exportXlsx(wb))).sheets)[0] as Record<string, any>;

describe('fidélité du .xlsx (audit)', () => {
  it('« Figer la ligne supérieure » (Univer : startColumn = -1) reste figé', async () => {
    const back = await roundTrip(book({ freeze: { startRow: 1, startColumn: -1, ySplit: 1, xSplit: 0 } }));
    expect(back.freeze).toMatchObject({ ySplit: 1, startRow: 1, xSplit: 0 });
  });

  it('figer la première colonne seule, et lignes + colonnes', async () => {
    expect((await roundTrip(book({ freeze: { startRow: -1, startColumn: 1, ySplit: 0, xSplit: 1 } }))).freeze).toMatchObject({ xSplit: 1, startColumn: 1, ySplit: 0 });
    expect((await roundTrip(book({ freeze: { startRow: 2, startColumn: 1, ySplit: 2, xSplit: 1 } }))).freeze).toMatchObject({ xSplit: 1, ySplit: 2 });
  });

  it('une ligne masquée sans contenu reste masquée', async () => {
    const back = await roundTrip(book({ rowData: { 3: { hd: 1 }, 7: { hd: 1 } } }));
    expect(back.rowData?.[3]?.hd).toBe(1);
    expect(back.rowData?.[7]?.hd).toBe(1);
  });

  it('bordures dont le style est rangé en texte (« thin ») ou en code', async () => {
    const back = await importXlsx(
      await exportXlsx(
        book(
          { cellData: { 0: { 0: { v: 'a', s: 'texte' }, 1: { v: 'b', s: 'code' } } } },
          { texte: { bd: { t: { s: 'thin', cl: { rgb: '#000000' } }, b: { s: 'medium', cl: { rgb: '#FF0000' } } } }, code: { bd: { l: { s: 1, cl: { rgb: '#000000' } } } } },
        ),
      ),
    );
    const cells = Object.values(back.sheets)[0]!.cellData as Record<number, Record<number, { s?: string }>>;
    const style = (r: number, c: number) => back.styles[cells[r][c].s as string] as { bd?: Record<string, { s: number; cl: { rgb: string } }> };
    expect(Object.keys(style(0, 0).bd ?? {}).sort()).toEqual(['b', 't']);
    expect(style(0, 0).bd!.b.cl.rgb.toUpperCase()).toBe('#FF0000');
    expect(Object.keys(style(0, 1).bd ?? {})).toEqual(['l']);
  });
});
