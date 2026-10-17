import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { checkPassword, legacyHash, makeProtection, readSheetProtection, sheetProtectionXml, type SheetProtection } from '../../src/data/sheet-protection';
import { exportXlsx } from '../../src/io/xlsx/export';
import { importXlsx } from '../../src/io/xlsx/import';
import { chartWorkbook } from '../fixtures/charts';

describe('protection de feuille', () => {
  it('mot de passe : hachage façon Excel (SHA-512, sel, 100 000 tours) et vérification', async () => {
    const p = await makeProtection('Grillon2026');
    expect(p).toMatchObject({ algorithmName: 'SHA-512', spinCount: 100000 });
    expect(p.hashValue).toMatch(/^[A-Za-z0-9+/]{86}==$/);
    expect(await checkPassword(p, 'Grillon2026')).toBe(true);
    expect(await checkPassword(p, 'grillon2026')).toBe(false);
    expect(await checkPassword({}, 'nimporte')).toBe(true);
  }, 30_000);

  it('ancien hachage 16 bits des fichiers Excel anciens', async () => {
    const h = legacyHash('abc');
    expect(h).toMatch(/^[0-9A-F]{1,4}$/);
    expect(await checkPassword({ legacyPassword: h }, 'abc')).toBe(true);
    expect(await checkPassword({ legacyPassword: h }, 'abd')).toBe(false);
  });

  it('élément <sheetProtection> : écriture et lecture', () => {
    const p: SheetProtection = { algorithmName: 'SHA-512', hashValue: 'aGFzaA==', saltValue: 'c2Fs', spinCount: 100000 };
    const xml = sheetProtectionXml(p);
    expect(xml).toBe('<sheetProtection algorithmName="SHA-512" hashValue="aGFzaA==" saltValue="c2Fs" spinCount="100000" sheet="1" objects="1" scenarios="1"/>');
    expect(readSheetProtection(`<worksheet><sheetData/>${xml}</worksheet>`)).toEqual(p);
    expect(readSheetProtection('<worksheet><sheetProtection sheet="0"/></worksheet>')).toBeNull();
    expect(readSheetProtection('<worksheet><sheetProtection password="CBEB" sheet="1"/></worksheet>')).toEqual({ legacyPassword: 'CBEB' });
  });

  it('round-trip .xlsx : la feuille reste protégée, au bon endroit dans le XML', async () => {
    const wb = chartWorkbook([]);
    (wb.sheets.s1 as { custom?: unknown }).custom = { grillon: { protection: {} } };
    const bytes = await exportXlsx(wb);
    const sheet = await (await JSZip.loadAsync(bytes)).file('xl/worksheets/sheet1.xml')!.async('string');
    expect(sheet).toMatch(/<\/sheetData><sheetProtection sheet="1" objects="1" scenarios="1"\/>/);
    const back = await importXlsx(bytes);
    expect((Object.values(back.sheets)[0] as { custom?: { grillon?: { protection?: unknown } } }).custom?.grillon?.protection).toEqual({});
  });

  it('round-trip .xlsx : cellules déverrouillées (protection locked="0")', async () => {
    const wb = chartWorkbook([]);
    const rect = { startRow: 1, endRow: 3, startColumn: 1, endColumn: 2 };
    (wb.sheets.s1 as { custom?: unknown }).custom = { grillon: { protection: {}, unlocked: [rect] } };
    const bytes = await exportXlsx(wb);
    const styles = await (await JSZip.loadAsync(bytes)).file('xl/styles.xml')!.async('string');
    expect(styles).toContain('locked="0"');
    const back = await importXlsx(bytes);
    const grillon = (Object.values(back.sheets)[0] as { custom?: { grillon?: { unlocked?: unknown; protection?: unknown } } }).custom?.grillon;
    expect(grillon?.unlocked).toEqual([rect]);
    expect(grillon?.protection).toEqual({});
  });
});
