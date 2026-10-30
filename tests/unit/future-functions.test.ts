import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { exportXlsx } from '../../src/io/xlsx/export';
import { fromFileFormula, needsArrayFlag, toFileFormula } from '../../src/io/xlsx/future-functions';
import { importXlsx } from '../../src/io/xlsx/import';

describe('fonctions récentes dans le .xlsx', () => {
  it('préfixes à l’écriture', () => {
    expect(toFileFormula('IFS(A1>1,"IFS(",TRUE,2)')).toBe('_xlfn.IFS(A1>1,"IFS(",TRUE,2)');
    expect(toFileFormula('SUM(STDEV.S(A1:A2),SORT(B1:B2))')).toBe('SUM(_xlfn.STDEV.S(A1:A2),_xlfn._xlws.SORT(B1:B2))');
    expect(toFileFormula('FILTER(A1:A3,A1:A3>1)')).toBe('_xlfn._xlws.FILTER(A1:A3,A1:A3>1)');
    expect(toFileFormula("SUM('IFS(x'!A1)")).toBe("SUM('IFS(x'!A1)");
    expect(toFileFormula('SUM(A1,VLOOKUP(1,B1:C2,2,0))')).toBe('SUM(A1,VLOOKUP(1,B1:C2,2,0))');
  });
  it('paramètres de LET et LAMBDA', () => {
    expect(toFileFormula('LET(x,2,y,x+1,x*y)')).toBe('_xlfn.LET(_xlpm.x,2,_xlpm.y,_xlpm.x+1,_xlpm.x*_xlpm.y)');
    expect(toFileFormula('LAMBDA(a,a+1)(2)')).toBe('_xlfn.LAMBDA(_xlpm.a,_xlpm.a+1)(2)');
  });
  it('retirés à la lecture, aller-retour', () => {
    expect(fromFileFormula('_xlfn._xlws.SORT(A1:A2)+_xlfn.LET(_xlpm.x,2,_xlpm.x)')).toBe('SORT(A1:A2)+LET(x,2,x)');
    for (const f of ['IFS(A1>1,"a",TRUE,"b")', 'LET(x,2,x*3)', 'SUM(SORT(A1:A3))']) expect(fromFileFormula(toFileFormula(f))).toBe(f);
  });
  it('formules à marquer matricielles', () => {
    for (const f of ['SORT(A1:A3)', 'A1:A3*2', 'SUM((B1:B7="Lyon")*C1:C7)', 'MAX(IF(B1:B7="x",C1:C7))', 'SUM(Feuil!A1:A3*2)', 'CHOOSE({1,2},A1:A2,B1:B2)', 'TRANSPOSE(A1:A3)']) expect(needsArrayFlag(f), f).toBe(true);
    for (const f of ['SUM(A1:A3)', 'SUM(A1:A3)/2', 'A1*2', 'COUNTIF(A1:A3,">"&2)', 'IF(A1="a:b","x","y")', 'SUMIF(A1:A3,"A1:A3*2")']) expect(needsArrayFlag(f), f).toBe(false);
  });
  const book = {
    id: 'w', name: 'w', appVersion: '1', locale: 'frFR', styles: {}, sheetOrder: ['s'],
    sheets: { s: { id: 's', name: 'Feuil1', cellData: { 0: { 0: { f: '=IFS(1>2,"a",TRUE,"b")' }, 1: { f: '=SUM((C1:C2>0)*C1:C2)' }, 2: { v: 1, t: 2 } }, 1: { 2: { v: 2, t: 2 } } } } },
  } as never;
  it('dans le fichier : préfixe, marque matricielle, métadonnées, recalcul à l’ouverture', async () => {
    const zip = await JSZip.loadAsync(await exportXlsx(book));
    const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
    expect(sheet).toContain('<f>_xlfn.IFS(1&gt;2,"a",TRUE,"b")</f>');
    expect(sheet).toMatch(/<c r="B1"[^>]* cm="1"><f t="array" ref="B1">SUM\(\(C1:C2&gt;0\)\*C1:C2\)<\/f>/);
    expect(await zip.file('xl/metadata.xml')!.async('string')).toContain('XLDAPR');
    expect(await zip.file('[Content_Types].xml')!.async('string')).toContain('/xl/metadata.xml');
    expect(await zip.file('xl/_rels/workbook.xml.rels')!.async('string')).toContain('relationships/sheetMetadata');
    expect(await zip.file('xl/workbook.xml')!.async('string')).toContain('fullCalcOnLoad="1"');
  });
  it('relu sans préfixe (fichier d’Excel)', async () => {
    const back = await importXlsx(await exportXlsx(book));
    const cells = Object.values(back.sheets)[0]!.cellData!;
    expect(cells[0]![0]!.f).toBe('=IFS(1>2,"a",TRUE,"b")');
    expect(cells[0]![1]!.f).toBe('=SUM((C1:C2>0)*C1:C2)');
  });
});
