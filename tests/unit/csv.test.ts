import { describe, expect, it } from 'vitest';
import { csvToWorkbook, parseCsvValue, sheetToCsv } from '../../src/io/csv';
import { CellValueType } from '../../src/model/univer-types';

describe('CSV', () => {
  it('détecte le séparateur « ; » et la virgule décimale', () => {
    const wb = csvToWorkbook('﻿Nom;Prix;Qté\nCafé;1,5;2\n"Thé; vert";2,25;10\n', 'achats.csv');
    const s = wb.sheets[wb.sheetOrder[0]]!;
    expect(s.name).toBe('achats');
    expect(s.cellData?.[1]?.[0]).toMatchObject({ v: 'Café', t: CellValueType.STRING });
    expect(s.cellData?.[1]?.[1]).toMatchObject({ v: 1.5, t: CellValueType.NUMBER });
    expect(s.cellData?.[2]?.[0]?.v).toBe('Thé; vert');
    expect(s.cellData?.[2]?.[1]?.v).toBe(2.25);
  });

  it('lit le CSV anglo-saxon (« , » et point décimal)', () => {
    const wb = csvToWorkbook('a,b\n1.5,"x, y"\n', 'en.csv');
    const s = wb.sheets[wb.sheetOrder[0]]!;
    expect(s.cellData?.[1]?.[0]?.v).toBe(1.5);
    expect(s.cellData?.[1]?.[1]?.v).toBe('x, y');
  });

  it('type les valeurs : dates, pourcentages, booléens, zéros initiaux', () => {
    expect(parseCsvValue('14/10/2026', ',')).toMatchObject({ v: 46309, s: { n: { pattern: 'dd/mm/yyyy' } } });
    expect(parseCsvValue('2026-10-14', '.')).toMatchObject({ v: 46309 });
    expect(parseCsvValue('12,5 %', ',')).toMatchObject({ v: 0.125 });
    expect(parseCsvValue('1 234,5', ',')).toMatchObject({ v: 1234.5 });
    expect(parseCsvValue('VRAI', ',')).toMatchObject({ v: 1, t: CellValueType.BOOLEAN });
    expect(parseCsvValue('00123', ',')).toMatchObject({ v: '00123', t: CellValueType.STRING });
    expect(parseCsvValue('31/02/2026', ',')).toMatchObject({ t: CellValueType.STRING });
    expect(parseCsvValue('=SUM(A1:A2)', ',')).toEqual({ f: '=SUM(A1:A2)' });
    expect(parseCsvValue('   ', ',')).toBeNull();
  });

  it('exporte avec « ; », virgule décimale, BOM, guillemets si nécessaire', () => {
    const csv = sheetToCsv({
      cellData: {
        0: { 0: { v: 'Nom' }, 1: { v: 'Prix' } },
        1: { 0: { v: 'Thé; vert', t: CellValueType.STRING }, 1: { v: 2.25, t: CellValueType.NUMBER } },
        2: { 0: { v: 'Actif' }, 1: { v: 1, t: CellValueType.BOOLEAN } },
      },
    });
    expect(csv).toBe('﻿Nom;Prix\r\n"Thé; vert";2,25\r\nActif;VRAI');
  });

  it('round-trip CSV → feuille → CSV', () => {
    const src = 'Ville;Valeur\r\nLyon;1,5\r\nNîmes;-3\r\nQuébec;1000000';
    const wb = csvToWorkbook(src, 'v.csv');
    expect(sheetToCsv(wb.sheets[wb.sheetOrder[0]]!, { bom: false })).toBe(src);
  });
});
