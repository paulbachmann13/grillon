import { describe, expect, it } from 'vitest';
import { parseFrenchEntry } from '../../src/data/french-entry';
import { serialToYmd } from '../../src/data/slicer';

describe('saisie à la française', () => {
  it('espace des milliers (normale, insécable, fine)', () => {
    expect(parseFrenchEntry('1 234,5')).toEqual({ kind: 'number', value: 1234.5, pattern: '#,##0.0' });
    expect(parseFrenchEntry('1 000 000')).toEqual({ kind: 'number', value: 1000000, pattern: '#,##0' });
    expect(parseFrenchEntry('-1 234')).toEqual({ kind: 'number', value: -1234, pattern: '#,##0' });
    expect(parseFrenchEntry('1 234,56')).toEqual({ kind: 'number', value: 1234.56, pattern: '#,##0.00' });
    expect(parseFrenchEntry('1 234,56')).toEqual({ kind: 'number', value: 1234.56, pattern: '#,##0.00' });
    expect(parseFrenchEntry('  2 500  ')).toEqual({ kind: 'number', value: 2500, pattern: '#,##0' });
  });
  it('les groupes mal formés restent du texte', () => {
    expect(parseFrenchEntry('12 34')).toBeNull();
    expect(parseFrenchEntry('1 2345')).toBeNull();
    expect(parseFrenchEntry('rue 12 345')).toBeNull();
  });
  it('montants en euros', () => {
    expect(parseFrenchEntry('-12,50 €')).toEqual({ kind: 'number', value: -12.5, pattern: '#,##0.00 €' });
    expect(parseFrenchEntry('1 250 €')).toEqual({ kind: 'number', value: 1250, pattern: '#,##0 €' });
    expect(parseFrenchEntry('3,5€')).toEqual({ kind: 'number', value: 3.5, pattern: '#,##0.0 €' });
    // Déjà bien lu par Univer : on n'y touche pas.
    expect(parseFrenchEntry('25 €')).toBeNull();
  });
  it('point décimal : 1.5 vaut 1,5 (et pas 15)', () => {
    expect(parseFrenchEntry('1.5')).toEqual({ kind: 'number', value: 1.5 });
    expect(parseFrenchEntry('-0.25')).toEqual({ kind: 'number', value: -0.25 });
    expect(parseFrenchEntry('1.2.3')).toBeNull();
  });
  it('date sans année : jour/mois de l’année en cours', () => {
    const d = (text: string) => {
      const r = parseFrenchEntry(text, 2026);
      return r?.kind === 'date' ? serialToYmd(r.value) : r;
    };
    expect(d('3/4')).toEqual({ y: 2026, m: 4, d: 3 });
    expect(d('31/12')).toEqual({ y: 2026, m: 12, d: 31 });
    expect(d('13/1')).toEqual({ y: 2026, m: 1, d: 13 });
    expect(d('3-4')).toEqual({ y: 2026, m: 4, d: 3 });
    expect(parseFrenchEntry('3/4', 2026)).toMatchObject({ pattern: 'dd/mm' });
  });
  it('date impossible : reste du texte', () => {
    expect(parseFrenchEntry('31/02', 2026)).toBeNull();
    expect(parseFrenchEntry('29/02', 2026)).toBeNull();
    expect(parseFrenchEntry('29/02', 2028)).toMatchObject({ kind: 'date' });
    expect(parseFrenchEntry('12/13')).toBeNull();
    expect(parseFrenchEntry('0/5')).toBeNull();
  });
  it('VRAI et FAUX, toutes casses', () => {
    expect(parseFrenchEntry('VRAI')).toEqual({ kind: 'boolean', value: true });
    expect(parseFrenchEntry('faux')).toEqual({ kind: 'boolean', value: false });
    expect(parseFrenchEntry('Vrai')).toEqual({ kind: 'boolean', value: true });
    expect(parseFrenchEntry('vraiment')).toBeNull();
  });
  it('le reste est laissé à Univer', () => {
    for (const t of ['1,5', '12/03/2024', '15%', '=1+1', "'007", 'Bonjour', '14:30', '', '42']) expect(parseFrenchEntry(t)).toBeNull();
  });
});
