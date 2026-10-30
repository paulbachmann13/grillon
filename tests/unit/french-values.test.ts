import { describe, expect, it } from 'vitest';
import { amorDegrc, fixedText, frenchCriteria, frenchFormatToEnglish, frenchNumberText, parseFrenchDate, parseFrenchNumber, searchWithWildcards } from '../../src/data/french-values';

describe('valeurs à la française', () => {
  it('nombre → texte', () => {
    expect(frenchNumberText(1.5)).toBe('1,5');
    expect(frenchNumberText(1 / 3)).toBe('0,333333333333333');
    expect(frenchNumberText(0.1 + 0.2)).toBe('0,3');
    expect(frenchNumberText(1000)).toBe('1000');
    expect(frenchNumberText(-2.25)).toBe('-2,25');
  });
  it('texte → nombre', () => {
    expect(parseFrenchNumber('12,5')).toBe(12.5);
    expect(parseFrenchNumber('1 000')).toBe(1000);
    expect(parseFrenchNumber('-1 234,5')).toBe(-1234.5);
    expect(parseFrenchNumber('12,5 %')).toBeCloseTo(0.125);
    expect(parseFrenchNumber('3,50 €')).toBe(3.5);
    expect(parseFrenchNumber(',5')).toBe(0.5);
    expect(parseFrenchNumber('a')).toBeNull();
    expect(parseFrenchNumber('')).toBeNull();
    expect(parseFrenchNumber('1 00')).toBeNull();
  });
  it('texte → date', () => {
    expect(parseFrenchDate('15/03/2024')).toBe(45366);
    expect(parseFrenchDate('15-03-24')).toBe(45366);
    expect(parseFrenchDate('15 mars 2024')).toBe(45366);
    expect(parseFrenchDate('1 février 2024')).toBe(45323);
    expect(parseFrenchDate('31/02/2024')).toBeNull();
    expect(parseFrenchDate('a')).toBeNull();
  });
  it('codes de format français', () => {
    expect(frenchFormatToEnglish('jj/mm/aaaa')).toBe('dd/mm/yyyy');
    expect(frenchFormatToEnglish('jjjj j mmmm aa')).toBe('dddd d mmmm yy');
    expect(frenchFormatToEnglish('# ##0,00')).toBe('#,##0.00');
    expect(frenchFormatToEnglish('0,0')).toBe('0.0');
    expect(frenchFormatToEnglish('0,00E+00')).toBe('0.00E+00');
    expect(frenchFormatToEnglish('"jour "j')).toBe('"jour "d');
    expect(frenchFormatToEnglish('[Rouge]0')).toBe('[Rouge]0');
    expect(frenchFormatToEnglish('hh:mm AM/PM')).toBe('hh:mm AM/PM');
    // Les codes anglais restent compris.
    expect(frenchFormatToEnglish('#,##0.00')).toBe('#,##0.00');
    expect(frenchFormatToEnglish('dd/mm/yyyy')).toBe('dd/mm/yyyy');
  });
  it('critères', () => {
    expect(frenchCriteria('>=12,5')).toBe('>=12.5');
    expect(frenchCriteria('12,5')).toBe('12.5');
    expect(frenchCriteria('<1 000')).toBe('<1000');
    expect(frenchCriteria('>=15/03/2024')).toBe('>=45366');
    expect(frenchCriteria('Lyon')).toBe('Lyon');
    expect(frenchCriteria('>10')).toBe('>10');
    expect(frenchCriteria('a,b')).toBe('a,b');
  });
  it('CHERCHE avec jokers', () => {
    expect(searchWithWildcards('?o', 'Bonjour')).toBe(1);
    expect(searchWithWildcards('j*r', 'Bonjour')).toBe(4);
    expect(searchWithWildcards('z?', 'Bonjour')).toBe(0);
    expect(searchWithWildcards('~?', 'a?b')).toBe(2);
    expect(searchWithWildcards('O?', 'Bonjour', 3)).toBe(5);
  });
  it('CTXT', () => {
    expect(fixedText(1234.567, 1)).toBe('1 234,6');
    expect(fixedText(1234.567, 1, true)).toBe('1234,6');
    expect(fixedText(-0.5, 0)).toBe('-1');
    expect(fixedText(1234.567, -2)).toBe('1 200');
  });
  it('AMORDEGRC (valeur relevée dans Excel)', () => {
    expect(amorDegrc(1000, 45292, 45500, 100, 1, 0.1)).toBe(214);
  });
});
