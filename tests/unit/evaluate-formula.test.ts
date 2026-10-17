import { describe, expect, it } from 'vitest';
import { nextStep, tokenize } from '../../src/data/evaluate-formula';
import { formulaLiteral, frenchError, numberFromText } from '../../src/data/formula-errors';

const step = (expr: string) => {
  const s = nextStep(expr);
  return s ? expr.slice(s[0], s[1]) : null;
};

describe('Évaluer la formule', () => {
  it('reconnaît références, fonctions, littéraux', () => {
    const kinds = (e: string) => tokenize(e).map((t) => `${t.kind}:${e.slice(t.start, t.end)}`);
    expect(kinds('SUM(A1:B2,Feuil2!$C$3)*2')).toEqual(['fn:SUM', 'op:(', 'ref:A1:B2', 'op:,', 'ref:Feuil2!$C$3', 'op:)', 'op:*', 'lit:2']);
    expect(kinds(`'Ma feuille'!A1&"x""y"`)).toEqual([`ref:'Ma feuille'!A1`, 'op:&', 'lit:"x""y"']);
    expect(kinds('Ventes[Prix]+Taux+TRUE+#N/A')).toEqual(['ref:Ventes[Prix]', 'op:+', 'ref:Taux', 'op:+', 'lit:TRUE', 'op:+', 'lit:#N/A']);
    expect(kinds('A:A+1:1')).toEqual(['ref:A:A', 'op:+', 'ref:1:1']);
    expect(kinds('LOG10(100)')).toEqual(['fn:LOG10', 'op:(', 'lit:100', 'op:)']);
  });

  it('avance comme Excel : références, puis fonctions, puis le reste', () => {
    expect(step('SUM(F1:F3)*2+IF(F1>1,MAX(F2,F3),0)')).toBe('F1:F3');
    expect(step('SUM({2;3;4})*2+IF(F1>1,MAX(F2,F3),0)')).toBe('SUM({2;3;4})');
    expect(step('9*2+IF(F1>1,MAX(F2,F3),0)')).toBe('F1');
    expect(step('9*2+IF(2>1,MAX(3,4),0)')).toBe('MAX(3,4)');
    expect(step('9*2+IF(2>1,4,0)')).toBe('IF(2>1,4,0)');
    expect(step('9*2+4')).toBe('9*2+4');
    expect(step('22')).toBeNull();
    expect(step('-1.5')).toBeNull();
    expect(step('"texte"')).toBeNull();
    expect(step('#DIV/0!')).toBeNull();
  });

  it('littéraux et erreurs', () => {
    expect(formulaLiteral([[2], [3]])).toBe('{2;3}');
    expect(formulaLiteral('a"b')).toBe('"a""b"');
    expect(formulaLiteral(true)).toBe('TRUE');
    expect(formulaLiteral('#N/A')).toBe('#N/A');
    expect(frenchError('#NAME?')).toBe('#NOM?');
    expect(numberFromText('1 234,5')).toBe(1234.5);
    expect(numberFromText('12%')).toBe(0.12);
    expect(numberFromText('abc')).toBeNull();
    expect(numberFromText('12/03')).toBeNull();
  });
});
