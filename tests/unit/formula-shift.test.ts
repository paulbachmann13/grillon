import { describe, expect, it } from 'vitest';
import { shiftFormula } from '../../src/data/formula-shift';

describe('décalage des références d’une formule', () => {
  it('référence relative : elle suit la recopie', () => {
    expect(shiftFormula('=A2*2', 2, 0)).toBe('=A4*2');
    expect(shiftFormula('=A2*2', 0, 3)).toBe('=D2*2');
    expect(shiftFormula('=A2+B2', 1, 1)).toBe('=B3+C3');
    expect(shiftFormula('=A2*2', 0, 0)).toBe('=A2*2');
  });
  it('référence absolue : elle ne bouge pas, en tout ou en partie', () => {
    expect(shiftFormula('=$A$1*B2', 3, 3)).toBe('=$A$1*E5');
    expect(shiftFormula('=$A2+B$2', 3, 3)).toBe('=$A5+E$2');
  });
  it('plages, y compris colonnes et lignes entières', () => {
    expect(shiftFormula('=SOMME(A1:A10)', 0, 1)).toBe('=SOMME(B1:B10)');
    expect(shiftFormula('=SUM(A1:B2)', 1, 1)).toBe('=SUM(B2:C3)');
    expect(shiftFormula('=SUM(A:A)', 5, 2)).toBe('=SUM(C:C)');
    expect(shiftFormula('=SUM($A:B)', 0, 1)).toBe('=SUM($A:C)');
    expect(shiftFormula('=SUM(1:3)', 2, 5)).toBe('=SUM(3:5)');
    expect(shiftFormula('=SUM($1:$3)', 2, 5)).toBe('=SUM($1:$3)');
  });
  it('références vers une autre feuille', () => {
    expect(shiftFormula('=Feuil2!A1+1', 1, 0)).toBe('=Feuil2!A2+1');
    expect(shiftFormula("='Ventes 2026'!$B2*C3", 1, 1)).toBe("='Ventes 2026'!$B3*D4");
    expect(shiftFormula('=TVA1!A1', 2, 0)).toBe('=TVA1!A3');
    expect(shiftFormula('=Feuil2!A:A', 0, 1)).toBe('=Feuil2!B:B');
  });
  it('textes, fonctions, nombres, noms et tableaux ne sont pas touchés', () => {
    expect(shiftFormula('=SI(A1="B2";"A1";LOG10(A1))', 1, 1)).toBe('=SI(B2="B2";"A1";LOG10(B2))');
    expect(shiftFormula('=A1*2.5+1E5', 3, 0)).toBe('=A4*2.5+1E5');
    expect(shiftFormula('=TAUX_TVA*A1', 1, 0)).toBe('=TAUX_TVA*A2');
    expect(shiftFormula('=Tableau1[[#Data],[Montant]]*A1', 1, 0)).toBe('=Tableau1[[#Data],[Montant]]*A2');
    expect(shiftFormula('=ATAN2(A1,B1)', 0, 1)).toBe('=ATAN2(B1,C1)');
    expect(shiftFormula('=A1.B1+X1', 1, 0)).toBe('=A1.B1+X2');
  });
  it('colonne ou ligne hors de la feuille : #REF!', () => {
    expect(shiftFormula('=A1', -1, 0)).toBe('=#REF!');
    expect(shiftFormula('=B1', 0, -2)).toBe('=#REF!');
    expect(shiftFormula('=XFD1', 0, 1)).toBe('=#REF!');
    expect(shiftFormula('=A1048576', 1, 0)).toBe('=#REF!');
  });
  it('plusieurs références dans une même formule, vers le haut ou la gauche', () => {
    expect(shiftFormula('=SOMMEPROD(H3:M3;$H$2:$M$2)/$N$2', 1, 0)).toBe('=SOMMEPROD(H4:M4;$H$2:$M$2)/$N$2');
    expect(shiftFormula('=C5+D5', -2, -1)).toBe('=B3+C3');
  });
});
