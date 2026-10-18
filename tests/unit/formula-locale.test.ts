import { describe, expect, it } from 'vitest';
import { englishToFrench, EN_TO_FR, frenchToEnglish, FR_TO_EN, looksEnglish } from '../../src/data/formula-locale';

describe('formules en français → anglais', () => {
  it('fonctions courantes et séparateur « ; »', () => {
    expect(frenchToEnglish('=SOMME(A1:A10)')).toBe('=SUM(A1:A10)');
    expect(frenchToEnglish('=SI(A1>10;"grand";"petit")')).toBe('=IF(A1>10,"grand","petit")');
    expect(frenchToEnglish('=RECHERCHEV(A2;Feuil2!A:C;3;FAUX)')).toBe('=VLOOKUP(A2,Feuil2!A:C,3,FALSE)');
    expect(frenchToEnglish('=SOMME.SI.ENS(C:C;A:A;"Nord";B:B;">=10")')).toBe('=SUMIFS(C:C,A:A,"Nord",B:B,">=10")');
    expect(frenchToEnglish('=SIERREUR(RECHERCHEX(A1;B:B;C:C);"?")')).toBe('=IFERROR(XLOOKUP(A1,B:B,C:C),"?")');
  });
  it('fonctions imbriquées, espaces et casse', () => {
    expect(frenchToEnglish('=si( et(A1>1 ; A2<5) ; somme(B1:B3) ; 0 )')).toBe('=IF( AND(A1>1 , A2<5) , SUM(B1:B3) , 0 )');
  });
  it('virgule décimale et VRAI / FAUX', () => {
    expect(frenchToEnglish('=ARRONDI(A1*1,196;2)')).toBe('=ROUND(A1*1.196,2)');
    expect(frenchToEnglish('=SI(A1=VRAI;0,5;1E-3)')).toBe('=IF(A1=TRUE,0.5,1E-3)');
    expect(frenchToEnglish('=A1*0,2+B2')).toBe('=A1*0.2+B2');
  });
  it('les textes, les noms de feuille et les références ne sont pas touchés', () => {
    expect(frenchToEnglish('=SI(A1="SOMME(1;2)";\'Ventes; 2026\'!B2;0)')).toBe('=IF(A1="SOMME(1;2)",\'Ventes; 2026\'!B2,0)');
    expect(frenchToEnglish('=SOMME(Feuil1!A1:B2)')).toBe('=SUM(Feuil1!A1:B2)');
    expect(frenchToEnglish('=SOMME(A1:A3)+TAUX_TVA')).toBe('=SUM(A1:A3)+TAUX_TVA');
    expect(frenchToEnglish('=SI(A1="a;b";"c,d";1,5)')).toBe('=IF(A1="a;b","c,d",1.5)');
  });
  it('constantes matricielles : « \\ » entre colonnes, « ; » entre lignes', () => {
    expect(frenchToEnglish('=SOMME({1\\2;3\\4})')).toBe('=SUM({1,2;3,4})');
    expect(frenchToEnglish('=INDEX({1,5\\2;3\\4};1;2)')).toBe('=INDEX({1.5,2;3,4},1,2)');
  });
  it('erreurs littérales et références de tableau', () => {
    expect(frenchToEnglish('=SI(A1=#N/A;#NOM?;#VALEUR!)')).toBe('=IF(A1=#N/A,#NAME?,#VALUE!)');
    expect(frenchToEnglish('=SOMME(Tableau1[[#Données];[Montant]])')).toBe('=SUM(Tableau1[[#Data];[Montant]])');
  });
  it('une formule déjà en anglais est laissée telle quelle', () => {
    expect(frenchToEnglish('=SUM(A1,B1)')).toBe('=SUM(A1,B1)');
    expect(frenchToEnglish('=IF(A1>1,"a","b")')).toBe('=IF(A1>1,"a","b")');
    expect(frenchToEnglish('=VLOOKUP(A1,B:C,2,FALSE)')).toBe('=VLOOKUP(A1,B:C,2,FALSE)');
    expect(looksEnglish('=SUM(A1,B1)')).toBe(true);
    expect(looksEnglish('=SOMME(A1;B1)')).toBe(false);
    expect(looksEnglish('=SUM(A1;B1)')).toBe(false);
  });
  it('texte simple et nombres inchangés ; fonctions inconnues conservées', () => {
    expect(frenchToEnglish('bonjour')).toBe('bonjour');
    expect(frenchToEnglish('1,5')).toBe('1,5');
    expect(frenchToEnglish('=MAFONCTION(A1;2)')).toBe('=MAFONCTION(A1,2)');
    expect(frenchToEnglish('=ABS(A1)+MAX(A1:A3;0)')).toBe('=ABS(A1)+MAX(A1:A3,0)');
  });
});

describe('formules anglais → français (barre de formule)', () => {
  it('noms, séparateurs et décimales', () => {
    expect(englishToFrench('=SUM(A1:A10)')).toBe('=SOMME(A1:A10)');
    expect(englishToFrench('=IF(A1>10,"grand","petit")')).toBe('=SI(A1>10;"grand";"petit")');
    expect(englishToFrench('=ROUND(A1*1.196,2)')).toBe('=ARRONDI(A1*1,196;2)');
    expect(englishToFrench('=IF(A1=TRUE,0.5,FALSE)')).toBe('=SI(A1=VRAI;0,5;FAUX)');
    expect(englishToFrench('=TRIM(A1)')).toBe('=SUPPRESPACE(A1)');
  });
  it('textes et noms de feuille conservés', () => {
    expect(englishToFrench('=IF(A1="SUM(1,2)",\'Ventes, 2026\'!B2,0)')).toBe('=SI(A1="SUM(1,2)";\'Ventes, 2026\'!B2;0)');
    expect(englishToFrench('=Feuil1!A1+VRAI1')).toBe('=Feuil1!A1+VRAI1');
  });
  it('constantes matricielles', () => {
    expect(englishToFrench('=SUM({1,2;3,4})')).toBe('=SOMME({1\\2;3\\4})');
  });
  it('erreurs littérales', () => {
    expect(englishToFrench('=IF(A1=#N/A,#NAME?,#VALUE!)')).toBe('=SI(A1=#N/A;#NOM?;#VALEUR!)');
  });
  it('un texte qui n’est pas une formule est inchangé', () => {
    expect(englishToFrench('SUM(1,2)')).toBe('SUM(1,2)');
  });
});

describe('aller-retour', () => {
  const formulas = [
    '=SUM(A1:A10)',
    '=IF(AND(A1>1,A2<5),SUM(B1:B3),0)',
    '=VLOOKUP(A2,Feuil2!A:C,3,FALSE)',
    '=IFERROR(XLOOKUP(A1,B:B,C:C),"n/a")',
    '=SUMIFS(C:C,A:A,"Nord",B:B,">=10")',
    '=ROUND(A1*1.196,2)+0.5',
    '=SUM({1,2;3,4})',
    '=TEXT(TODAY(),"jj/mm/aaaa")',
    '=DATEDIF(A1,B1,"d")+NETWORKDAYS(A1,B1)',
    '=LET(x,A1*2,x+1)',
  ];
  for (const f of formulas) {
    it(f, () => expect(frenchToEnglish(englishToFrench(f))).toBe(f));
  }
});

describe('table des fonctions', () => {
  it('pas de nom en double côté anglais, et les noms français sont en majuscules sans espace', () => {
    expect(EN_TO_FR.size).toBeGreaterThan(300);
    for (const [en, fr] of EN_TO_FR) {
      expect(en).toMatch(/^[A-Z0-9.]+$/);
      expect(fr).toMatch(/^[A-Z0-9._]+$/);
    }
    expect(FR_TO_EN.get('SOMME')).toBe('SUM');
    expect(FR_TO_EN.get('RECHERCHEV')).toBe('VLOOKUP');
  });
});
