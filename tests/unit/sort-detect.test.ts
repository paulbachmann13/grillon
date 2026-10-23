import { describe, expect, it } from 'vitest';
import { compareFrench, hasHeaderRow, sortRange, type SortCell } from '../../src/data/sort-detect';

const region = { r1: 0, c1: 0, r2: 7, c2: 2 };
const txt = (v: string, bold = false): SortCell => ({ v, t: 1, bold });
const num = (v: number): SortCell => ({ v, t: 2 });

describe('plage à trier (comme Excel)', () => {
  it('une cellule : tout le tableau', () => {
    expect(sortRange({ r1: 2, c1: 0, r2: 2, c2: 0 }, region)).toEqual(region);
  });
  it('une colonne seule : tout le tableau (les lignes restent entières)', () => {
    expect(sortRange({ r1: 0, c1: 2, r2: 7, c2: 2 }, region)).toEqual(region);
    expect(sortRange({ r1: 1, c1: 1, r2: 5, c2: 2 }, region)).toEqual(region);
  });
  it('une sélection qui couvre toutes les colonnes : triée telle quelle', () => {
    expect(sortRange({ r1: 1, c1: 0, r2: 5, c2: 2 }, region)).toEqual({ r1: 1, c1: 0, r2: 5, c2: 2 });
  });
});

describe('ligne d’en-tête', () => {
  it('texte au-dessus de nombres', () => {
    expect(hasHeaderRow([[txt('Nom'), txt('Moyenne')], [txt('Zoé'), num(14)], [txt('Adam'), num(9)]])).toBe(true);
  });
  it('en-tête en gras au-dessus de textes', () => {
    expect(hasHeaderRow([[txt('Nom', true), txt('Prénom', true)], [txt('Zoé'), txt('Lina')], [txt('Adam'), txt('Léa')]])).toBe(true);
  });
  it('liste de noms sans mise en forme : pas d’en-tête', () => {
    expect(hasHeaderRow([[txt('Zoé'), txt('Lina')], [txt('Adam'), txt('Léa')]])).toBe(false);
  });
  it('première ligne avec un nombre : pas d’en-tête', () => {
    expect(hasHeaderRow([[txt('Zoé'), num(12)], [txt('Adam'), num(9)]])).toBe(false);
  });
  it('une seule ligne : pas d’en-tête', () => {
    expect(hasHeaderRow([[txt('Nom'), txt('Note')]])).toBe(false);
  });
  it('cellules vides dans l’en-tête ou dessous', () => {
    expect(hasHeaderRow([[txt('Nom'), { v: null }, txt('Note')], [txt('Zoé'), { v: null }, num(12)]])).toBe(true);
  });
});

describe('ordre alphabétique français', () => {
  it('accents rangés avec leur lettre, casse ignorée', () => {
    const names = ['Zoé', 'Émile', 'adam', 'Éloïse', 'Ethan', 'Bérénice', 'Jean-Marc', 'éa', 'Ea'];
    expect([...names].sort(compareFrench)).toEqual(['adam', 'Bérénice', 'Ea', 'éa', 'Éloïse', 'Émile', 'Ethan', 'Jean-Marc', 'Zoé']);
  });
  it('tirets et apostrophes ignorés', () => {
    expect(compareFrench('Jean-Marc', 'Jeanmarc')).toBe(0);
    expect(compareFrench("d'Artagnan", 'Dartagnan')).toBe(0);
  });
});
