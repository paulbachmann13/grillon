import { describe, expect, it } from 'vitest';
import { advancedFilter } from '../../src/data/advanced-filter';
import { cellText, type Grid } from '../../src/data/cells';
import { flashFill, inferProgram, runProgram } from '../../src/data/flash-fill';
import { detailBlock, group, hiddenForLevel, ungroup } from '../../src/data/outline';
import { removeDuplicates } from '../../src/data/remove-duplicates';
import { computeSubtotals } from '../../src/data/subtotals';
import { splitText, textToColumns } from '../../src/data/text-to-columns';

const g = (rows: (string | number | null)[][]): Grid => rows.map((r) => r.map((v) => (v === null ? null : { v, t: typeof v === 'number' ? 2 : 1 })));
const texts = (grid: Grid) => grid.map((r) => r.map(cellText));

describe('Supprimer les doublons', () => {
  const grid = g([
    ['Nom', 'Ville'],
    ['Alice', 'Lyon'],
    ['alice', 'lyon'],
    ['Bob', 'Paris'],
    ['Alice', 'Paris'],
    ['Bob', 'Paris'],
  ]);
  it('compare toutes les colonnes, sans tenir compte de la casse, en gardant l’en-tête', () => {
    const r = removeDuplicates(grid, [], true);
    expect(texts(r.rows)).toEqual([['Nom', 'Ville'], ['Alice', 'Lyon'], ['Bob', 'Paris'], ['Alice', 'Paris']]);
    expect([r.removed, r.uniques]).toEqual([2, 3]);
  });
  it('compare seulement les colonnes choisies', () => {
    const r = removeDuplicates(grid, [0], true);
    expect(texts(r.rows).map((x) => x[0])).toEqual(['Nom', 'Alice', 'Bob']);
  });
  it('considère 1 et "1" comme identiques', () => {
    expect(removeDuplicates(g([[1], ['1'], [2]]), [0], false).removed).toBe(1);
  });
});

describe('Convertir (texte en colonnes)', () => {
  it('découpe selon les séparateurs, avec guillemets et séparateurs consécutifs', () => {
    expect(splitText('a;b;;c', { semicolon: true })).toEqual(['a', 'b', '', 'c']);
    expect(splitText('a  b', { space: true, consecutiveAsOne: true })).toEqual(['a', 'b']);
    expect(splitText('"Dupont, Jean",42', { comma: true })).toEqual(['Dupont, Jean', '42']);
    expect(splitText('x|y', { other: '|' })).toEqual(['x', 'y']);
  });
  it('type les morceaux (nombres à virgule, dates)', () => {
    const [row] = textToColumns(['Lyon;1,5;14/10/2026'], { semicolon: true, decimal: ',' });
    expect(row[0]).toMatchObject({ v: 'Lyon' });
    expect(row[1]).toMatchObject({ v: 1.5, t: 2 });
    expect(row[2]).toMatchObject({ v: 46309 });
  });
});

describe('Remplissage instantané', () => {
  it('extrait le prénom', () => {
    const out = flashFill([
      { sources: ['Jean Dupont'], target: 'Jean' },
      { sources: ['Marie Curie'], target: '' },
      { sources: ['Ada Lovelace'], target: '' },
    ]);
    expect(out).toEqual(['Jean', 'Marie', 'Ada']);
  });
  it('recompose « Nom, Prénom » en majuscules', () => {
    const out = flashFill([
      { sources: ['jean dupont'], target: 'DUPONT, Jean' },
      { sources: ['marie curie'], target: '' },
    ]);
    expect(out).toEqual(['DUPONT, Jean', 'CURIE, Marie']);
  });
  it('construit une adresse e-mail à partir de deux colonnes', () => {
    const out = flashFill([
      { sources: ['Jean', 'Dupont'], target: 'jean.dupont@ecole.fr' },
      { sources: ['Marie', 'Curie'], target: 'marie.curie@ecole.fr' },
      { sources: ['Pierre', 'Martin'], target: '' },
    ]);
    expect(out?.[2]).toBe('pierre.martin@ecole.fr');
  });
  it('extrait le domaine d’une adresse e-mail et les initiales', () => {
    expect(flashFill([{ sources: ['a.b@site.com'], target: 'site.com' }, { sources: ['x@autre.org'], target: '' }])?.[1]).toBe('autre.org');
    const p = inferProgram([{ sources: ['Jean Dupont'], target: 'J.D.' }, { sources: ['Marie Curie'], target: 'M.C.' }]);
    expect(p && runProgram(p, ['Ada Lovelace'])).toBe('A.L.');
  });
  it('renvoie null quand aucun motif ne correspond', () => {
    expect(flashFill([{ sources: ['abc'], target: 'zzz' }, { sources: ['def'], target: '' }])).toBeNull();
  });
});

describe('Sous-totaux', () => {
  it('insère un total par groupe, un total général et les niveaux de plan', () => {
    const grid = g([
      ['Ville', 'Montant'],
      ['Lyon', 10],
      ['Lyon', 20],
      ['Paris', 5],
    ]);
    const r = computeSubtotals(grid, { groupBy: 0, fn: 'sum', columns: [1], startRow: 0, startColumn: 0 });
    expect(r.rows.map((row) => [cellText(row[0]), row[1]?.f ?? cellText(row[1])])).toEqual([
      ['Ville', 'Montant'],
      ['Lyon', '10'],
      ['Lyon', '20'],
      ['Total Lyon', '=SUBTOTAL(9,B2:B3)'],
      ['Paris', '5'],
      ['Total Paris', '=SUBTOTAL(9,B5:B5)'],
      ['Total général', '=SUBTOTAL(9,B2:B6)'],
    ]);
    expect(r.outline).toEqual([0, 2, 2, 1, 2, 1, 0]);
  });
  it('tient compte de la position de la plage', () => {
    const r = computeSubtotals(g([['V', 'N'], ['a', 1]]), { groupBy: 0, fn: 'average', columns: [1], startRow: 4, startColumn: 2 });
    expect(r.rows[2][1]?.f).toBe('=SUBTOTAL(1,D6:D6)');
  });
});

describe('Filtre avancé', () => {
  const list = g([
    ['Nom', 'Ville', 'Montant'],
    ['Alice', 'Lyon', 120],
    ['Bob', 'Paris', 80],
    ['Chloé', 'Lyon', 40],
    ['David', 'Lille', 200],
    ['Alice', 'Lyon', 120],
  ]);
  it('combine ET sur une ligne et OU entre lignes', () => {
    const crit = g([
      ['Ville', 'Montant'],
      ['Lyon', '>100'],
      ['Paris', null],
    ]);
    expect(advancedFilter(list, crit).rows).toEqual([0, 1, 4]);
  });
  it('« commence par », jokers, égalité exacte et différent', () => {
    expect(advancedFilter(list, g([['Ville'], ['L']])).rows).toEqual([0, 2, 3, 4]);
    expect(advancedFilter(list, g([['Ville'], ['=L*e']])).rows).toEqual([3]);
    expect(advancedFilter(list, g([['Ville'], ['<>Lyon']])).rows).toEqual([1, 3]);
    expect(advancedFilter(list, g([['Nom'], ['=Chlo?']])).rows).toEqual([2]);
  });
  it('extraction sans doublon', () => {
    expect(advancedFilter(list, g([['Ville'], ['Lyon']]), true).rows).toEqual([0, 2]);
  });
});

describe('Plan (grouper/dissocier)', () => {
  it('groupe, dissocie et trouve le bloc de détail', () => {
    let levels = [0, 0, 0, 0, 0, 0];
    levels = group(levels, 1, 4);
    levels = group(levels, 2, 3);
    expect(levels).toEqual([0, 1, 2, 2, 1, 0]);
    expect(detailBlock(levels, 2)).toEqual({ from: 2, to: 3, level: 2 });
    expect(detailBlock(levels, 1)).toEqual({ from: 1, to: 4, level: 1 });
    expect(detailBlock(levels, 0)).toBeNull();
    expect(hiddenForLevel(levels, 2)).toEqual([false, false, true, true, false, false]);
    expect(ungroup(levels, 1, 4)).toEqual([0, 0, 1, 1, 0, 0]);
  });
});
