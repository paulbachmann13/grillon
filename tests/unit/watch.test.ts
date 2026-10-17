import { describe, expect, it } from 'vitest';
import { addWatch, dropSheet, removeWatch, renameSheet, sortByColumn, type WatchItem } from '../../src/data/watch';

const a: WatchItem = { sheetId: 's1', sheetName: 'Feuil1', row: 0, col: 0 };
const b: WatchItem = { sheetId: 's1', sheetName: 'Feuil1', row: 2, col: 1 };
const c: WatchItem = { sheetId: 's2', sheetName: 'Feuil2', row: 0, col: 0 };

describe('fenêtre Espion', () => {
  it('ajoute sans doublon', () => {
    expect(addWatch(addWatch([], a), { ...a })).toEqual([a]);
    expect(addWatch([a], c)).toEqual([a, c]);
  });
  it('retire une cellule', () => {
    expect(removeWatch([a, b, c], b)).toEqual([a, c]);
  });
  it('suit le renommage et la suppression d’une feuille', () => {
    expect(renameSheet([a, c], 's1', 'Ventes').map((w) => w.sheetName)).toEqual(['Ventes', 'Feuil2']);
    expect(dropSheet([a, b, c], 's1')).toEqual([c]);
  });
  it('trie par colonne', () => {
    expect(sortByColumn([b, a, c], (w) => w.row).map((w) => w.row)).toEqual([0, 0, 2]);
    expect(sortByColumn([a, c], (w) => w.sheetName, true)[0]).toBe(c);
  });
});
