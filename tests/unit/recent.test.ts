import { describe, expect, it } from 'vitest';
import { addRecent, folderOf, formatWhen, MAX_RECENT, parseRecent, removeRecent, type RecentFile } from '../../src/data/recent';

const f = (id: string, at = 1): RecentFile => ({ id, name: `${id}.xlsx`, path: `/Users/p/${id}.xlsx`, at });

describe('fichiers récents', () => {
  it('le dernier ouvert passe en tête, sans doublon', () => {
    const list = addRecent(addRecent(addRecent([], f('a')), f('b')), f('a', 5));
    expect(list.map((x) => x.id)).toEqual(['a', 'b']);
    expect(list[0].at).toBe(5);
  });
  it('garde les plus récents seulement', () => {
    let list: RecentFile[] = [];
    for (let i = 0; i < MAX_RECENT + 5; i++) list = addRecent(list, f(`f${i}`));
    expect(list).toHaveLength(MAX_RECENT);
    expect(list[0].id).toBe(`f${MAX_RECENT + 4}`);
  });
  it('retire un fichier', () => {
    expect(removeRecent([f('a'), f('b')], 'a').map((x) => x.id)).toEqual(['b']);
  });
});

describe('affichage', () => {
  it('dossier parent raccourci', () => {
    expect(folderOf('/Users/paul/Documents/Maths/notes.xlsx', '/Users/paul')).toBe('~/Documents/Maths');
    expect(folderOf('/Users/paul/notes.xlsx', '/Users/paul')).toBe('~');
    expect(folderOf('/Volumes/Disque/a/b/c/d/notes.xlsx')).toBe('…/b/c/d');
    expect(folderOf(undefined)).toBe('');
  });
  it('date lisible', () => {
    const now = new Date(2026, 9, 18, 15, 0).getTime();
    expect(formatWhen(new Date(2026, 9, 18, 9, 5).getTime(), now)).toBe('Aujourd’hui, 09:05');
    expect(formatWhen(new Date(2026, 9, 17, 23, 0).getTime(), now)).toBe('Hier');
    expect(formatWhen(new Date(2026, 9, 14, 12, 0).getTime(), now)).toBe('Il y a 4 jours');
    expect(formatWhen(new Date(2026, 8, 2, 12, 0).getTime(), now)).toBe('2 sept.');
    expect(formatWhen(new Date(2025, 11, 25, 12, 0).getTime(), now)).toBe('25 déc. 2025');
  });
  it('un stockage abîmé donne une liste vide', () => {
    expect(parseRecent(null)).toEqual([]);
    expect(parseRecent('pas du json')).toEqual([]);
    expect(parseRecent('{"a":1}')).toEqual([]);
    expect(parseRecent(JSON.stringify([f('a'), { id: 3 }, null, { id: 'x', name: 'x', at: 'hier' }]))).toEqual([f('a')]);
  });
});
