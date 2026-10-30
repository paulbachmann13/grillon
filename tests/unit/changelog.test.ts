import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { compareVersions, notesSince, parseChangelog, parseInline, parseNotes } from '../../src/data/changelog';

const SAMPLE = `# Changelog

## [1.2.0] — 2026-10-30

Intro de la version.

### Corrigé
- **Tri** : la colonne
  continue ici ;
  - détail un
    - détail deux
- Autre \`ligne\`

### Ajouté
- Une fonction

## [1.1.0]

### Ajouté
- Vieille chose

## [0.3.0] — Phase 3 : onglet Données — 2026-10-14

### Ajouté
- Données
`;

describe('historique des versions', () => {
  it('lit les versions, dates et titres', () => {
    const all = parseChangelog(SAMPLE);
    expect(all.map((r) => r.version)).toEqual(['1.2.0', '1.1.0', '0.3.0']);
    expect(all[0]).toMatchObject({ date: '2026-10-30' });
    expect(all[1]!.date).toBeUndefined();
    expect(all[2]).toMatchObject({ date: '2026-10-14', title: 'Phase 3 : onglet Données' });
  });
  it('lit les titres, paragraphes et puces (avec suite de ligne et niveaux)', () => {
    expect(parseChangelog(SAMPLE)[0]!.blocks).toEqual([
      { kind: 'paragraph', text: 'Intro de la version.' },
      { kind: 'heading', text: 'Corrigé' },
      { kind: 'item', text: '**Tri** : la colonne continue ici ;', depth: 0 },
      { kind: 'item', text: 'détail un', depth: 1 },
      { kind: 'item', text: 'détail deux', depth: 2 },
      { kind: 'item', text: 'Autre `ligne`', depth: 0 },
      { kind: 'heading', text: 'Ajouté' },
      { kind: 'item', text: 'Une fonction', depth: 0 },
    ]);
    expect(parseNotes('### Ajouté\n- a')).toEqual([{ kind: 'heading', text: 'Ajouté' }, { kind: 'item', text: 'a', depth: 0 }]);
  });
  it('gras et code, sans HTML', () => {
    expect(parseInline('un **gras** et `code` <b>x</b>')).toEqual([{ text: 'un ' }, { text: 'gras', bold: true }, { text: ' et ' }, { text: 'code', code: true }, { text: ' <b>x</b>' }]);
    expect(parseInline('rien')).toEqual([{ text: 'rien' }]);
  });
  it('compare les versions comme des nombres', () => {
    expect(compareVersions('1.14.10', '1.14.9')).toBeGreaterThan(0);
    expect(compareVersions('1.2', '1.2.0')).toBe(0);
    expect(compareVersions('0.9.0', '1.0.0')).toBeLessThan(0);
  });
  it('ce qui est nouveau depuis la version précédente', () => {
    const all = parseChangelog(SAMPLE);
    expect(notesSince(all, '1.1.0', '1.2.0').map((r) => r.version)).toEqual(['1.2.0']);
    expect(notesSince(all, '0.3.0', '1.2.0').map((r) => r.version)).toEqual(['1.2.0', '1.1.0']);
    expect(notesSince(all, null, '1.2.0').map((r) => r.version)).toEqual(['1.2.0']);
    expect(notesSince(all, '1.2.0', '1.2.0')).toEqual([]);
    // Retour en arrière (version plus ancienne que la dernière vue) : rien à raconter.
    expect(notesSince(all, '1.2.0', '1.1.0')).toEqual([]);
    expect(notesSince(all, '0.3.0', '1.2.0', 1).map((r) => r.version)).toEqual(['1.2.0']);
  });
});

describe('CHANGELOG.md du projet', () => {
  const all = parseChangelog(readFileSync('CHANGELOG.md', 'utf8'));
  const version = (JSON.parse(readFileSync('package.json', 'utf8')) as { version: string }).version;
  it('raconte la version en cours (sinon la mise à jour et « Nouveautés » seraient vides)', () => {
    expect(all[0]!.version).toBe(version);
    expect(all[0]!.date).toMatch(/^\d{4}-\d{2}-\d{2}$/);
    expect(all[0]!.blocks.some((b) => b.kind === 'item')).toBe(true);
  });
  it('les versions se suivent de la plus récente à la plus ancienne, sans doublon', () => {
    const versions = all.map((r) => r.version);
    expect(new Set(versions).size).toBe(versions.length);
    for (let i = 1; i < versions.length; i++) expect(compareVersions(versions[i - 1]!, versions[i]!), `${versions[i - 1]} avant ${versions[i]}`).toBeGreaterThan(0);
  });
});
