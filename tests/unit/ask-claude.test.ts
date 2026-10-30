import { describe, expect, it } from 'vitest';
import { buildPrompt, parseReply, toCellContent, type AskCell } from '../../src/data/ask-claude';

const grid: AskCell[][] = [
  [{ shown: 'Nom' }, { shown: 'Ville' }, { shown: 'Note', }],
  [{ shown: 'Léa' }, { shown: 'Lyon' }, { shown: '15', numeric: true }],
  [{ shown: 'Tom' }, { shown: 'Lyon' }, { shown: '8', numeric: true }],
  [{ shown: 'Total' }, { shown: '' }, { shown: '23', numeric: true, formula: '=SOMME(C2:C3)' }],
];

describe('message pour Claude', () => {
  it('contient la demande, la plage, les lettres et numéros, les formules', () => {
    const p = buildPrompt({ question: 'Ajoute la moyenne', sheetName: 'Notes', row: 0, col: 0, cells: grid, header: true });
    expect(p.text).toContain('Ma demande : Ajoute la moyenne');
    expect(p.text).toContain('feuille « Notes », cellules A1:C4');
    expect(p.text).toContain('\tA\tB\tC');
    expect(p.text).toContain('2\tLéa\tLyon\t15');
    expect(p.text).toContain('4\tTotal\t\t=SOMME(C2:C3) → 23');
    expect(p.text).toContain('```grillon');
    expect(p.truncated).toBe(false);
    expect([p.rows, p.cols]).toEqual([4, 3]);
  });
  it('part du bon coin quand la sélection ne commence pas en A1', () => {
    const p = buildPrompt({ question: 'x', sheetName: 'S', row: 4, col: 2, cells: grid.slice(0, 2) });
    expect(p.text).toContain('cellules C5:E6');
    expect(p.text).toContain('\tC\tD\tE');
    expect(p.text).toContain('5\tNom\tVille\tNote');
  });
  it('masque les textes (même texte, même numéro), garde les en-têtes et les nombres', () => {
    const p = buildPrompt({ question: 'x', sheetName: 'S', row: 0, col: 0, cells: grid, hideTexts: true, header: true });
    expect(p.text).toContain('1\tNom\tVille\tNote');
    expect(p.text).toContain('2\tTexte 1\tTexte 2\t15');
    expect(p.text).toContain('3\tTexte 3\tTexte 2\t8');
    expect(p.text).not.toContain('Léa');
    expect(p.text).not.toContain('Lyon');
    expect(p.hiddenTexts).toBe(5);
  });
  it('limite les lignes et le dit', () => {
    const big: AskCell[][] = Array.from({ length: 300 }, (_, i) => [{ shown: String(i), numeric: true }]);
    const p = buildPrompt({ question: 'x', sheetName: 'S', row: 0, col: 0, cells: big, maxRows: 10 });
    expect(p.rows).toBe(10);
    expect(p.totalRows).toBe(300);
    expect(p.truncated).toBe(true);
    expect(p.text).toContain('seulement 10 lignes sur 300');
  });
  it('une cellule sur plusieurs lignes tient sur une ligne', () => {
    const p = buildPrompt({ question: 'x', sheetName: 'S', row: 0, col: 0, cells: [[{ shown: 'a\nb\tc' }]] });
    expect(p.text).toContain('1\ta b c');
  });
});

describe('réponse de Claude', () => {
  it('lit le bloc grillon', () => {
    const r = parseReply('Voici la moyenne :\n\n```grillon\nE1: Moyenne\nE2: =MOYENNE(B2:D2)\nE3: =MOYENNE(B3:D3)\n```\nBonne journée');
    expect(r.edits).toEqual([
      { row: 0, col: 4, content: 'Moyenne' },
      { row: 1, col: 4, content: '=MOYENNE(B2:D2)' },
      { row: 2, col: 4, content: '=MOYENNE(B3:D3)' },
    ]);
  });
  it('tolère puces, $, minuscules, espaces insécables et lignes illisibles', () => {
    const r = parseReply('```grillon\n- $e$2: =SOMME(B2;C2)\nb3:12,5\nn’importe quoi\nA1: Total \n```');
    expect(r.edits.map((e) => [e.row, e.col, e.content])).toEqual([[1, 4, '=SOMME(B2;C2)'], [2, 1, '12,5'], [0, 0, 'Total']]);
    expect(r.ignored).toEqual(['n’importe quoi']);
  });
  it('la dernière écriture d’une même cellule l’emporte', () => {
    expect(parseReply('```grillon\nA1: 1\nA1: 2\n```').edits).toEqual([{ row: 0, col: 0, content: '2' }]);
  });
  it('plusieurs blocs grillon : tous lus ; un bloc sans balise mais bien formé : lu', () => {
    expect(parseReply('```grillon\nA1: 1\n```\ntexte\n```grillon\nA2: 2\n```').edits).toHaveLength(2);
    expect(parseReply('```\nA1: 1\nB2: =A1\n```').edits).toHaveLength(2);
    expect(parseReply('A1: 1\nB2: =A1').edits).toHaveLength(2);
  });
  it('un bloc de code qui n’en est pas un (autre langage) est ignoré', () => {
    expect(parseReply('```python\nprint(1)\n```').edits).toEqual([]);
    expect(parseReply('Je n’ai pas compris, peux-tu préciser ?').edits).toEqual([]);
    expect(parseReply('').edits).toEqual([]);
  });
  it('ne lit pas les cellules hors de la grille et plafonne le nombre', () => {
    expect(parseReply('```grillon\nA0: 1\nXFE1: 2\nA1: 3\n```').edits).toEqual([{ row: 0, col: 0, content: '3' }]);
    const many = Array.from({ length: 700 }, (_, i) => `A${i + 1}: ${i}`).join('\n');
    expect(parseReply(`\`\`\`grillon\n${many}\n\`\`\``).edits).toHaveLength(500);
  });
  it('le contenu : formule en anglais stocké, nombre, logique, texte, vide', () => {
    expect(toCellContent('=SOMME(A1;A2)')).toEqual({ f: '=SUM(A1,A2)' });
    expect(toCellContent('=SI(A1>10;"oui";"non")')).toEqual({ f: '=IF(A1>10,"oui","non")' });
    expect(toCellContent('12,5')).toEqual({ v: 12.5 });
    expect(toCellContent('1 000')).toEqual({ v: 1000 });
    expect(toCellContent('-3')).toEqual({ v: -3 });
    expect(toCellContent('VRAI')).toEqual({ v: true });
    expect(toCellContent('Moyenne')).toEqual({ v: 'Moyenne', text: true });
    expect(toCellContent("'0123")).toEqual({ v: '0123', text: true });
    expect(toCellContent('  ')).toBeNull();
  });
});
