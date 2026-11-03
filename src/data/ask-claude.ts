/**
 * « Demander à Claude » : prépare le message à coller dans Claude (la question et les cellules sélectionnées, en texte),
 * puis lit sa réponse (un bloc « grillon » : une cellule par ligne) pour la coller dans la feuille. Rien n'est envoyé par Grillon :
 * on copie, on colle. Logique pure, sans DOM ni Univer.
 */
import { columnToLetters } from '../io/xlsx/units';
import { frenchToEnglish } from './formula-locale';
import { parseFrenchNumber } from './french-values';

/** Une cellule à envoyer : ce qu'on y voit, et sa formule (en français) si elle en a une. */
export interface AskCell {
  shown: string;
  formula?: string;
  /** Vrai pour un nombre, une date ou un logique (jamais masqué). */
  numeric?: boolean;
}

export interface PromptOptions {
  question: string;
  sheetName: string;
  /** Position de la première cellule (indices à partir de 0). */
  row: number;
  col: number;
  cells: AskCell[][];
  /** Remplace les textes (noms, villes…) par « Texte 1 », « Texte 2 »… sauf la ligne d'en-têtes. */
  hideTexts?: boolean;
  /** La première ligne est une ligne d'en-têtes (gardée telle quelle). */
  header?: boolean;
  maxRows?: number;
  maxCols?: number;
  /** Taille réelle de la zone quand `cells` n'en contient qu'une partie (lecture limitée). */
  totalRows?: number;
  totalCols?: number;
}

export interface Prompt {
  text: string;
  rows: number;
  totalRows: number;
  cols: number;
  totalCols: number;
  truncated: boolean;
  hiddenTexts: number;
}

export const DEFAULT_MAX_ROWS = 100;
export const DEFAULT_MAX_COLS = 30;

const INSTRUCTIONS = [
  'Tu m’aides dans Grillon, un tableur compatible Excel, en français : les formules s’écrivent en français (SOMME, SI, RECHERCHEV, NB.SI.ENS…), avec « ; » entre les arguments et la virgule comme séparateur décimal.',
  'Réponds en français, simplement, sans jargon (je ne suis pas informaticien) : explique en une ou deux phrases ce que tu proposes.',
  'Si ma demande se règle en écrivant dans des cellules, termine ta réponse par UN SEUL bloc de code balisé « grillon », une cellule par ligne, sous la forme « E2: =MOYENNE(B2:D2) » (une formule commence par =). N’écris que dans les cellules nécessaires et n’invente aucune donnée. Si une information te manque, pose-moi une question au lieu d’écrire le bloc.',
  'Exemple de bloc :',
  '```grillon',
  'E1: Moyenne',
  'E2: =MOYENNE(B2:D2)',
  'E3: =MOYENNE(B3:D3)',
  '```',
].join('\n');

/** Texte → « Texte N » (même texte, même numéro dans toute la sélection, pour que les recherches et les comparaisons marchent encore). */
function hider(): (text: string) => string {
  const seen = new Map<string, number>();
  return (text) => {
    const key = text.trim().toLocaleLowerCase('fr');
    if (!seen.has(key)) seen.set(key, seen.size + 1);
    return `Texte ${seen.get(key)}`;
  };
}

/** Textes écrits en dur dans une formule (="Léa", NB.SI(B:B;"<>Lyon")) : masqués aussi, en gardant l'opérateur de comparaison ; les valeurs sans lettre (">5", ";") restent. */
function hideLiterals(formula: string, hide: (text: string) => string): string {
  return formula.replace(/"((?:[^"]|"")*)"/g, (whole, body: string) => {
    const [, op = '', rest = ''] = /^(<>|>=|<=|=|<|>)?([\s\S]*)$/.exec(body) ?? [];
    return /\p{L}/u.test(rest) ? `"${op}${hide(rest.replace(/""/g, '"'))}"` : whole;
  });
}

const oneLine = (s: string) => s.replace(/[\t\r\n]+/g, ' ').trim();

/** Le message complet à coller dans Claude, avec le nombre de lignes envoyées (pour le montrer avant de copier). */
export function buildPrompt(o: PromptOptions): Prompt {
  const maxRows = o.maxRows ?? DEFAULT_MAX_ROWS;
  const maxCols = o.maxCols ?? DEFAULT_MAX_COLS;
  const totalRows = o.totalRows ?? o.cells.length;
  const totalCols = o.totalCols ?? Math.max(0, ...o.cells.map((r) => r.length));
  const rows = Math.min(totalRows, o.cells.length, maxRows + (o.header ? 1 : 0));
  const cols = Math.min(totalCols, maxCols);
  const hide = hider();
  let hiddenTexts = 0;
  const lines: string[] = [];
  const letters = Array.from({ length: cols }, (_, j) => columnToLetters(o.col + j));
  lines.push(['', ...letters].join('\t'));
  for (let i = 0; i < rows; i++) {
    const out: string[] = [String(o.row + i + 1)];
    for (let j = 0; j < cols; j++) {
      const cell = o.cells[i]?.[j];
      if (!cell) {
        out.push('');
        continue;
      }
      let shown = oneLine(cell.shown);
      if (o.hideTexts && !cell.numeric && shown && !(o.header && i === 0)) {
        // Une formule garde sa forme (Claude en a besoin) ; seul son résultat texte est masqué.
        shown = hide(shown);
        hiddenTexts++;
      }
      out.push(cell.formula ? `${o.hideTexts ? hideLiterals(cell.formula, hide) : cell.formula}${shown ? ` → ${shown}` : ''}` : shown);
    }
    lines.push(out.join('\t'));
  }
  const first = `${columnToLetters(o.col)}${o.row + 1}`;
  const last = `${columnToLetters(o.col + Math.max(0, cols - 1))}${o.row + rows}`;
  const truncated = rows < totalRows || cols < totalCols;
  const text = [
    INSTRUCTIONS,
    '',
    `Ma demande : ${o.question.trim() || '(à préciser)'}`,
    '',
    `Voici ${o.hideTexts ? 'ma feuille' : `ma feuille « ${o.sheetName} »`}, cellules ${first}:${last}${truncated ? ` (seulement ${rows} lignes sur ${totalRows} et ${cols} colonnes sur ${totalCols}, le reste n’est pas montré)` : ''}.`,
    'Les colonnes sont désignées par leur lettre et les lignes par leur numéro (première ligne du tableau ci-dessous : les lettres). Les cellules sont séparées par des tabulations ; une cellule qui commence par = est une formule, suivie de → et de son résultat.',
    '',
    lines.join('\n'),
  ].join('\n');
  return { text, rows, totalRows, cols, totalCols, truncated, hiddenTexts };
}

// --- La réponse de Claude ---

export interface ReplyEdit {
  /** Ligne et colonne (indices à partir de 0). */
  row: number;
  col: number;
  /** Ce que Claude écrit : tel quel (« =SOMME(A1;A2) », « 12,5 », « Total »). Vide : vider la cellule. */
  content: string;
}

export interface ParsedReply {
  edits: ReplyEdit[];
  /** Lignes du bloc qui n'ont pas pu être lues. */
  ignored: string[];
}

const LINE = /^\s*(?:[-*]\s+)?\$?([A-Za-z]{1,3})\$?(\d{1,7})\s*:\s?(.*)$/;
export const MAX_EDITS = 500;

function columnIndex(letters: string): number {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

function parseLines(block: string): ParsedReply {
  const edits = new Map<string, ReplyEdit>();
  const ignored: string[] = [];
  for (const raw of block.replace(/\r/g, '').split('\n')) {
    if (!raw.trim()) continue;
    const m = LINE.exec(raw);
    if (!m) {
      ignored.push(raw.trim());
      continue;
    }
    const col = columnIndex(m[1]!);
    const row = Number(m[2]) - 1;
    if (row < 0 || col > 16383 || row > 1048575) {
      ignored.push(raw.trim());
      continue;
    }
    // Les guillemets typographiques ou les espaces insécables que mettent parfois les messageries.
    const content = m[3]!.replace(/[  ]/g, ' ').replace(/\s+$/, '');
    edits.set(`${row},${col}`, { row, col, content });
  }
  return { edits: [...edits.values()].slice(0, MAX_EDITS), ignored };
}

/** Lit la réponse de Claude : le bloc « grillon » (ou, à défaut, un bloc de code dont toutes les lignes sont « cellule: contenu »). */
export function parseReply(reply: string): ParsedReply {
  const text = reply.replace(/\r/g, '');
  const fences = [...text.matchAll(/```[ \t]*([A-Za-z0-9_-]*)[^\n]*\n([\s\S]*?)```/g)];
  const named = fences.filter((f) => f[1]!.toLowerCase() === 'grillon');
  if (named.length) return parseLines(named.map((f) => f[2]).join('\n'));
  for (const f of fences) {
    const lines = f[2]!.split('\n').filter((l) => l.trim());
    if (lines.length && lines.every((l) => LINE.test(l))) return parseLines(f[2]!);
  }
  // Pas de bloc : un message qui ne contient que « cellule: contenu » (ou rien d'exploitable).
  const lines = text.split('\n').filter((l) => l.trim());
  if (lines.length && lines.every((l) => LINE.test(l))) return parseLines(text);
  return { edits: [], ignored: [] };
}

/** Ce que la cellule devient dans Univer : formule (anglais stocké), nombre, logique ou texte ; null pour vider. */
export type CellContent = { f: string } | { v: number | string | boolean; text?: true } | null;

export function toCellContent(content: string): CellContent {
  const text = content.trim();
  if (!text) return null;
  if (text.startsWith('=')) return { f: frenchToEnglish(text) };
  if (/^vrai$/i.test(text)) return { v: true };
  if (/^faux$/i.test(text)) return { v: false };
  const n = parseFrenchNumber(text) ?? (/^-?\d+(\.\d+)?$/.test(text) ? Number(text) : null);
  if (n !== null) return { v: n };
  // Texte (une apostrophe de tête force le texte, comme dans Excel).
  return { v: text.startsWith("'") ? text.slice(1) : text, text: true };
}
