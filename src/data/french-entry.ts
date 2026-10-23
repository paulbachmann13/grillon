/**
 * Saisie « à la française », comme Excel en français, pour les cas qu'Univer lit mal :
 * - espace des milliers : « 1 234,5 », « 1 000 000 », « -1 234 » (espaces normales, insécables ou fines) ;
 * - montants en euros avec virgule ou espaces : « -12,50 € », « 1 250 € » ;
 * - point décimal : « 1.5 » vaut 1,5 (Univer en faisait 15, en lisant le point comme séparateur de milliers) ;
 * - date sans année : « 3/4 » est le 3 avril (Univer y voyait le 4 mars), « 31/12 », « 13/1 » de l'année en cours ;
 * - VRAI / FAUX (toutes casses) deviennent des valeurs logiques.
 * Le reste est laissé à Univer (renvoie null). Logique pure.
 */
import { ymdToSerial } from './slicer';

export type FrenchEntry =
  | { kind: 'number'; value: number; pattern?: string }
  | { kind: 'boolean'; value: boolean }
  | { kind: 'date'; value: number; pattern: string };

const SPACES = '[ \\u00a0\\u202f]';
const GROUPED = new RegExp(`^(-?)(\\d{1,3}(?:${SPACES}\\d{3})+)(?:,(\\d+))?$`);
const EURO = new RegExp(`^(-?)(\\d{1,3}(?:${SPACES}?\\d{3})*|\\d+)(?:,(\\d+))?${SPACES}*€$`);
const DOT_DECIMAL = /^(-?)(\d+)\.(\d+)$/;
const DAY_MONTH = /^(\d{1,2})[/-](\d{1,2})$/;

const digits = (s: string) => s.replace(/[   ]/g, '');
const decimalsPattern = (base: string, decimals: string | undefined) => (decimals ? `${base}.${'0'.repeat(Math.min(decimals.length, 10))}` : base);

export function parseFrenchEntry(input: string, year = new Date().getFullYear()): FrenchEntry | null {
  const text = input.trim();
  if (!text || text.startsWith('=') || text.startsWith("'")) return null;
  if (/^vrai$/i.test(text)) return { kind: 'boolean', value: true };
  if (/^faux$/i.test(text)) return { kind: 'boolean', value: false };

  let m = GROUPED.exec(text);
  if (m) {
    const value = Number(`${m[1]}${digits(m[2])}.${m[3] ?? '0'}`);
    return { kind: 'number', value, pattern: decimalsPattern('#,##0', m[3]) };
  }
  m = EURO.exec(text);
  // « 5 € » et « 25 € » sont déjà bien lus par Univer : on ne traite que ce qu'il laisse en texte (virgule, signe, milliers).
  if (m && (m[1] || m[3] !== undefined || /[   ]/.test(m[2]))) {
    const value = Number(`${m[1]}${digits(m[2])}.${m[3] ?? '0'}`);
    return { kind: 'number', value, pattern: `${decimalsPattern('#,##0', m[3])} €` };
  }
  m = DOT_DECIMAL.exec(text);
  if (m) return { kind: 'number', value: Number(`${m[1]}${m[2]}.${m[3]}`) };

  m = DAY_MONTH.exec(text);
  if (m) {
    const day = Number(m[1]);
    const month = Number(m[2]);
    if (month < 1 || month > 12 || day < 1) return null;
    // Le 31/02 n'existe pas : on laisse le texte tel quel, comme Excel.
    if (new Date(Date.UTC(year, month - 1, day)).getUTCMonth() !== month - 1) return null;
    return { kind: 'date', value: ymdToSerial(year, month, day), pattern: 'dd/mm' };
  }
  return null;
}
