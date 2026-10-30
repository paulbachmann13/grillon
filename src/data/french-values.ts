/**
 * Valeurs « à la française » du moteur de formules, comme Excel en français : conversion nombre → texte (virgule décimale),
 * texte → nombre ou date (« 12,5 », « 1 000 », « 15/03/2024 »), codes de format de TEXTE (« jj/mm/aaaa », « # ##0,00 »),
 * critères de NB.SI (« >=12,5 »). Logique pure.
 */
import { ymdToSerial } from './slicer';

/** Nombre → texte comme Excel en français (« "x"&1,5 » donne « x1,5 ») : 15 chiffres significatifs, sans largeur limite. */
export function frenchNumberText(n: number): string {
  if (!Number.isFinite(n)) return String(n);
  const clean = Number(n.toPrecision(15));
  const abs = Math.abs(clean);
  if (abs !== 0 && (abs >= 1e21 || abs < 1e-19)) {
    const [mant, exp] = clean.toExponential().split('e');
    return `${mant!.replace('.', ',')}E${Number(exp) < 0 ? '-' : '+'}${String(Math.abs(Number(exp))).padStart(2, '0')}`;
  }
  const text = abs >= 1e15 || (abs !== 0 && abs < 1e-6) ? clean.toLocaleString('en-US', { useGrouping: false, maximumSignificantDigits: 15 }) : String(clean);
  return text.replace('.', ',');
}

const SPACE = '[ \\u00a0\\u202f]';
const NUMBER = new RegExp(`^(-?)(\\d{1,3}(?:${SPACE}\\d{3})+|\\d*)(?:,(\\d+))?(?:${SPACE}*(%|€))?$`);

/** Texte → nombre à la française (« 12,5 », « 1 000 », « -1 234,5 », « 12,5 % », « 3,50 € ») ; null si ce n'en est pas un. */
export function parseFrenchNumber(input: string): number | null {
  const m = NUMBER.exec(input.trim());
  if (!m || (!m[2] && !m[3])) return null;
  const value = Number(`${m[1]}${(m[2] || '0').replace(new RegExp(SPACE, 'g'), '')}.${m[3] ?? '0'}`);
  return m[4] === '%' ? value / 100 : value;
}

const MONTHS: Record<string, number> = {
  janvier: 1, janv: 1, fevrier: 2, fevr: 2, mars: 3, avril: 4, avr: 4, mai: 5, juin: 6, juillet: 7, juil: 7, aout: 8, septembre: 9, sept: 9, octobre: 10, oct: 10, novembre: 11, nov: 11, decembre: 12, dec: 12,
};

/** Texte → date à la française (« 15/03/2024 », « 15-03-24 », « 15 mars 2024 ») ; numéro de série, ou null. */
export function parseFrenchDate(input: string): number | null {
  const text = input.trim().toLowerCase().normalize('NFD').replace(/\p{M}/gu, '');
  let m = /^(\d{1,2})[/-](\d{1,2})[/-](\d{4}|\d{2})$/.exec(text);
  let day: number, month: number, year: number;
  if (m) {
    [day, month, year] = [Number(m[1]), Number(m[2]), Number(m[3])];
  } else {
    m = /^(\d{1,2})[ -]([a-z]+)\.?[ -](\d{4}|\d{2})$/.exec(text);
    if (!m || !MONTHS[m[2]!]) return null;
    [day, month, year] = [Number(m[1]), MONTHS[m[2]!]!, Number(m[3])];
  }
  // Comme Excel : deux chiffres de 00 à 29 → 20xx, de 30 à 99 → 19xx.
  if (m[3]!.length === 2) year += year < 30 ? 2000 : 1900;
  if (month < 1 || month > 12 || day < 1 || day > new Date(Date.UTC(year, month, 0)).getUTCDate()) return null;
  return ymdToSerial(year, month, day);
}

/**
 * Code de format français de TEXTE → code du moteur : jour j, année a, virgule décimale, espace des milliers.
 * Les parties entre guillemets ou crochets et les caractères échappés restent tels quels. Les codes anglais restent acceptés.
 */
export function frenchFormatToEnglish(format: string): string {
  let out = '';
  for (let i = 0; i < format.length; i++) {
    const ch = format[i]!;
    if (ch === '"') {
      const end = format.indexOf('"', i + 1);
      const stop = end < 0 ? format.length : end + 1;
      out += format.slice(i, stop);
      i = stop - 1;
    } else if (ch === '[') {
      const end = format.indexOf(']', i + 1);
      const stop = end < 0 ? format.length : end + 1;
      out += format.slice(i, stop);
      i = stop - 1;
    } else if (ch === '\\' && i + 1 < format.length) {
      out += format.slice(i, i + 2);
      i++;
    } else if (ch === 'j' || ch === 'J') out += 'd';
    else if (ch === 'a' || ch === 'A') {
      // « AM/PM » et « A/P » gardent leur A.
      if (/^(am\/pm|a\/p)/i.test(format.slice(i))) {
        const len = /^am\/pm/i.test(format.slice(i)) ? 5 : 3;
        out += format.slice(i, i + len);
        i += len - 1;
      } else out += 'y';
    } else if (ch === ',') {
      // Virgule : décimale entre ou après des zéros. Déjà suivie d'un groupe de 3 chiffres (« #,##0 ») : code anglais, gardée.
      out += /^,[#0?]{3}(?![#0?])/.test(format.slice(i)) && /[#0?]$/.test(out) && !/\.[0#?]*$/.test(out) && /#/.test(out) ? ',' : '.';
    } else if ((ch === ' ' || ch === ' ' || ch === ' ') && /[#0?]$/.test(out) && /^.[#0?]/.test(format.slice(i))) out += ',';
    else out += ch;
  }
  return out;
}

/** Critère de NB.SI, SOMME.SI… : « >=12,5 » → « >=12.5 », « 1 000 » → « 1000 » (le moteur ne lit que le point). */
export function frenchCriteria(criteria: string): string {
  const m = /^(<=|>=|<>|=|<|>)?(.*)$/s.exec(criteria)!;
  const n = parseFrenchNumber(m[2]!);
  if (n !== null && /[,\s  %€]/.test(m[2]!)) return `${m[1] ?? ''}${n}`;
  const d = parseFrenchDate(m[2]!);
  return d !== null ? `${m[1] ?? ''}${d}` : criteria;
}

/** Position (à partir de 1) pour CHERCHE : sans tenir compte de la casse, avec les jokers ? * et ~ ; 0 si absent. */
export function searchWithWildcards(find: string, within: string, start = 1): number {
  const source = find
    .toLocaleUpperCase()
    .replace(/~([?*~])|([?*])|([.+^${}()|[\]\\])/g, (_m, escaped: string, wild: string, special: string) => (escaped ? `\\${escaped}` : wild ? (wild === '?' ? '[\\s\\S]' : '[\\s\\S]*?') : `\\${special}`));
  const re = new RegExp(source, 'g');
  re.lastIndex = start - 1;
  const m = re.exec(within.toLocaleUpperCase());
  return m ? m.index + 1 : 0;
}

/** CTXT / FIXED : nombre arrondi en texte à la française (espace insécable fine des milliers, comme le format « # ##0,00 »). */
export function fixedText(n: number, decimals = 2, noGrouping = false): string {
  const factor = 10 ** decimals;
  const rounded = Math.round(Math.abs(n) * factor + 1e-9) / factor;
  const [int, dec] = rounded.toFixed(Math.max(0, decimals)).split('.');
  const grouped = noGrouping ? int! : int!.replace(/\B(?=(\d{3})+$)/g, ' ');
  return `${n < 0 && rounded !== 0 ? '-' : ''}${grouped}${dec ? `,${dec}` : ''}`;
}

/** AMORDEGRC : amortissement dégressif à la française (méthode d'Excel). */
export function amorDegrc(cost: number, purchased: number, firstPeriod: number, salvage: number, period: number, rate: number, basis = 0): number {
  const life = 1 / rate;
  const coef = life < 3 ? 1 : life < 5 ? 1.5 : life <= 6 ? 2 : 2.5;
  const r = rate * coef;
  const yearFrac = (a: number, b: number) => (basis === 1 ? (b - a) / 365.25 : basis === 3 ? (b - a) / 365 : basis === 2 ? (b - a) / 360 : days360(a, b) / 360);
  let n = Math.round(yearFrac(purchased, firstPeriod) * r * cost);
  let remaining = cost - n;
  let rest = remaining - salvage;
  for (let i = 0; i < period; i++) {
    n = Math.round(r * remaining);
    rest -= n;
    if (rest < 0) return period - i <= 1 ? Math.round(remaining * 0.5) : 0;
    remaining -= n;
  }
  return n;
}

function days360(a: number, b: number): number {
  const date = (s: number) => new Date(Date.UTC(1899, 11, 30) + s * 86_400_000);
  const d1 = date(a);
  const d2 = date(b);
  const day1 = Math.min(d1.getUTCDate(), 30);
  const day2 = d2.getUTCDate() === 31 && day1 === 30 ? 30 : d2.getUTCDate();
  return (d2.getUTCFullYear() - d1.getUTCFullYear()) * 360 + (d2.getUTCMonth() - d1.getUTCMonth()) * 30 + day2 - day1;
}
