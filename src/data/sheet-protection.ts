/**
 * Protection de feuille au format d'Excel (élément <sheetProtection>) : mot de passe facultatif,
 * haché comme Excel 2013+ (SHA-512, sel, 100 000 itérations). Le mot de passe n'est jamais stocké.
 */
export interface SheetProtection {
  /** Hachage Excel du mot de passe (absent : protection sans mot de passe). */
  algorithmName?: string;
  hashValue?: string;
  saltValue?: string;
  spinCount?: number;
  /** Ancien hachage Excel sur 16 bits (fichiers anciens). */
  legacyPassword?: string;
}

const SPIN = 100_000;

const toBase64 = (bytes: Uint8Array) => btoa(String.fromCharCode(...bytes));
const fromBase64 = (s: string) => Uint8Array.from(atob(s), (c) => c.charCodeAt(0));

function utf16le(text: string): Uint8Array {
  const out = new Uint8Array(text.length * 2);
  for (let i = 0; i < text.length; i++) {
    const c = text.charCodeAt(i);
    out[i * 2] = c & 255;
    out[i * 2 + 1] = c >> 8;
  }
  return out;
}

const concat = (a: Uint8Array, b: Uint8Array) => {
  const out = new Uint8Array(a.length + b.length);
  out.set(a);
  out.set(b, a.length);
  return out;
};

/** Hachage d'un mot de passe selon l'algorithme d'Excel (ECMA-376, « agile »). */
export async function excelPasswordHash(password: string, salt: Uint8Array, spinCount = SPIN): Promise<string> {
  const subtle = globalThis.crypto.subtle;
  let hash = new Uint8Array(await subtle.digest('SHA-512', concat(salt, utf16le(password))));
  const counter = new Uint8Array(4);
  for (let i = 0; i < spinCount; i++) {
    counter[0] = i & 255;
    counter[1] = (i >>> 8) & 255;
    counter[2] = (i >>> 16) & 255;
    counter[3] = (i >>> 24) & 255;
    hash = new Uint8Array(await subtle.digest('SHA-512', concat(hash, counter)));
  }
  return toBase64(hash);
}

/** Protection avec (ou sans) mot de passe. */
export async function makeProtection(password: string | undefined): Promise<SheetProtection> {
  if (!password) return {};
  const salt = globalThis.crypto.getRandomValues(new Uint8Array(16));
  return { algorithmName: 'SHA-512', hashValue: await excelPasswordHash(password, salt), saltValue: toBase64(salt), spinCount: SPIN };
}

/** Ancien hachage d'Excel (16 bits) d'un mot de passe, en hexadécimal. */
export function legacyHash(password: string): string {
  let hash = 0;
  for (let i = password.length - 1; i >= 0; i--) {
    hash = ((hash >> 14) & 1) | ((hash << 1) & 0x7fff);
    hash ^= password.charCodeAt(i);
  }
  hash = ((hash >> 14) & 1) | ((hash << 1) & 0x7fff);
  hash ^= password.length;
  hash ^= 0xce4b;
  return hash.toString(16).toUpperCase();
}

/** Vérifie un mot de passe ; une protection sans mot de passe accepte tout. */
export async function checkPassword(p: SheetProtection, password: string): Promise<boolean> {
  if (p.hashValue && p.saltValue) {
    if (p.algorithmName && p.algorithmName !== 'SHA-512') return false;
    return (await excelPasswordHash(password, fromBase64(p.saltValue), p.spinCount ?? SPIN)) === p.hashValue;
  }
  if (p.legacyPassword) return legacyHash(password) === p.legacyPassword.toUpperCase();
  return true;
}

export const hasPassword = (p: SheetProtection) => !!(p.hashValue || p.legacyPassword);

/** Élément <sheetProtection> (mêmes options par défaut qu'Excel). */
export function sheetProtectionXml(p: SheetProtection): string {
  const attrs: string[] = [];
  if (p.algorithmName && p.hashValue && p.saltValue) {
    attrs.push(`algorithmName="${p.algorithmName}"`, `hashValue="${p.hashValue}"`, `saltValue="${p.saltValue}"`, `spinCount="${p.spinCount ?? SPIN}"`);
  } else if (p.legacyPassword) attrs.push(`password="${p.legacyPassword}"`);
  attrs.push('sheet="1"', 'objects="1"', 'scenarios="1"');
  return `<sheetProtection ${attrs.join(' ')}/>`;
}

/** Lit l'élément <sheetProtection> d'une feuille (null si la feuille n'est pas protégée). */
export function readSheetProtection(sheetXml: string): SheetProtection | null {
  const m = /<(?:\w+:)?sheetProtection\b([^>]*)\/?>/.exec(sheetXml);
  if (!m) return null;
  const attr = (name: string) => new RegExp(`\\b${name}="([^"]*)"`).exec(m[1]!)?.[1];
  const sheet = attr('sheet');
  if (sheet !== '1' && sheet !== 'true') return null;
  const p: SheetProtection = {};
  const hash = attr('hashValue');
  if (hash) {
    p.algorithmName = attr('algorithmName') ?? 'SHA-512';
    p.hashValue = hash;
    p.saltValue = attr('saltValue');
    p.spinCount = Number(attr('spinCount') ?? SPIN);
  }
  const legacy = attr('password');
  if (legacy) p.legacyPassword = legacy;
  return p;
}
