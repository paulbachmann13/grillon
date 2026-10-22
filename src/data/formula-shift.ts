/**
 * Décalage des références d'une formule, comme quand on recopie une formule vers le bas ou vers la droite :
 * les références relatives (A1) bougent, les références absolues ($A$1) restent. Sert à retrouver la formule d'une
 * cellule « partagée » d'Univer (une cellule remplie ne garde que l'identifiant de la formule de la première).
 * Logique pure, sans DOM ni Univer.
 */
const MAX_ROW = 1_048_576;
const MAX_COL = 16_384;

function colToNumber(letters: string): number {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + ch.charCodeAt(0) - 64;
  return n;
}

function numberToCol(n: number): string {
  let s = '';
  for (let x = n; x > 0; x = Math.floor((x - 1) / 26)) s = String.fromCharCode(65 + ((x - 1) % 26)) + s;
  return s;
}

/** Colonne (« $B » ou « b ») décalée ; null si elle sort de la feuille. */
function shiftCol(ref: string, d: number): string | null {
  const absolute = ref.startsWith('$');
  const n = colToNumber(absolute ? ref.slice(1) : ref);
  if (absolute) return ref;
  const next = n + d;
  return next >= 1 && next <= MAX_COL ? numberToCol(next) : null;
}

function shiftRow(ref: string, d: number): string | null {
  if (ref.startsWith('$')) return ref;
  const next = Number(ref) + d;
  return next >= 1 && next <= MAX_ROW ? String(next) : null;
}

const CELL = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)/;
const COL_RANGE = /^(\$?[A-Za-z]{1,3}):(\$?[A-Za-z]{1,3})(?![A-Za-z0-9_(.])/;
const ROW_RANGE = /^(\$?\d+):(\$?\d+)(?![A-Za-z0-9_.(])/;

/** Vrai si « A1 », « $A$1 »… est une vraie cellule (colonne ≤ XFD, ligne ≥ 1). */
function isCell(m: RegExpExecArray): boolean {
  return colToNumber(m[2]) <= MAX_COL && Number(m[4]) >= 1 && Number(m[4]) <= MAX_ROW;
}

/**
 * Formule recopiée de `dRow` lignes et `dCol` colonnes. Une référence qui sortirait de la feuille devient #REF!.
 * Les textes entre guillemets, les noms de feuille entre apostrophes, les références de tableau [..] et les noms de
 * fonctions ne sont pas touchés.
 */
export function shiftFormula(formula: string, dRow: number, dCol: number): string {
  if (!dRow && !dCol) return formula;
  let out = '';
  let i = 0;
  while (i < formula.length) {
    const c = formula[i];
    if (c === '"') {
      let j = i + 1;
      while (j < formula.length && (formula[j] !== '"' || formula[j + 1] === '"')) j += formula[j] === '"' ? 2 : 1;
      out += formula.slice(i, j + 1);
      i = j + 1;
      continue;
    }
    if (c === "'") {
      let j = i + 1;
      while (j < formula.length && (formula[j] !== "'" || formula[j + 1] === "'")) j += formula[j] === "'" ? 2 : 1;
      out += formula.slice(i, j + 1);
      i = j + 1;
      continue;
    }
    if (c === '[') {
      let depth = 0;
      let j = i;
      for (; j < formula.length; j++) {
        if (formula[j] === '[') depth++;
        else if (formula[j] === ']' && --depth === 0) break;
      }
      out += formula.slice(i, j + 1);
      i = j + 1;
      continue;
    }
    const prev = i > 0 ? formula[i - 1] : '';
    // Début d'un mot seulement : on ne coupe pas un nom au milieu (« TVA1 », « A.B1 »).
    if (!/[A-Za-z0-9_.]/.test(prev)) {
      const rest = formula.slice(i);
      let m: RegExpExecArray | null;
      if ((m = COL_RANGE.exec(rest))) {
        const a = shiftCol(m[1], dCol);
        const b = shiftCol(m[2], dCol);
        out += a && b ? `${a}:${b}` : '#REF!';
        i += m[0].length;
        continue;
      }
      if ((m = ROW_RANGE.exec(rest))) {
        const a = shiftRow(m[1], dRow);
        const b = shiftRow(m[2], dRow);
        out += a && b ? `${a}:${b}` : '#REF!';
        i += m[0].length;
        continue;
      }
      if ((m = CELL.exec(rest)) && isCell(m)) {
        const after = rest[m[0].length] ?? '';
        // « Feuil1!A1 » : le nom avant « ! » n'est pas une cellule ; une fonction « LOG10( » non plus.
        if (after !== '!' && after !== '(' && !/[A-Za-z0-9_.]/.test(after)) {
          const col = shiftCol(`${m[1]}${m[2]}`, dCol);
          const row = shiftRow(`${m[3]}${m[4]}`, dRow);
          out += col && row ? `${col}${row}` : '#REF!';
          i += m[0].length;
          continue;
        }
      }
    }
    // Mot ordinaire (nom de fonction, de feuille, nombre…) : copié en bloc pour ne pas y chercher de référence.
    const word = /^[A-Za-z0-9_.]+/.exec(formula.slice(i));
    if (word) {
      out += word[0];
      i += word[0].length;
      continue;
    }
    out += c;
    i++;
  }
  return out;
}
