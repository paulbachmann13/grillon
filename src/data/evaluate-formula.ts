/**
 * « Évaluer la formule » (Formules → Audit de formules), comme Excel : la formule est calculée pas
 * à pas, de gauche à droite. Chaque étape remplace la partie soulignée par sa valeur : d'abord les
 * références (A1, B2:B10, noms), puis les appels de fonction dont les arguments sont des valeurs,
 * enfin les opérations restantes.
 */

export interface Token {
  kind: 'ref' | 'fn' | 'lit' | 'op';
  start: number;
  end: number; // exclusif
}

const ERRORS = /^#(DIV\/0!|N\/A|NAME\?|NULL!|NUM!|REF!|VALUE!|SPILL!|CALC!)/i;
const SHEET = String.raw`(?:'(?:[^']|'')+'|[A-Za-z_À-ɏ][\w.À-ɏ]*)!`;
const CELL = String.raw`\$?[A-Za-z]{1,3}\$?\d+`;
const CELL_REF = new RegExp(String.raw`^(?:${SHEET})?${CELL}(?::${CELL})?`);
const COL_REF = new RegExp(String.raw`^(?:${SHEET})?\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}(?![\w(])`);
const ROW_REF = new RegExp(String.raw`^(?:${SHEET})?\$?\d+:\$?\d+`);
const NUMBER = /^(\d+(\.\d*)?|\.\d+)(E[-+]?\d+)?/i;
const IDENT = /^[A-Za-z_\\À-ɏ][\w.À-ɏ]*/;

/** Découpe une formule (sans « = ») en éléments. */
export function tokenize(expr: string): Token[] {
  const out: Token[] = [];
  let i = 0;
  while (i < expr.length) {
    const ch = expr[i];
    const rest = expr.slice(i);
    if (/\s/.test(ch)) {
      i++;
      continue;
    }
    if (ch === '"') {
      let j = i + 1;
      while (j < expr.length && !(expr[j] === '"' && expr[j + 1] !== '"')) j += expr[j] === '"' ? 2 : 1;
      out.push({ kind: 'lit', start: i, end: j + 1 });
      i = j + 1;
      continue;
    }
    if (ch === '{') {
      const j = expr.indexOf('}', i);
      const end = j < 0 ? expr.length : j + 1;
      out.push({ kind: 'lit', start: i, end });
      i = end;
      continue;
    }
    const err = ERRORS.exec(rest);
    if (err) {
      out.push({ kind: 'lit', start: i, end: i + err[0].length });
      i += err[0].length;
      continue;
    }
    const ref = CELL_REF.exec(rest) ?? COL_REF.exec(rest) ?? (/^\d/.test(ch) || ch === "'" ? ROW_REF.exec(rest) : null);
    if (ref && !/^[\w(]/.test(expr.slice(i + ref[0].length))) {
      out.push({ kind: 'ref', start: i, end: i + ref[0].length });
      i += ref[0].length;
      continue;
    }
    const num = NUMBER.exec(rest);
    if (num) {
      out.push({ kind: 'lit', start: i, end: i + num[0].length });
      i += num[0].length;
      continue;
    }
    // Nom de feuille entre apostrophes devant un nom défini : 'Ma feuille'!Total
    const quoted = /^'(?:[^']|'')+'!/.exec(rest);
    const ident = quoted ? null : IDENT.exec(rest);
    if (quoted || ident) {
      let end = i + (quoted ?? ident)![0].length;
      if (quoted) end += IDENT.exec(expr.slice(end))?.[0].length ?? 0;
      else if (expr[end] === '!') end += 1 + (IDENT.exec(expr.slice(end + 1))?.[0].length ?? 0);
      const word = expr.slice(i, end);
      if (expr[end] === '(') out.push({ kind: 'fn', start: i, end });
      else if (expr[end] === '[') {
        // Référence structurée : Tableau1[Prix], Tableau1[[#Totaux],[Prix]]
        let depth = 0;
        let j = end;
        for (; j < expr.length; j++) {
          if (expr[j] === '[') depth++;
          else if (expr[j] === ']' && --depth === 0) break;
        }
        out.push({ kind: 'ref', start: i, end: j + 1 });
        end = j + 1;
      } else out.push({ kind: /^(TRUE|FALSE)$/i.test(word) ? 'lit' : 'ref', start: i, end });
      i = end;
      continue;
    }
    out.push({ kind: 'op', start: i, end: i + 1 });
    i++;
  }
  return out;
}

/** Position de la parenthèse fermante qui correspond à celle ouverte en `open`. */
function closingParen(expr: string, tokens: Token[], open: number): number {
  let depth = 0;
  for (const tk of tokens) {
    if (tk.start < open || tk.kind !== 'op') continue;
    const ch = expr[tk.start];
    if (ch === '(') depth++;
    else if (ch === ')' && --depth === 0) return tk.start;
  }
  return expr.length - 1;
}

/**
 * Partie à évaluer à l'étape suivante, ou null si l'expression est déjà une valeur.
 * Renvoie [début, fin) dans `expr`.
 */
export function nextStep(expr: string): [number, number] | null {
  const tokens = tokenize(expr);
  for (const tk of tokens) {
    if (tk.kind === 'ref') return [tk.start, tk.end];
    if (tk.kind === 'fn') {
      const close = closingParen(expr, tokens, tk.end);
      const inner = tokens.filter((x) => x.start > tk.end && x.start < close);
      if (!inner.some((x) => x.kind === 'ref' || x.kind === 'fn')) return [tk.start, close + 1];
    }
  }
  const meaningful = tokens.filter((x) => !(x.kind === 'op' && /[()]/.test(expr[x.start])));
  if (meaningful.length <= 1 && meaningful[0]?.kind !== 'ref') {
    // Un seul littéral, éventuellement entre parenthèses ou précédé d'un signe : terminé.
    const signOnly = meaningful.length === 1 && tokens.every((x) => x.kind !== 'op' || /[()]/.test(expr[x.start]));
    if (signOnly || meaningful.length === 0) return null;
  }
  // Signe moins devant un nombre : « -5 » est déjà une valeur.
  if (/^\s*-?\s*(\d+(\.\d*)?|\.\d+)(E[-+]?\d+)?\s*$/i.test(expr)) return null;
  return [0, expr.length];
}
