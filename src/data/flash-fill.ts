/**
 * Remplissage instantané (Ctrl+E), version simple.
 *
 * À partir de quelques exemples saisis à côté des données, on cherche un « programme »
 * qui les reproduit tous : une suite de morceaux, chacun étant soit un texte constant,
 * soit un extrait d'une colonne source (valeur entière ou n-ième mot selon un séparateur),
 * éventuellement transformé (MAJUSCULES, minuscules, Nom Propre, initiale).
 */

export type Transform = 'none' | 'upper' | 'lower' | 'proper' | 'initial';

export type Part =
  | { kind: 'const'; text: string }
  | { kind: 'extract'; col: number; sep: string | null; index: number; transform: Transform };

export type Program = Part[];

const SEPARATORS = [' ', ',', ';', '-', '_', '.', '@', '/', '|'];
const TRANSFORMS: Transform[] = ['none', 'proper', 'upper', 'lower', 'initial'];

function applyTransform(s: string, t: Transform): string {
  switch (t) {
    case 'upper':
      return s.toLocaleUpperCase('fr');
    case 'lower':
      return s.toLocaleLowerCase('fr');
    case 'proper':
      return s.toLocaleLowerCase('fr').replace(/(^|[\s'’-])(\p{L})/gu, (_, a: string, b: string) => a + b.toLocaleUpperCase('fr'));
    case 'initial':
      return s.slice(0, 1).toLocaleUpperCase('fr');
    default:
      return s;
  }
}

function tokens(value: string, sep: string): string[] {
  return value.split(sep).filter((t) => t !== '');
}

function extract(sources: string[], p: Extract<Part, { kind: 'extract' }>): string | null {
  const value = sources[p.col] ?? '';
  let raw: string | undefined;
  if (p.sep === null) raw = value;
  else {
    const toks = tokens(value, p.sep);
    raw = p.index >= 0 ? toks[p.index] : toks[toks.length + p.index];
  }
  if (raw === undefined || raw === '') return null;
  return applyTransform(raw, p.transform);
}

export function runProgram(program: Program, sources: string[]): string | null {
  let out = '';
  for (const p of program) {
    if (p.kind === 'const') out += p.text;
    else {
      const v = extract(sources, p);
      if (v === null) return null;
      out += v;
    }
  }
  return out;
}

/** Tous les extraits possibles d'une ligne source, avec leur valeur. */
function candidateParts(sources: string[]): { part: Extract<Part, { kind: 'extract' }>; value: string }[] {
  const out: { part: Extract<Part, { kind: 'extract' }>; value: string }[] = [];
  sources.forEach((value, col) => {
    if (!value) return;
    const specs: { sep: string | null; index: number }[] = [{ sep: null, index: 0 }];
    for (const sep of SEPARATORS) {
      const toks = tokens(value, sep);
      if (toks.length < 2) continue;
      for (let i = 0; i < toks.length; i++) {
        specs.push({ sep, index: i });
        specs.push({ sep, index: i - toks.length });
      }
    }
    for (const spec of specs) {
      for (const transform of TRANSFORMS) {
        const part = { kind: 'extract' as const, col, ...spec, transform };
        const v = extract(sources, part);
        if (v) out.push({ part, value: v });
      }
    }
  });
  return out;
}

/** Programmes qui produisent `target` à partir de `sources` (les plus simples d'abord). */
function programsFor(sources: string[], target: string, limit = 4000): Program[] {
  const candidates = candidateParts(sources);
  const results: Program[] = [];
  const dfs = (pos: number, acc: Program) => {
    if (results.length >= limit) return;
    if (pos === target.length) {
      results.push(acc);
      return;
    }
    // Extraits qui correspondent au texte restant (les plus longs d'abord).
    const matching = candidates.filter((c) => target.startsWith(c.value, pos)).sort((a, b) => b.value.length - a.value.length);
    for (const c of matching) dfs(pos + c.value.length, [...acc, c.part]);
    // Sinon, un caractère constant (fusionné avec la constante précédente).
    const last = acc[acc.length - 1];
    const ch = target[pos];
    if (last?.kind === 'const') dfs(pos + 1, [...acc.slice(0, -1), { kind: 'const', text: last.text + ch }]);
    else dfs(pos + 1, [...acc, { kind: 'const', text: ch }]);
  };
  dfs(0, []);
  return results;
}

/** Plus un programme a de texte constant, moins il est « général ». */
function cost(p: Program): number {
  return p.reduce((s, part) => s + (part.kind === 'const' ? 10 * part.text.length : part.transform === 'none' ? 1 : 2), 0);
}

export interface FlashFillExample {
  sources: string[];
  target: string;
}

/** Cherche un programme cohérent avec tous les exemples (null si aucun). */
export function inferProgram(examples: FlashFillExample[]): Program | null {
  const usable = examples.filter((e) => e.target !== '');
  if (!usable.length) return null;
  const [first, ...rest] = usable;
  const programs = programsFor(first.sources, first.target)
    .filter((p) => p.some((part) => part.kind === 'extract'))
    .sort((a, b) => cost(a) - cost(b));
  return programs.find((p) => rest.every((e) => runProgram(p, e.sources) === e.target)) ?? null;
}

/**
 * Remplit une colonne : `rows[i].sources` sont les valeurs des colonnes à gauche,
 * `rows[i].target` la valeur déjà saisie (exemple) ou ''. Renvoie les valeurs proposées
 * pour les lignes vides (null si le motif n'a pas été trouvé).
 */
export function flashFill(rows: FlashFillExample[]): (string | null)[] | null {
  const program = inferProgram(rows.filter((r) => r.target !== ''));
  if (!program) return null;
  return rows.map((r) => (r.target !== '' ? r.target : r.sources.some(Boolean) ? runProgram(program, r.sources) : null));
}
