/**
 * Historique des versions : lit le CHANGELOG.md (le même texte que celui des Releases GitHub, donc de la fenêtre « Mise à jour »).
 * Logique pure, sans DOM.
 */

export type NoteBlock =
  | { kind: 'heading'; text: string }
  | { kind: 'item'; text: string; depth: number }
  | { kind: 'paragraph'; text: string };

export interface ReleaseNotes {
  version: string;
  /** Date au format AAAA-MM-JJ, si le titre en donne une. */
  date?: string;
  /** Titre éventuel (« Phase 3 : onglet Données »). */
  title?: string;
  blocks: NoteBlock[];
}

export type Inline = { text: string; bold?: boolean; code?: boolean };

/** Une section « ### Corrigé » ou un paragraphe, une puce, une sous-puce : tout ce qu'écrit le CHANGELOG. */
export function parseNotes(body: string): NoteBlock[] {
  const blocks: NoteBlock[] = [];
  for (const raw of body.replace(/\r/g, '').split('\n')) {
    const line = raw.replace(/\s+$/, '');
    if (!line.trim()) continue;
    const heading = /^#{2,4}\s+(.*)$/.exec(line);
    if (heading) {
      blocks.push({ kind: 'heading', text: heading[1]!.trim() });
      continue;
    }
    const item = /^(\s*)[-*]\s+(.*)$/.exec(line);
    if (item) {
      blocks.push({ kind: 'item', text: item[2]!.trim(), depth: Math.min(2, Math.floor(item[1]!.length / 2)) });
      continue;
    }
    // Suite d'une puce sur la ligne suivante : rattachée à la puce précédente.
    const last = blocks[blocks.length - 1];
    if (last?.kind === 'item' && /^\s+\S/.test(line)) last.text += ` ${line.trim()}`;
    else blocks.push({ kind: 'paragraph', text: line.trim() });
  }
  return blocks;
}

/** Texte → morceaux (gras `**…**`, code `` `…` ``), pour l'afficher sans HTML. */
export function parseInline(text: string): Inline[] {
  const out: Inline[] = [];
  const re = /\*\*(.+?)\*\*|`([^`]+)`/g;
  let last = 0;
  for (let m = re.exec(text); m; m = re.exec(text)) {
    if (m.index > last) out.push({ text: text.slice(last, m.index) });
    out.push(m[1] !== undefined ? { text: m[1], bold: true } : { text: m[2]!, code: true });
    last = m.index + m[0].length;
  }
  if (last < text.length) out.push({ text: text.slice(last) });
  return out;
}

/** Le CHANGELOG entier, du plus récent au plus ancien (l'ordre du fichier). */
export function parseChangelog(text: string): ReleaseNotes[] {
  const parts = text.replace(/\r/g, '').split(/^## \[/m).slice(1);
  const out: ReleaseNotes[] = [];
  for (const part of parts) {
    const end = part.indexOf('\n');
    const head = end < 0 ? part : part.slice(0, end);
    const m = /^([^\]]+)\](.*)$/.exec(head);
    if (!m) continue;
    const rest = m[2]!;
    const date = /(\d{4}-\d{2}-\d{2})/.exec(rest)?.[1];
    const title = rest.replace(/\d{4}-\d{2}-\d{2}/, '').replace(/^[\s—-]+|[\s—-]+$/g, '').replace(/\s+—\s+/g, ' — ') || undefined;
    const blocks = parseNotes(end < 0 ? '' : part.slice(end + 1));
    out.push({ version: m[1]!.trim(), ...(date ? { date } : {}), ...(title ? { title } : {}), blocks });
  }
  return out;
}

const numbers = (v: string) => v.split('.').map((n) => Number.parseInt(n, 10) || 0);

/** Compare deux numéros de version (1.14.10 > 1.14.9) : négatif, nul ou positif. */
export function compareVersions(a: string, b: string): number {
  const x = numbers(a);
  const y = numbers(b);
  for (let i = 0; i < Math.max(x.length, y.length); i++) {
    const d = (x[i] ?? 0) - (y[i] ?? 0);
    if (d) return d;
  }
  return 0;
}

/**
 * Les versions à raconter après une mise à jour : plus récentes que `from` et au plus `to`, la plus récente d'abord.
 * Sans version de départ connue, seule la version actuelle.
 */
export function notesSince(all: ReleaseNotes[], from: string | null, to: string, limit = 6): ReleaseNotes[] {
  const upTo = all.filter((r) => compareVersions(r.version, to) <= 0);
  if (!from) return upTo.filter((r) => r.version === to);
  return upTo.filter((r) => compareVersions(r.version, from) > 0).sort((a, b) => compareVersions(b.version, a.version)).slice(0, limit);
}
