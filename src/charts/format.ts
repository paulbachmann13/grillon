/**
 * Format de nombre « à la Excel » pour les étiquettes et les axes des graphiques :
 * 0, 0.00, #,##0, 0%, 0.0E+00, textes entre guillemets, [$€-40C]… Rendu à la française
 * (virgule décimale, espace insécable comme séparateur de milliers).
 */
const NBSP = ' ';

/** Nombre au format Standard, à la française (comme les cellules). */
export function formatGeneral(v: number): string {
  if (!Number.isFinite(v)) return String(v);
  const abs = Math.abs(v);
  if (abs !== 0 && (abs >= 1e11 || abs < 1e-9)) return v.toExponential(4).replace('.', ',').replace(/e([+-])(\d)$/, 'E$10$2').replace(/e/, 'E');
  const text = String(Number(v.toPrecision(11)));
  const [int, dec] = text.split('.');
  const grouped = int!.replace(/\B(?=(\d{3})+(?!\d))/g, NBSP);
  return dec ? `${grouped},${dec}` : grouped;
}

function group(int: string): string {
  return int.replace(/\B(?=(\d{3})+(?!\d))/g, NBSP);
}

export function formatNumberPattern(value: number, pattern?: string): string {
  if (!pattern || /^general$/i.test(pattern) || /^standard$/i.test(pattern)) return formatGeneral(value);
  if (!Number.isFinite(value)) return String(value);
  const sections = pattern.split(';');
  const negativeSection = value < 0 && sections[1] !== undefined;
  let p = negativeSection ? sections[1]! : sections[0]!;
  let n = negativeSection ? -value : value;

  // Textes littéraux, devises [$€-40C], caractères échappés.
  const literals: string[] = [];
  const keep = (text: string) => `\u0000${literals.push(text) - 1}\u0000`;
  p = p
    .replace(/\[\$([^\]-]*)[^\]]*\]/g, (_, s: string) => keep(s))
    .replace(/"([^"]*)"/g, (_, s: string) => keep(s))
    .replace(/\\(.)/g, (_, s: string) => keep(s))
    .replace(/_.|\*./g, '')
    .replace(/\[[^\]]*\]/g, '');
  if (p.includes('%')) n *= 100;

  const core = /[#0?][#0?,]*(?:\.[0#?]+)?(?:E[+-]0+)?/i.exec(p);
  if (!core) return p.replace(/\u0000(\d+)\u0000/g, (_, i: string) => literals[Number(i)]!);
  const [numberPart, expPart] = core[0].split(/E/i) as [string, string | undefined];
  const [intPattern = '', decPattern = ''] = numberPart.split('.');
  const thousands = intPattern.includes(',');
  const minInt = (intPattern.match(/0/g) ?? []).length;
  const minDec = (decPattern.match(/0/g) ?? []).length;
  const maxDec = decPattern.length;

  let body: string;
  if (expPart !== undefined) {
    const [mantissa = '0', exp = '0'] = n.toExponential(maxDec).split('e');
    const e = Number(exp);
    body = `${mantissa.replace('.', ',')}E${e < 0 ? '-' : '+'}${String(Math.abs(e)).padStart(expPart.length - 1, '0')}`;
  } else {
    let [int = '0', dec = ''] = n.toFixed(maxDec).split('.');
    while (dec.length > minDec && dec.endsWith('0')) dec = dec.slice(0, -1);
    int = int.padStart(minInt, '0');
    if (int === '0' && minInt === 0 && dec) int = '';
    body = (thousands ? group(int) : int) + (dec ? `,${dec}` : '');
  }
  const out = p.slice(0, core.index) + body + p.slice(core.index + core[0].length);
  const text = out.replace(/\u0000(\d+)\u0000/g, (_, i: string) => literals[Number(i)]!);
  return negativeSection || value >= 0 ? text : `-${text}`;
}
