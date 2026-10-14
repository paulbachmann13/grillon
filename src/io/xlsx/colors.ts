/**
 * Couleurs : ARGB d'Excel, couleurs de thème (avec teinte), palette indexée
 * historique, et les formats CSS qu'Univer peut produire (#rrggbb, rgb(), rgba()).
 */

export interface ExcelColor {
  argb?: string;
  theme?: number;
  tint?: number;
  indexed?: number;
}

/** Palette du thème Office par défaut, dans l'ordre des index de thème Excel. */
export const DEFAULT_THEME_COLORS = [
  'FFFFFF', // 0 lt1 (fond 1)
  '000000', // 1 dk1 (texte 1)
  'E7E6E6', // 2 lt2
  '44546A', // 3 dk2
  '4472C4', // 4 accent1
  'ED7D31', // 5 accent2
  'A5A5A5', // 6 accent3
  'FFC000', // 7 accent4
  '5B9BD5', // 8 accent5
  '70AD47', // 9 accent6
  '0563C1', // 10 hlink
  '954F72', // 11 folHlink
];

/** Palette indexée héritée d'Excel 97 (index 0 à 63). */
const INDEXED_COLORS = [
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080',
  '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
  '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF',
  '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
  '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696',
  '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333',
];

/**
 * Extrait la palette d'un theme1.xml. L'ordre XML est dk1, lt1, dk2, lt2,
 * accent1..6, hlink, folHlink ; Excel inverse les paires clair/foncé dans ses index.
 */
export function parseThemeColors(themeXml: string | undefined): string[] {
  if (!themeXml) return DEFAULT_THEME_COLORS;
  const scheme = /<a:clrScheme[\s\S]*?<\/a:clrScheme>/.exec(themeXml)?.[0];
  if (!scheme) return DEFAULT_THEME_COLORS;
  const tags = ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
  const found: Record<string, string> = {};
  for (const tag of tags) {
    const block = new RegExp(`<a:${tag}>([\\s\\S]*?)</a:${tag}>`).exec(scheme)?.[1] ?? '';
    const srgb = /srgbClr val="([0-9A-Fa-f]{6})"/.exec(block)?.[1];
    const sys = /lastClr="([0-9A-Fa-f]{6})"/.exec(block)?.[1];
    found[tag] = (srgb ?? sys ?? '').toUpperCase();
  }
  const order = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
  return order.map((tag, i) => found[tag] || DEFAULT_THEME_COLORS[i]);
}

function hexToRgb(hex: string): [number, number, number] {
  return [parseInt(hex.slice(0, 2), 16), parseInt(hex.slice(2, 4), 16), parseInt(hex.slice(4, 6), 16)];
}

function rgbToHex(r: number, g: number, b: number): string {
  const h = (n: number) => Math.max(0, Math.min(255, Math.round(n))).toString(16).padStart(2, '0');
  return (h(r) + h(g) + h(b)).toUpperCase();
}

function rgbToHsl(r: number, g: number, b: number): [number, number, number] {
  r /= 255; g /= 255; b /= 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return [0, 0, l];
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  let h: number;
  if (max === r) h = (g - b) / d + (g < b ? 6 : 0);
  else if (max === g) h = (b - r) / d + 2;
  else h = (r - g) / d + 4;
  return [h / 6, s, l];
}

function hslToRgb(h: number, s: number, l: number): [number, number, number] {
  if (s === 0) return [l * 255, l * 255, l * 255];
  const hue = (p: number, q: number, t: number) => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  return [hue(p, q, h + 1 / 3) * 255, hue(p, q, h) * 255, hue(p, q, h - 1 / 3) * 255];
}

/** Applique la teinte Excel (-1..1) sur la luminance HSL, comme le fait Excel. */
export function applyTint(hex: string, tint: number): string {
  if (!tint) return hex;
  const [h, s, l] = rgbToHsl(...hexToRgb(hex));
  const l2 = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
  return rgbToHex(...hslToRgb(h, s, l2));
}

/** Couleur Excel → « #RRGGBB » (ou undefined si absente/automatique). */
export function excelColorToCss(color: ExcelColor | undefined, theme: string[] = DEFAULT_THEME_COLORS): string | undefined {
  if (!color) return undefined;
  let hex: string | undefined;
  if (color.argb) {
    const argb = color.argb.toUpperCase();
    hex = argb.length === 8 ? argb.slice(2) : argb.length === 6 ? argb : undefined;
  } else if (color.theme !== undefined) {
    hex = theme[color.theme] ?? DEFAULT_THEME_COLORS[color.theme];
  } else if (color.indexed !== undefined) {
    // 64 = couleur système « automatique » : on laisse la valeur par défaut.
    hex = INDEXED_COLORS[color.indexed];
  }
  if (!hex || !/^[0-9A-F]{6}$/.test(hex)) return undefined;
  if (color.tint) hex = applyTint(hex, color.tint);
  return `#${hex}`;
}

/** Couleur CSS (#rgb, #rrggbb, rgb(), rgba()) → ARGB Excel « FFRRGGBB ». */
export function cssToArgb(css: string | null | undefined | void): string | undefined {
  if (!css) return undefined;
  const s = css.trim();
  let m = /^#([0-9a-f]{3})$/i.exec(s);
  if (m) return `FF${m[1].split('').map((c) => c + c).join('').toUpperCase()}`;
  m = /^#([0-9a-f]{6})$/i.exec(s);
  if (m) return `FF${m[1].toUpperCase()}`;
  m = /^#([0-9a-f]{8})$/i.exec(s);
  if (m) return `${m[1].slice(6).toUpperCase()}${m[1].slice(0, 6).toUpperCase()}`;
  m = /^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*(?:,\s*([\d.]+)\s*)?\)$/i.exec(s);
  if (m) {
    const a = m[4] === undefined ? 255 : Math.round(Number(m[4]) * 255);
    return rgbToHex(a, 0, 0).slice(0, 2) + rgbToHex(Number(m[1]), Number(m[2]), Number(m[3]));
  }
  return undefined;
}
