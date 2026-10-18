/**
 * Thèmes du classeur (Mise en page → Thèmes, Couleurs, Polices), comme Excel.
 *
 * Univer n'a pas de notion de thème : les couleurs d'un fichier Excel sont résolues en #RRGGBB à
 * l'ouverture. Changer de thème recolore donc les cellules dont la couleur est une couleur du
 * thème actuel (avec ses teintes standard) ; les autres couleurs ne bougent pas.
 * Logique pure, sans DOM ni Univer.
 */
import { applyTint } from '../io/xlsx/colors';

/** Ordre des index de thème d'Excel : lt1, dk1, lt2, dk2, accent1..6, hlink, folHlink. */
export const THEME_SLOTS = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'] as const;

export interface ThemeColors {
  name: string;
  /** 12 couleurs « RRGGBB » dans l'ordre de THEME_SLOTS. */
  colors: string[];
}

export interface ThemeFonts {
  name: string;
  major: string;
  minor: string;
}

export interface Theme {
  name: string;
  colors: ThemeColors;
  fonts: ThemeFonts;
}

export const COLOR_SCHEMES: ThemeColors[] = [
  { name: 'Office 2013 – 2022', colors: ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '0563C1', '954F72'] },
  { name: 'Office', colors: ['FFFFFF', '000000', 'E8E8E8', '0E2841', '156082', 'E97132', '196B24', '0F9ED5', 'A02B93', '4EA72E', '467886', '96607D'] },
  { name: 'Office 2007 – 2010', colors: ['FFFFFF', '000000', 'EEECE1', '1F497D', '4F81BD', 'C0504D', '9BBB59', '8064A2', '4BACC6', 'F79646', '0000FF', '800080'] },
  { name: 'Océan', colors: ['FFFFFF', '000000', 'E3EEF2', '12394D', '0F6E8C', '2BA3B8', '5CC4B0', 'F2B544', 'E8705B', '7A8DB8', '0F6E8C', '6B5B95'] },
  { name: 'Forêt', colors: ['FFFFFF', '000000', 'ECEFE3', '2F3E2C', '3E7C4A', '8DAA3B', 'C9A227', 'A5562F', '5E8C8A', '7D6B56', '2E6B8A', '7D5A7B'] },
  { name: 'Crépuscule', colors: ['FFFFFF', '000000', 'F2E9E4', '3A2E4A', '8E4585', 'E4572E', 'F3A712', '29A5A0', '5B6EAE', 'B8587A', '5B6EAE', '8E4585'] },
  { name: 'Terre', colors: ['FFFFFF', '000000', 'EFE6DA', '4A3B2E', 'B5651D', '8C6D46', 'D4A94E', '6F8F72', 'A44A3F', '5B7C99', '2F6D8F', '8B5E83'] },
  { name: 'Monochrome', colors: ['FFFFFF', '000000', 'EDEDED', '262626', '404040', '595959', '737373', '8C8C8C', 'A6A6A6', 'BFBFBF', '404040', '737373'] },
];

export const FONT_SCHEMES: ThemeFonts[] = [
  { name: 'Office 2013 – 2022', major: 'Calibri Light', minor: 'Calibri' },
  { name: 'Office', major: 'Aptos Display', minor: 'Aptos' },
  { name: 'Office 2007 – 2010', major: 'Cambria', minor: 'Calibri' },
  { name: 'Arial', major: 'Arial', minor: 'Arial' },
  { name: 'Cambria', major: 'Cambria', minor: 'Cambria' },
  { name: 'Candara', major: 'Candara', minor: 'Candara' },
  { name: 'Georgia', major: 'Georgia', minor: 'Georgia' },
  { name: 'Georgia – Verdana', major: 'Georgia', minor: 'Verdana' },
  { name: 'Times New Roman', major: 'Times New Roman', minor: 'Times New Roman' },
  { name: 'Trebuchet MS', major: 'Trebuchet MS', minor: 'Trebuchet MS' },
  { name: 'Verdana', major: 'Verdana', minor: 'Verdana' },
];

/** Thèmes complets proposés dans la galerie « Thèmes » (couleurs + polices). */
export const THEMES: Theme[] = [
  { name: 'Office 2013 – 2022', colors: COLOR_SCHEMES[0], fonts: FONT_SCHEMES[0] },
  { name: 'Office', colors: COLOR_SCHEMES[1], fonts: FONT_SCHEMES[1] },
  { name: 'Office 2007 – 2010', colors: COLOR_SCHEMES[2], fonts: FONT_SCHEMES[2] },
  { name: 'Océan', colors: COLOR_SCHEMES[3], fonts: FONT_SCHEMES[3] },
  { name: 'Forêt', colors: COLOR_SCHEMES[4], fonts: FONT_SCHEMES[4] },
  { name: 'Crépuscule', colors: COLOR_SCHEMES[5], fonts: FONT_SCHEMES[6] },
  { name: 'Terre', colors: COLOR_SCHEMES[6], fonts: FONT_SCHEMES[5] },
  { name: 'Monochrome', colors: COLOR_SCHEMES[7], fonts: FONT_SCHEMES[8] },
];

/** Thème d'un classeur qui n'en déclare pas (celui d'Excel jusqu'en 2022). */
export const DEFAULT_THEME: Theme = THEMES[0];

export const sameColors = (a: ThemeColors, b: ThemeColors) => a.colors.every((c, i) => c.toUpperCase() === b.colors[i]?.toUpperCase());
export const sameFonts = (a: ThemeFonts, b: ThemeFonts) => a.major === b.major && a.minor === b.minor;
export const sameTheme = (a: Theme, b: Theme) => sameColors(a.colors, b.colors) && sameFonts(a.fonts, b.fonts);

// --- Recoloration ---

/** Teintes standard des galeries de couleurs d'Excel (plus claires, plus foncées). */
const TINTS = [0.8, 0.6, 0.4, -0.25, -0.5, 0.9, 0.35, 0.25, 0.15, 0.05, -0.05, -0.15, -0.35, -0.1, -0.75];
/** Les index 0 et 1 (blanc, noir) ne sont reconnus qu'à l'identique : leurs teintes sont des gris. */
const TINTED_SLOTS = [2, 3, 4, 5, 6, 7, 8, 9];

export interface ColorMatch {
  slot: number;
  tint: number;
}

/** Table « couleur → (index de thème, teinte) » d'un jeu de couleurs. L'identique l'emporte sur une teinte. */
export function colorLookup(colors: ThemeColors): Map<string, ColorMatch> {
  const map = new Map<string, ColorMatch>();
  colors.colors.forEach((hex, slot) => {
    const key = hex.toUpperCase();
    if (!map.has(key)) map.set(key, { slot, tint: 0 });
  });
  for (const tint of TINTS) {
    for (const slot of TINTED_SLOTS) {
      const key = applyTint(colors.colors[slot].toUpperCase(), tint);
      if (!map.has(key)) map.set(key, { slot, tint });
    }
  }
  return map;
}

/** Recolore « #RRGGBB » d'un thème à l'autre ; renvoie la couleur inchangée si elle n'appartient pas au thème. */
export function remapColor(css: string, lookup: Map<string, ColorMatch>, to: ThemeColors): string {
  const m = /^#([0-9a-f]{6})$/i.exec(css.trim());
  if (!m) return css;
  const hit = lookup.get(m[1].toUpperCase());
  if (!hit) return css;
  return `#${applyTint(to.colors[hit.slot].toUpperCase(), hit.tint)}`;
}

type Rgb = { rgb?: string | null } | null | undefined;
interface StyleLike {
  bg?: Rgb;
  cl?: Rgb;
  ul?: { cl?: Rgb } | null;
  st?: { cl?: Rgb } | null;
  ol?: { cl?: Rgb } | null;
  bd?: Record<string, { cl?: Rgb } | null | undefined> | null;
  ff?: string | null;
  [key: string]: unknown;
}

/** Style Univer recoloré / re-poli. Renvoie le même objet s'il n'y a rien à changer. */
export function remapStyle<T extends StyleLike>(style: T, from: Theme, to: Theme, lookup = colorLookup(from.colors)): T {
  let out: T = style;
  const set = (patch: Partial<StyleLike>) => {
    if (out === style) out = { ...style };
    Object.assign(out, patch);
  };
  const color = (c: Rgb): Rgb => {
    if (!c?.rgb) return c;
    const next = remapColor(c.rgb, lookup, to.colors);
    return next === c.rgb ? c : { ...c, rgb: next };
  };
  for (const key of ['bg', 'cl'] as const) {
    const next = color(style[key]);
    if (next !== style[key]) set({ [key]: next });
  }
  for (const key of ['ul', 'st', 'ol'] as const) {
    const sub = style[key];
    if (sub?.cl) {
      const next = color(sub.cl);
      if (next !== sub.cl) set({ [key]: { ...sub, cl: next } });
    }
  }
  if (style.bd) {
    let bd: Record<string, { cl?: Rgb } | null | undefined> | undefined;
    for (const [side, border] of Object.entries(style.bd)) {
      const next = border?.cl ? color(border.cl) : undefined;
      if (border && next && next !== border.cl) (bd ??= { ...style.bd })[side] = { ...border, cl: next };
    }
    if (bd) set({ bd });
  }
  if (style.ff) {
    const ff = remapFont(style.ff, from.fonts, to.fonts);
    if (ff !== style.ff) set({ ff });
  }
  return out;
}

/** La police du corps devient la nouvelle police du corps, idem pour les titres. */
export function remapFont(name: string, from: ThemeFonts, to: ThemeFonts): string {
  if (name === from.minor) return to.minor;
  if (name === from.major) return to.major;
  return name;
}

// --- Application à un classeur Univer ---

interface WorkbookLike {
  styles?: Record<string, StyleLike | null | undefined>;
  defaultStyle?: StyleLike | string | null;
  sheets?: Record<string, { defaultStyle?: StyleLike | string | null; custom?: unknown } | undefined>;
  custom?: Record<string, unknown>;
}

/** Clé de `custom` du classeur où l'on garde le thème (sinon : thème Office 2013 – 2022). */
export const THEME_CUSTOM_KEY = 'theme';

export function themeOf(wb: WorkbookLike | undefined): Theme {
  const stored = (wb?.custom?.grillon as { [THEME_CUSTOM_KEY]?: Theme } | undefined)?.[THEME_CUSTOM_KEY];
  if (stored?.colors?.colors?.length === 12 && stored.fonts?.minor) return stored;
  return DEFAULT_THEME;
}

/** Même classeur avec les couleurs et polices du nouveau thème (copie ; l'original n'est pas modifié). */
export function applyThemeToWorkbook<T extends WorkbookLike>(wb: T, to: Theme): T & { custom: Record<string, unknown> } {
  const from = themeOf(wb);
  const lookup = colorLookup(from.colors);
  const out = { ...wb } as T & { custom: Record<string, unknown> };
  if (wb.styles) {
    out.styles = Object.fromEntries(Object.entries(wb.styles).map(([id, s]) => [id, s ? remapStyle(s, from, to, lookup) : s]));
  }
  const font = (s: StyleLike | string | null | undefined) => (s && typeof s === 'object' ? remapStyle(s, from, to, lookup) : s);
  if (wb.defaultStyle) out.defaultStyle = font(wb.defaultStyle);
  if (wb.sheets) {
    out.sheets = Object.fromEntries(Object.entries(wb.sheets).map(([id, sh]) => [id, sh && sh.defaultStyle ? { ...sh, defaultStyle: font(sh.defaultStyle) } : sh]));
  }
  const grillon = { ...((wb.custom?.grillon as Record<string, unknown> | undefined) ?? {}) };
  if (sameTheme(to, DEFAULT_THEME)) delete grillon[THEME_CUSTOM_KEY];
  else grillon[THEME_CUSTOM_KEY] = to;
  out.custom = { ...(wb.custom ?? {}), grillon };
  return out;
}

// --- Fichier .xlsx (xl/theme/theme1.xml) ---

const escape = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

function clrSchemeXml(c: ThemeColors): string {
  const by: Record<string, string> = {};
  THEME_SLOTS.forEach((slot, i) => (by[slot] = c.colors[i].toUpperCase()));
  const order = ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
  return `<a:clrScheme name="${escape(c.name)}">${order.map((tag) => `<a:${tag}><a:srgbClr val="${by[tag]}"/></a:${tag}>`).join('')}</a:clrScheme>`;
}

/** Remplace le jeu de couleurs et les polices latines d'un theme1.xml (le reste du thème est conservé). */
export function patchThemeXml(xml: string, theme: Theme): string {
  let out = /<a:clrScheme[\s\S]*?<\/a:clrScheme>/.test(xml) ? xml.replace(/<a:clrScheme[\s\S]*?<\/a:clrScheme>/, () => clrSchemeXml(theme.colors)) : xml;
  out = out.replace(/<a:fontScheme\b([^>]*)\bname="[^"]*"/, (_m, rest) => `<a:fontScheme${rest}name="${escape(theme.fonts.name)}"`);
  for (const [tag, face] of [['majorFont', theme.fonts.major], ['minorFont', theme.fonts.minor]] as const) {
    out = out.replace(new RegExp(`(<a:${tag}>\\s*)<a:latin\\b[^>]*?/>`), (_m, head) => `${head}<a:latin typeface="${escape(face)}"/>`);
  }
  return out;
}

/** Thème lu dans un theme1.xml, ou undefined s'il n'y en a pas. */
export function parseTheme(xml: string | undefined, colors: string[]): Theme | undefined {
  if (!xml || colors.length !== 12) return undefined;
  const attr = (re: RegExp) => re.exec(xml)?.[1];
  const decode = (s: string) => s.replace(/&quot;/g, '"').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');
  const clrName = decode(attr(/<a:clrScheme[^>]*\bname="([^"]*)"/) ?? 'Personnalisé');
  const fontName = decode(attr(/<a:fontScheme[^>]*\bname="([^"]*)"/) ?? 'Personnalisé');
  const major = decode(attr(/<a:majorFont>\s*<a:latin[^>]*\btypeface="([^"]*)"/) ?? DEFAULT_THEME.fonts.major);
  const minor = decode(attr(/<a:minorFont>\s*<a:latin[^>]*\btypeface="([^"]*)"/) ?? DEFAULT_THEME.fonts.minor);
  const named = (list: ThemeColors[]) => list.find((s) => s.colors.every((c, i) => c === colors[i].toUpperCase()));
  const fonts = FONT_SCHEMES.find((f) => f.major === major && f.minor === minor);
  return {
    name: clrName,
    colors: { name: named(COLOR_SCHEMES)?.name ?? clrName, colors: colors.map((c) => c.toUpperCase()) },
    fonts: fonts ?? { name: fontName, major, minor },
  };
}
