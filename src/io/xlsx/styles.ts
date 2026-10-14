/**
 * Styles de cellule : ExcelJS (Style) ⇄ Univer (IStyleData).
 */
import type { Alignment, Border, Borders, Fill, Font, Style } from 'exceljs';
import {
  BooleanNumber,
  BorderStyleTypes,
  HorizontalAlign,
  VerticalAlign,
  WrapStrategy,
  type IBorderData,
  type IBorderStyleData,
  type IStyleData,
} from '../../model/univer-types';
import { cssToArgb, excelColorToCss, type ExcelColor } from './colors';

type BorderStyleName = NonNullable<Border['style']>;

const BORDER_TO_UNIVER: Record<string, number> = {
  thin: BorderStyleTypes.THIN,
  hair: BorderStyleTypes.HAIR,
  dotted: BorderStyleTypes.DOTTED,
  dashed: BorderStyleTypes.DASHED,
  dashDot: BorderStyleTypes.DASH_DOT,
  dashDotDot: BorderStyleTypes.DASH_DOT_DOT,
  double: BorderStyleTypes.DOUBLE,
  medium: BorderStyleTypes.MEDIUM,
  mediumDashed: BorderStyleTypes.MEDIUM_DASHED,
  mediumDashDot: BorderStyleTypes.MEDIUM_DASH_DOT,
  mediumDashDotDot: BorderStyleTypes.MEDIUM_DASH_DOT_DOT,
  slantDashDot: BorderStyleTypes.SLANT_DASH_DOT,
  thick: BorderStyleTypes.THICK,
};

const BORDER_FROM_UNIVER: Record<number, BorderStyleName> = Object.fromEntries(
  Object.entries(BORDER_TO_UNIVER).map(([k, v]) => [v, k as BorderStyleName]),
);

const H_ALIGN_TO_UNIVER: Record<string, number> = {
  left: HorizontalAlign.LEFT,
  center: HorizontalAlign.CENTER,
  centerContinuous: HorizontalAlign.CENTER,
  right: HorizontalAlign.RIGHT,
  justify: HorizontalAlign.JUSTIFIED,
  fill: HorizontalAlign.LEFT,
  distributed: HorizontalAlign.DISTRIBUTED,
};

const H_ALIGN_FROM_UNIVER: Record<number, Alignment['horizontal']> = {
  [HorizontalAlign.LEFT]: 'left',
  [HorizontalAlign.CENTER]: 'center',
  [HorizontalAlign.RIGHT]: 'right',
  [HorizontalAlign.JUSTIFIED]: 'justify',
  [HorizontalAlign.BOTH]: 'justify',
  [HorizontalAlign.DISTRIBUTED]: 'distributed',
};

const V_ALIGN_TO_UNIVER: Record<string, number> = {
  top: VerticalAlign.TOP,
  middle: VerticalAlign.MIDDLE,
  bottom: VerticalAlign.BOTTOM,
  distributed: VerticalAlign.MIDDLE,
  justify: VerticalAlign.MIDDLE,
};

const V_ALIGN_FROM_UNIVER: Record<number, Alignment['vertical']> = {
  [VerticalAlign.TOP]: 'top',
  [VerticalAlign.MIDDLE]: 'middle',
  [VerticalAlign.BOTTOM]: 'bottom',
};

const BORDER_SIDES = [
  ['top', 't'],
  ['right', 'r'],
  ['bottom', 'b'],
  ['left', 'l'],
] as const;

function borderToUniver(b: Partial<Border> | undefined, theme: string[]): IBorderStyleData | undefined {
  if (!b?.style) return undefined;
  const s = BORDER_TO_UNIVER[b.style];
  if (!s) return undefined;
  return { s: s as IBorderStyleData['s'], cl: { rgb: excelColorToCss(b.color as ExcelColor, theme) ?? '#000000' } };
}

function borderFromUniver(b: IBorderStyleData | null | undefined | void): Partial<Border> | undefined {
  if (!b || !b.s) return undefined;
  const style = BORDER_FROM_UNIVER[b.s];
  if (!style) return undefined;
  const argb = cssToArgb(b.cl?.rgb);
  return argb ? { style, color: { argb } } : { style };
}

/** Style ExcelJS → IStyleData Univer (null si rien à garder). */
export function styleToUniver(style: Partial<Style> | undefined, theme: string[]): IStyleData | null {
  if (!style) return null;
  const out: IStyleData = {};
  const font = style.font as Partial<Font> | undefined;
  if (font) {
    if (font.name) out.ff = font.name;
    if (font.size) out.fs = font.size;
    if (font.bold) out.bl = BooleanNumber.TRUE;
    if (font.italic) out.it = BooleanNumber.TRUE;
    if (font.underline && font.underline !== 'none') out.ul = { s: BooleanNumber.TRUE };
    if (font.strike) out.st = { s: BooleanNumber.TRUE };
    const cl = excelColorToCss(font.color as ExcelColor, theme);
    if (cl) out.cl = { rgb: cl };
    if (font.vertAlign === 'superscript') out.va = 3 as IStyleData['va'];
    else if (font.vertAlign === 'subscript') out.va = 2 as IStyleData['va'];
  }

  const fill = style.fill as Fill | undefined;
  if (fill && fill.type === 'pattern' && fill.pattern && fill.pattern !== 'none') {
    // Motif plein : fgColor est la couleur du fond. Autres motifs : approximation par fgColor.
    const bg = excelColorToCss((fill.fgColor ?? fill.bgColor) as ExcelColor, theme);
    if (bg) out.bg = { rgb: bg };
  } else if (fill && fill.type === 'gradient' && fill.stops?.length) {
    const bg = excelColorToCss(fill.stops[0].color as ExcelColor, theme);
    if (bg) out.bg = { rgb: bg };
  }

  const border = style.border as Partial<Borders> | undefined;
  if (border) {
    const bd: IBorderData = {};
    for (const [xl, un] of BORDER_SIDES) {
      const b = borderToUniver(border[xl], theme);
      if (b) bd[un] = b;
    }
    if (border.diagonal?.style) {
      const b = borderToUniver(border.diagonal, theme);
      if (b) {
        if (border.diagonal.down) bd.tl_br = b;
        if (border.diagonal.up) bd.bl_tr = b;
      }
    }
    if (Object.keys(bd).length) out.bd = bd;
  }

  const al = style.alignment as Partial<Alignment> | undefined;
  if (al) {
    if (al.horizontal && H_ALIGN_TO_UNIVER[al.horizontal]) out.ht = H_ALIGN_TO_UNIVER[al.horizontal] as IStyleData['ht'];
    if (al.vertical && V_ALIGN_TO_UNIVER[al.vertical]) out.vt = V_ALIGN_TO_UNIVER[al.vertical] as IStyleData['vt'];
    if (al.wrapText) out.tb = WrapStrategy.WRAP as IStyleData['tb'];
    else if (al.shrinkToFit) out.tb = WrapStrategy.CLIP as IStyleData['tb'];
    if (al.textRotation === 'vertical') out.tr = { a: 0, v: BooleanNumber.TRUE };
    else if (typeof al.textRotation === 'number' && al.textRotation !== 0) {
      // Excel : 1..90 = sens antihoraire, 91..180 = -1..-90.
      const a = al.textRotation > 90 ? 90 - al.textRotation : al.textRotation;
      out.tr = { a };
    }
    if (al.indent) out.pd = { l: al.indent * 9 };
  }

  if (style.numFmt && style.numFmt !== 'General') out.n = { pattern: style.numFmt };

  return Object.keys(out).length ? out : null;
}

/** IStyleData Univer → style ExcelJS. */
export function styleFromUniver(s: IStyleData | null | undefined): Partial<Style> | null {
  if (!s) return null;
  const out: Partial<Style> = {};

  const font: Partial<Font> = {};
  if (s.ff) font.name = s.ff;
  if (s.fs) font.size = s.fs;
  if (s.bl) font.bold = true;
  if (s.it) font.italic = true;
  if (s.ul?.s) font.underline = true;
  if (s.st?.s) font.strike = true;
  const fontArgb = cssToArgb(s.cl?.rgb);
  if (fontArgb) font.color = { argb: fontArgb };
  if ((s.va as number) === 3) font.vertAlign = 'superscript';
  else if ((s.va as number) === 2) font.vertAlign = 'subscript';
  if (Object.keys(font).length) out.font = font;

  const bgArgb = cssToArgb(s.bg?.rgb);
  if (bgArgb) out.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgArgb }, bgColor: { indexed: 64 } as never };

  if (s.bd) {
    const border: Partial<Borders> = {};
    for (const [xl, un] of BORDER_SIDES) {
      const b = borderFromUniver(s.bd[un]);
      if (b) border[xl] = b as Border;
    }
    const down = borderFromUniver(s.bd.tl_br);
    const up = borderFromUniver(s.bd.bl_tr);
    if (down || up) border.diagonal = { ...(down ?? up), up: !!up, down: !!down } as Borders['diagonal'];
    if (Object.keys(border).length) out.border = border;
  }

  const al: Partial<Alignment> = {};
  if (s.ht && H_ALIGN_FROM_UNIVER[s.ht]) al.horizontal = H_ALIGN_FROM_UNIVER[s.ht];
  if (s.vt && V_ALIGN_FROM_UNIVER[s.vt]) al.vertical = V_ALIGN_FROM_UNIVER[s.vt];
  if (s.tb === WrapStrategy.WRAP) al.wrapText = true;
  else if (s.tb === WrapStrategy.CLIP) al.shrinkToFit = true;
  if (s.tr) {
    if (s.tr.v) al.textRotation = 'vertical';
    else if (s.tr.a) al.textRotation = s.tr.a < 0 ? 90 - s.tr.a : s.tr.a;
  }
  if (s.pd?.l) {
    const indent = Math.round(s.pd.l / 9);
    if (indent > 0) al.indent = indent;
  }
  if (Object.keys(al).length) out.alignment = al;

  if (s.n?.pattern) out.numFmt = s.n.pattern;

  return Object.keys(out).length ? out : null;
}
