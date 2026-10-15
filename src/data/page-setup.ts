/**
 * Mise en page d'une feuille (logique pure) : papier, orientation, marges, échelle, zone
 * d'impression, titres répétés, en-têtes et pieds de page, avec la conversion vers / depuis
 * les réglages d'Excel (« pageSetup », « headerFooter »).
 */
export type PaperSize = 'A4' | 'A3' | 'A5' | 'Letter' | 'Legal';

export interface HeaderFooterSpec {
  left?: string;
  center?: string;
  right?: string;
}

export interface Margins {
  top: number;
  bottom: number;
  left: number;
  right: number;
  header: number;
  footer: number;
}

export interface PageSetup {
  orientation?: 'portrait' | 'landscape';
  paper?: PaperSize;
  /** Marges en pouces, comme dans Excel. */
  margins?: Margins;
  /** Échelle en % (10 à 400). */
  scale?: number;
  /** Ajuster à N pages en largeur / hauteur (0 = automatique). */
  fitWidth?: number;
  fitHeight?: number;
  centerH?: boolean;
  centerV?: boolean;
  /** Zone d'impression (« A1:F40 »). */
  printArea?: string;
  /** Lignes à répéter en haut (« 1:2 »). */
  titleRows?: string;
  gridlines?: boolean;
  headings?: boolean;
  header?: HeaderFooterSpec;
  footer?: HeaderFooterSpec;
}

export const PAPER_SIZES: Record<PaperSize, { code: number; widthMm: number; heightMm: number; label: string }> = {
  A4: { code: 9, widthMm: 210, heightMm: 297, label: 'A4 (210 × 297 mm)' },
  A3: { code: 8, widthMm: 297, heightMm: 420, label: 'A3 (297 × 420 mm)' },
  A5: { code: 11, widthMm: 148, heightMm: 210, label: 'A5 (148 × 210 mm)' },
  Letter: { code: 1, widthMm: 215.9, heightMm: 279.4, label: 'Lettre (215,9 × 279,4 mm)' },
  Legal: { code: 5, widthMm: 215.9, heightMm: 355.6, label: 'Légal (215,9 × 355,6 mm)' },
};

export const MARGIN_PRESETS: Record<'normal' | 'narrow' | 'wide', Margins> = {
  normal: { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
  narrow: { left: 0.25, right: 0.25, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
  wide: { left: 1, right: 1, top: 1, bottom: 1, header: 0.5, footer: 0.5 },
};

export const DEFAULT_PAGE: Required<Pick<PageSetup, 'orientation' | 'paper' | 'margins' | 'scale'>> = {
  orientation: 'portrait',
  paper: 'A4',
  margins: MARGIN_PRESETS.normal,
  scale: 100,
};

/** Nom du préréglage de marges (ou « custom »). */
export function marginPresetOf(m: Margins | undefined): 'normal' | 'narrow' | 'wide' | 'custom' {
  const margins = m ?? MARGIN_PRESETS.normal;
  for (const [name, preset] of Object.entries(MARGIN_PRESETS)) {
    if ((Object.keys(preset) as (keyof Margins)[]).every((k) => Math.abs(preset[k] - margins[k]) < 1e-6)) return name as 'normal' | 'narrow' | 'wide';
  }
  return 'custom';
}

// --- En-têtes et pieds de page : « &L gauche &C centre &R droite », codes &P &N &D &T &F &A ---

export function hfToExcel(h: HeaderFooterSpec | undefined): string | undefined {
  if (!h || (!h.left && !h.center && !h.right)) return undefined;
  return `${h.left ? `&L${h.left}` : ''}${h.center ? `&C${h.center}` : ''}${h.right ? `&R${h.right}` : ''}`;
}

export function hfFromExcel(text: string | undefined): HeaderFooterSpec | undefined {
  if (!text) return undefined;
  const out: HeaderFooterSpec = {};
  const parts = text.split(/(&[LCR])/);
  // Texte avant tout marqueur : au centre (comme Excel).
  if (parts[0]) out.center = parts[0];
  for (let i = 1; i < parts.length; i += 2) {
    const value = parts[i + 1] ?? '';
    if (!value) continue;
    const key = parts[i] === '&L' ? 'left' : parts[i] === '&C' ? 'center' : 'right';
    out[key] = value;
  }
  return Object.keys(out).length ? out : undefined;
}

/** Valeur CSS `content` d'un texte d'en-tête (@page margin boxes) : « Page &P / &N » → "Page " counter(page) " / " counter(pages). */
export function hfToCssContent(text: string, ctx: { file: string; sheet: string; date: string; time: string }): string {
  const q = (s: string) => `"${s.replace(/\\/g, '\\\\').replace(/"/g, '\\"').replace(/\n/g, '\\A ')}"`;
  const pieces: string[] = [];
  let literal = '';
  const flush = () => {
    if (literal) pieces.push(q(literal));
    literal = '';
  };
  for (let i = 0; i < text.length; i++) {
    const ch = text[i]!;
    if (ch !== '&') {
      literal += ch;
      continue;
    }
    const code = text[++i];
    if (code === '&') literal += '&';
    else if (code === 'P') (flush(), pieces.push('counter(page)'));
    else if (code === 'N') (flush(), pieces.push('counter(pages)'));
    else if (code === 'D') literal += ctx.date;
    else if (code === 'T') literal += ctx.time;
    else if (code === 'F') literal += ctx.file;
    else if (code === 'A') literal += ctx.sheet;
    // Les autres codes de mise en forme (&B, &I, &"police", &12…) sont ignorés.
    else if (code === '"') {
      const end = text.indexOf('"', i + 1);
      i = end < 0 ? text.length : end;
    } else if (code && /\d/.test(code)) while (/\d/.test(text[i + 1] ?? '')) i++;
  }
  flush();
  return pieces.length ? pieces.join(' ') : '""';
}

// --- Conversion avec ExcelJS ---

export interface ExcelPageSetupLike {
  paperSize?: number;
  orientation?: 'portrait' | 'landscape';
  fitToPage?: boolean;
  fitToWidth?: number;
  fitToHeight?: number;
  scale?: number;
  margins?: Margins;
  printArea?: string;
  printTitlesRow?: string;
  horizontalCentered?: boolean;
  verticalCentered?: boolean;
  showGridLines?: boolean;
  showRowColHeaders?: boolean;
}

export function toExcelPageSetup(p: PageSetup): { pageSetup: ExcelPageSetupLike; oddHeader?: string; oddFooter?: string } {
  const fit = p.fitWidth !== undefined || p.fitHeight !== undefined;
  const pageSetup: ExcelPageSetupLike = {
    paperSize: PAPER_SIZES[p.paper ?? DEFAULT_PAGE.paper].code,
    orientation: p.orientation ?? 'portrait',
    margins: p.margins ?? DEFAULT_PAGE.margins,
  };
  if (fit) {
    pageSetup.fitToPage = true;
    pageSetup.fitToWidth = p.fitWidth ?? 0;
    pageSetup.fitToHeight = p.fitHeight ?? 0;
  } else if (p.scale && p.scale !== 100) {
    pageSetup.scale = Math.round(p.scale);
  }
  if (p.printArea) pageSetup.printArea = p.printArea;
  if (p.titleRows) pageSetup.printTitlesRow = p.titleRows;
  if (p.centerH) pageSetup.horizontalCentered = true;
  if (p.centerV) pageSetup.verticalCentered = true;
  if (p.gridlines) pageSetup.showGridLines = true;
  if (p.headings) pageSetup.showRowColHeaders = true;
  return { pageSetup, oddHeader: hfToExcel(p.header), oddFooter: hfToExcel(p.footer) };
}

export function fromExcelPageSetup(ps: ExcelPageSetupLike | undefined, oddHeader?: string, oddFooter?: string): PageSetup | undefined {
  const p: PageSetup = {};
  if (ps) {
    const paper = (Object.entries(PAPER_SIZES).find(([, v]) => v.code === ps.paperSize)?.[0] ?? undefined) as PaperSize | undefined;
    if (paper && paper !== 'A4') p.paper = paper;
    if (ps.orientation === 'landscape') p.orientation = 'landscape';
    if (ps.margins && marginPresetOf(ps.margins) !== 'normal') p.margins = { ...ps.margins };
    if (ps.fitToPage) {
      p.fitWidth = ps.fitToWidth ?? 1;
      p.fitHeight = ps.fitToHeight ?? 1;
    } else if (ps.scale && ps.scale !== 100) p.scale = ps.scale;
    if (ps.printArea) p.printArea = ps.printArea;
    if (ps.printTitlesRow) p.titleRows = ps.printTitlesRow;
    if (ps.horizontalCentered) p.centerH = true;
    if (ps.verticalCentered) p.centerV = true;
    if (ps.showGridLines) p.gridlines = true;
    if (ps.showRowColHeaders) p.headings = true;
  }
  const header = hfFromExcel(oddHeader);
  const footer = hfFromExcel(oddFooter);
  if (header) p.header = header;
  if (footer) p.footer = footer;
  return Object.keys(p).length ? p : undefined;
}

/** « 1:3 » → lignes 0 à 2 ; null si invalide. */
export function parseRowRange(text: string | undefined): { first: number; last: number } | null {
  const m = /^\$?(\d+):\$?(\d+)$/.exec((text ?? '').trim());
  if (!m) return null;
  const a = Number(m[1]) - 1;
  const b = Number(m[2]) - 1;
  return a >= 0 && b >= a ? { first: a, last: b } : null;
}

/** Zone imprimable de la page, en pixels CSS (96 ppp), marges comprises. */
export function printableSizePx(p: PageSetup): { width: number; height: number } {
  const paper = PAPER_SIZES[p.paper ?? 'A4'];
  const landscape = p.orientation === 'landscape';
  const m = p.margins ?? DEFAULT_PAGE.margins;
  const w = ((landscape ? paper.heightMm : paper.widthMm) / 25.4 - m.left - m.right) * 96;
  const h = ((landscape ? paper.widthMm : paper.heightMm) / 25.4 - m.top - m.bottom) * 96;
  return { width: Math.max(96, w), height: Math.max(96, h) };
}

/** Facteur d'échelle d'impression : échelle fixe ou ajustement à N pages de large. */
export function printZoom(p: PageSetup, contentWidthPx: number): number {
  if (p.fitWidth) {
    const target = printableSizePx(p).width * p.fitWidth;
    return contentWidthPx > target ? target / contentWidthPx : 1;
  }
  return (p.scale ?? 100) / 100;
}
