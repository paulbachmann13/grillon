/**
 * Formes et zones de texte ⇄ DrawingML (xdr:sp / xdr:cxnSp dans xl/drawings/drawingN.xml).
 */
import type { ShapeKind, ShapeSpec } from '../../charts/shape';
import { resolveShape, SHAPE_FILL, SHAPE_STROKE, TEXTBOX_STROKE } from '../../charts/shape';
import { escapeXml, parseXml } from './ooxml';

type Xml = Record<string, unknown>;
const arr = <T>(v: T | T[] | undefined): T[] => (v === undefined ? [] : Array.isArray(v) ? v : [v]);
const hex = (c: string) => c.replace('#', '').toUpperCase();
const fillXml = (c: string) => (c === 'none' ? '<a:noFill/>' : `<a:solidFill><a:srgbClr val="${hex(c)}"/></a:solidFill>`);

const PRST: Record<ShapeKind, string> = { rect: 'rect', roundRect: 'roundRect', ellipse: 'ellipse', line: 'line', arrow: 'line', textbox: 'rect' };

/** Contenu d'un <xdr:sp> ou <xdr:cxnSp> (sans l'ancre). */
export function shapeXml(spec: ShapeSpec, id: number, name: string): string {
  const r = resolveShape(spec);
  const linear = spec.kind === 'line' || spec.kind === 'arrow';
  const widthEmu = Math.round(r.strokeWidth * 12700);
  const ln = r.stroke === 'none' ? '<a:ln><a:noFill/></a:ln>' : `<a:ln w="${widthEmu}">${fillXml(r.stroke)}${spec.kind === 'arrow' ? '<a:tailEnd type="triangle"/>' : ''}</a:ln>`;
  const xfrm = '<a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm>';
  if (linear) {
    return `<xdr:cxnSp macro=""><xdr:nvCxnSpPr><xdr:cNvPr id="${id}" name="${escapeXml(name)}"/><xdr:cNvCxnSpPr/></xdr:nvCxnSpPr><xdr:spPr>${xfrm}<a:prstGeom prst="line"><a:avLst/></a:prstGeom>${ln}</xdr:spPr></xdr:cxnSp>`;
  }
  const rpr = `<a:rPr lang="fr-FR" sz="${Math.round(r.fontSize * 100)}"${spec.bold ? ' b="1"' : ''}${spec.italic ? ' i="1"' : ''}>${fillXml(r.color)}</a:rPr>`;
  const algn = r.align === 'center' ? 'ctr' : r.align === 'right' ? 'r' : 'l';
  const paragraphs = (spec.text ?? '').split('\n').map((line) => `<a:p><a:pPr algn="${algn}"/>${line ? `<a:r>${rpr}<a:t>${escapeXml(line)}</a:t></a:r>` : ''}<a:endParaRPr lang="fr-FR" sz="${Math.round(r.fontSize * 100)}"/></a:p>`).join('');
  const body = `<xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip" wrap="square" rtlCol="0" anchor="${spec.kind === 'textbox' ? 't' : 'ctr'}"/><a:lstStyle/>${paragraphs || '<a:p><a:endParaRPr lang="fr-FR"/></a:p>'}</xdr:txBody>`;
  return `<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="${id}" name="${escapeXml(name)}"/><xdr:cNvSpPr${spec.kind === 'textbox' ? ' txBox="1"' : ''}/></xdr:nvSpPr><xdr:spPr>${xfrm}<a:prstGeom prst="${PRST[spec.kind]}"><a:avLst/></a:prstGeom>${fillXml(r.fill)}${ln}</xdr:spPr>${body}</xdr:sp>`;
}

const colorOf = (node: unknown): string | undefined => {
  const v = ((node as { srgbClr?: { val?: string } } | undefined)?.srgbClr)?.val;
  return v ? `#${v.toUpperCase()}` : undefined;
};

/** Lit une forme dans une ancre (null si l'ancre ne contient ni forme ni connecteur). */
export function parseShapeAnchor(anchorXml: string): ShapeSpec | null {
  const doc = parseXml<{ twoCellAnchor?: Xml; oneCellAnchor?: Xml }>(anchorXml);
  const anchor = (doc.twoCellAnchor ?? doc.oneCellAnchor) as Xml | undefined;
  const sp = anchor?.sp as Xml | undefined;
  const cxn = anchor?.cxnSp as Xml | undefined;
  const node = sp ?? cxn;
  if (!node) return null;
  const spPr = node.spPr as Xml | undefined;
  const prst = String((spPr?.prstGeom as { prst?: string } | undefined)?.prst ?? 'rect');
  const txBox = (((node.nvSpPr as Xml | undefined)?.cNvSpPr) as { txBox?: unknown } | undefined)?.txBox;
  const ln = spPr?.ln as Xml | undefined;
  let kind: ShapeKind;
  if (cxn || prst === 'line' || prst === 'straightConnector1') kind = ln?.tailEnd !== undefined && (ln.tailEnd as { type?: string }).type !== 'none' ? 'arrow' : 'line';
  else if (prst === 'ellipse') kind = 'ellipse';
  else if (prst === 'roundRect') kind = 'roundRect';
  else if (prst === 'rect' || prst === 'flowChartProcess') kind = txBox === '1' || txBox === 1 || txBox === true ? 'textbox' : 'rect';
  else return null; // géométrie non prise en charge
  const spec: ShapeSpec = { kind };
  const defaults = resolveShape({ kind });
  const fill = spPr?.noFill !== undefined ? 'none' : colorOf(spPr?.solidFill);
  if (fill && fill !== defaults.fill) spec.fill = fill;
  const stroke = ln?.noFill !== undefined ? 'none' : colorOf(ln?.solidFill);
  if (stroke && stroke !== defaults.stroke) spec.stroke = stroke;
  const w = Number((ln as { w?: unknown } | undefined)?.w);
  if (Number.isFinite(w) && w > 0) {
    const pt = Math.round((w / 12700) * 100) / 100;
    if (pt !== defaults.strokeWidth) spec.strokeWidth = pt;
  }
  if (kind !== 'line' && kind !== 'arrow') {
    const paras = arr(((node.txBody as Xml | undefined)?.p) as Xml | Xml[] | undefined);
    const lines: string[] = [];
    for (const p of paras) {
      const runs = arr(p.r as Xml | Xml[] | undefined);
      lines.push(runs.map((r) => String(r.t ?? '')).join(''));
      const rPr = runs[0]?.rPr as Xml | undefined;
      if (rPr && spec.fontSize === undefined) {
        const sz = Number(rPr.sz);
        if (Number.isFinite(sz) && sz > 0 && sz / 100 !== defaults.fontSize) spec.fontSize = sz / 100;
        const c = colorOf(rPr.solidFill);
        if (c && c !== defaults.color) spec.color = c;
        if (String(rPr.b) === '1') spec.bold = true;
        if (String(rPr.i) === '1') spec.italic = true;
      }
      const algn = String((p.pPr as { algn?: string } | undefined)?.algn ?? '');
      const align = algn === 'ctr' ? 'center' : algn === 'r' ? 'right' : algn === 'l' ? 'left' : undefined;
      if (align && align !== defaults.align && !spec.align) spec.align = align;
    }
    const text = lines.join('\n').replace(/\n+$/, '');
    if (text) spec.text = text;
  }
  void SHAPE_FILL;
  void SHAPE_STROKE;
  void TEXTBOX_STROKE;
  return spec;
}
