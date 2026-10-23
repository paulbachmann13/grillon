/**
 * Accès direct au paquet .xlsx (zip + XML OOXML), pour ce qu'ExcelJS ne gère pas ou mal :
 * critères de filtre, tableaux structurés, « interrompre si vrai » des MFC, validations
 * sur de grandes plages (ExcelJS les déplie cellule par cellule).
 *
 * On ne parse jamais une feuille entière (elle peut faire des dizaines de Mo) : on extrait
 * les éléments utiles par recherche textuelle, puis on parse ces seuls fragments.
 */
import { XMLParser } from 'fast-xml-parser';
import JSZip from 'jszip';

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '',
  removeNSPrefix: true,
  parseAttributeValue: false,
  parseTagValue: false,
  trimValues: false,
  isArray: (name) =>
    ['dataValidation', 'filterColumn', 'filter', 'customFilter', 'tableColumn', 'Relationship', 'sheet', 'dateGroupItem'].includes(name),
});

/** Texte d'un élément lu par parseXml : une chaîne, ou { '#text' } quand l'élément a des attributs (xml:space="preserve"). */
export function textOf(node: unknown): string | undefined {
  if (typeof node === 'string') return node;
  if (typeof node === 'number' || typeof node === 'boolean') return String(node);
  const t = (node as { '#text'?: unknown } | null | undefined)?.['#text'];
  return t === undefined ? undefined : String(t);
}

export function parseXml<T = Record<string, unknown>>(xml: string): T {
  return parser.parse(xml) as T;
}

export function escapeXml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

export function attrs(obj: Record<string, string | number | boolean | undefined | null>): string {
  return Object.entries(obj)
    .filter(([, v]) => v !== undefined && v !== null && v !== false)
    .map(([k, v]) => ` ${k}="${escapeXml(v === true ? '1' : String(v))}"`)
    .join('');
}

/** Extrait le premier élément `<tag …>…</tag>` ou `<tag …/>` (préfixe d'espace de noms facultatif). */
export function extractElement(xml: string, tag: string, from = 0): { start: number; end: number; xml: string } | null {
  const open = new RegExp(`<(?:\\w+:)?${tag}[\\s/>]`, 'g');
  open.lastIndex = from;
  const m = open.exec(xml);
  if (!m) return null;
  const start = m.index;
  const prefix = /<(\w+:)?/.exec(xml.slice(start))![1] ?? '';
  const gt = xml.indexOf('>', start);
  if (gt < 0) return null;
  if (xml[gt - 1] === '/') return { start, end: gt + 1, xml: xml.slice(start, gt + 1) };
  const close = `</${prefix}${tag}>`;
  const endIdx = xml.indexOf(close, gt);
  if (endIdx < 0) return null;
  const end = endIdx + close.length;
  return { start, end, xml: xml.slice(start, end) };
}

export function extractAll(xml: string, tag: string): { start: number; end: number; xml: string }[] {
  const out = [];
  let from = 0;
  for (;;) {
    const el = extractElement(xml, tag, from);
    if (!el) break;
    out.push(el);
    from = el.end;
  }
  return out;
}

// --- Navigation dans le paquet ---

export interface SheetEntry {
  name: string;
  path: string;
}

function resolveTarget(baseDir: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const parts = `${baseDir}/${target}`.split('/');
  const out: string[] = [];
  for (const p of parts) {
    if (p === '..') out.pop();
    else if (p && p !== '.') out.push(p);
  }
  return out.join('/');
}

export function relsPathFor(partPath: string): string {
  const i = partPath.lastIndexOf('/');
  return `${partPath.slice(0, i)}/_rels/${partPath.slice(i + 1)}.rels`;
}

export interface Relationship {
  Id: string;
  Type: string;
  Target: string;
  TargetMode?: string;
}

export async function readRelationships(zip: JSZip, partPath: string): Promise<Relationship[]> {
  const file = zip.file(relsPathFor(partPath));
  if (!file) return [];
  const doc = parseXml<{ Relationships?: { Relationship?: Relationship[] } }>(await file.async('string'));
  return doc.Relationships?.Relationship ?? [];
}

export function resolveRelTarget(partPath: string, target: string): string {
  return resolveTarget(partPath.slice(0, partPath.lastIndexOf('/')), target);
}

/** Feuilles du classeur, dans l'ordre, avec le chemin de leur XML. */
export async function listSheets(zip: JSZip): Promise<SheetEntry[]> {
  const wbFile = zip.file('xl/workbook.xml');
  if (!wbFile) return [];
  const wbXml = await wbFile.async('string');
  const sheetsEl = extractElement(wbXml, 'sheets');
  if (!sheetsEl) return [];
  const doc = parseXml<{ sheets?: { sheet?: Record<string, string>[] } }>(sheetsEl.xml);
  const rels = await readRelationships(zip, 'xl/workbook.xml');
  const byId = new Map(rels.map((r) => [r.Id, r]));
  return (doc.sheets?.sheet ?? []).flatMap((s) => {
    const rid = s.id ?? Object.entries(s).find(([k]) => k.endsWith(':id') || k === 'id')?.[1];
    const rel = rid ? byId.get(rid) : undefined;
    return rel ? [{ name: s.name, path: resolveRelTarget('xl/workbook.xml', rel.Target) }] : [];
  });
}

export async function loadZip(bytes: ArrayBuffer | Uint8Array): Promise<JSZip> {
  return JSZip.loadAsync(bytes);
}

export async function zipToBytes(zip: JSZip): Promise<Uint8Array> {
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 } });
}

/** Ajoute une relation à une partie et renvoie son identifiant. */
export async function addRelationship(zip: JSZip, partPath: string, type: string, target: string): Promise<string> {
  const path = relsPathFor(partPath);
  const file = zip.file(path);
  let xml = file
    ? await file.async('string')
    : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
  const ids = [...xml.matchAll(/Id="rId(\d+)"/g)].map((m) => Number(m[1]));
  const id = `rId${(ids.length ? Math.max(...ids) : 0) + 1}`;
  xml = xml.replace('</Relationships>', `<Relationship${attrs({ Id: id, Type: type, Target: target })}/></Relationships>`);
  zip.file(path, xml);
  return id;
}

export async function addContentTypeOverride(zip: JSZip, partName: string, contentType: string): Promise<void> {
  const file = zip.file('[Content_Types].xml');
  if (!file) return;
  let xml = await file.async('string');
  if (xml.includes(`PartName="${partName}"`)) return;
  xml = xml.replace('</Types>', `<Override${attrs({ PartName: partName, ContentType: contentType })}/></Types>`);
  zip.file('[Content_Types].xml', xml);
}

/**
 * Ordre des éléments enfants de <worksheet> (schéma OOXML). Sert à insérer un élément
 * au bon endroit : Excel refuse (« réparation ») un fichier où l'ordre n'est pas respecté.
 */
const WORKSHEET_ORDER = [
  'sheetPr', 'dimension', 'sheetViews', 'sheetFormatPr', 'cols', 'sheetData', 'sheetCalcPr', 'sheetProtection',
  'protectedRanges', 'scenarios', 'autoFilter', 'sortState', 'dataConsolidate', 'customSheetViews', 'mergeCells',
  'phoneticPr', 'conditionalFormatting', 'dataValidations', 'hyperlinks', 'printOptions', 'pageMargins', 'pageSetup',
  'headerFooter', 'rowBreaks', 'colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags',
  'drawing', 'legacyDrawing', 'legacyDrawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems',
  'tableParts', 'extLst',
];

export interface TopLevelElement {
  tag: string;
  start: number;
  end: number;
}

/**
 * Enfants directs de l'élément racine (<worksheet>). Un même nom peut apparaître plus bas
 * dans l'arbre (ex. <extLst> dans une règle de MFC) : seuls les enfants directs comptent.
 */
export function topLevelChildren(xml: string): TopLevelElement[] {
  const out: TopLevelElement[] = [];
  // Saute le prologue et la balise ouvrante de la racine.
  let pos = 0;
  for (;;) {
    pos = xml.indexOf('<', pos);
    if (pos < 0) return out;
    if (xml.startsWith('<?', pos) || xml.startsWith('<!', pos)) {
      pos = xml.indexOf('>', pos) + 1;
      continue;
    }
    pos = xml.indexOf('>', pos) + 1;
    break;
  }
  for (;;) {
    pos = xml.indexOf('<', pos);
    if (pos < 0 || xml[pos + 1] === '/') return out;
    const m = /^<((?:\w+:)?([\w.]+))/.exec(xml.slice(pos, pos + 80));
    if (!m) return out;
    const qname = m[1];
    const gt = xml.indexOf('>', pos);
    let end: number;
    if (xml[gt - 1] === '/') end = gt + 1;
    else {
      // Les éléments de même nom ne s'imbriquent pas directement dans une feuille.
      const close = `</${qname}>`;
      const c = xml.indexOf(close, gt);
      if (c < 0) return out;
      end = c + close.length;
    }
    out.push({ tag: m[2], start: pos, end });
    pos = end;
  }
}

/** Remplace (ou insère à sa place selon le schéma) l'élément de premier niveau `tag`. */
export function insertWorksheetElement(sheetXml: string, tag: string, fragment: string): string {
  const children = topLevelChildren(sheetXml);
  const existing = children.find((c) => c.tag === tag);
  if (existing) return sheetXml.slice(0, existing.start) + fragment + sheetXml.slice(existing.end);
  const rank = WORKSHEET_ORDER.indexOf(tag);
  const next = children.find((c) => {
    const r = WORKSHEET_ORDER.indexOf(c.tag);
    return r > rank;
  });
  const pos = next ? next.start : sheetXml.lastIndexOf('</');
  return sheetXml.slice(0, pos) + fragment + sheetXml.slice(pos);
}

/** Retire un élément de premier niveau (s'il existe). */
export function removeWorksheetElement(sheetXml: string, tag: string): { xml: string; removed?: string } {
  const el = topLevelChildren(sheetXml).find((c) => c.tag === tag);
  if (!el) return { xml: sheetXml };
  return { xml: sheetXml.slice(0, el.start) + sheetXml.slice(el.end), removed: sheetXml.slice(el.start, el.end) };
}

export function getWorksheetElement(sheetXml: string, tag: string): string | undefined {
  const el = topLevelChildren(sheetXml).find((c) => c.tag === tag);
  return el ? sheetXml.slice(el.start, el.end) : undefined;
}
