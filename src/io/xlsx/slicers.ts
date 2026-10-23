/**
 * Segments et chronologies ⇄ .xlsx.
 *
 * Excel : un segment est fait d'un cache (xl/slicerCaches/…, relié au classeur), d'une partie « segments » par feuille
 * (xl/slicers/…, listée dans l'extension de la feuille) et d'un cadre sle:slicer dans le dessin. Un segment de tableau
 * exige un vrai tableau Excel (on en crée un, sans style, si la plage filtrée n'en est pas un) ; un segment de tableau
 * croisé, le tableau croisé écrit par Grillon. Les choix du segment sont le filtre du tableau (autoFilter) ou les
 * éléments masqués du tableau croisé.
 *
 * Ce qu'Excel ne sait pas représenter (chronologie sur un tableau, segment sur une plage déjà filtrée…) et l'état exact
 * de Grillon (identifiants, positions, tableaux créés pour l'occasion) sont gardés dans une partie XML personnalisée
 * (customXml), qu'Excel conserve à l'enregistrement. Sans elle (fichier venu d'Excel), on relit les segments d'Excel.
 */
import type JSZip from 'jszip';
import { EMPTY_LABEL, type SlicerSpec } from '../../data/slicer';
import type { IWorkbookData } from '../../model/univer-types';
import { addContentTypeOverride, addRelationship, attrs, escapeXml, extractAll, extractElement, readRelationships, resolveRelTarget } from './ooxml';
import type { UniverTable, UniverTableResource } from './tables';
import { rangeFromA1, rangeToA1 } from './units';

export const SLICER_COMPONENT_KEY = 'GrillonSlicer';

const NS_X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
const NS_X15 = 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main';
const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const NS_MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const REL_SLICER = 'http://schemas.microsoft.com/office/2007/relationships/slicer';
const REL_SLICER_CACHE = 'http://schemas.microsoft.com/office/2007/relationships/slicerCache';
const REL_CUSTOM_XML = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml';
const REL_CUSTOM_XML_PROPS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps';
const CT_SLICER = 'application/vnd.ms-excel.slicer+xml';
const CT_SLICER_CACHE = 'application/vnd.ms-excel.slicerCache+xml';
const CT_CUSTOM_XML_PROPS = 'application/vnd.openxmlformats-officedocument.customXmlProperties+xml';
const GRILLON_NS = 'urn:grillon:segments:1';
/** Extensions (uri) d'Excel : listes de segments de feuille, caches du classeur, cache de tableau croisé. */
const EXT = {
  sheetPivotSlicers: '{A8765BA9-456A-4dab-B4F3-ACF838C121DE}',
  sheetTableSlicers: '{3A4CF648-6AED-40f4-86FF-DC5316D8AED3}',
  bookPivotCaches: '{BBE1A952-AA13-448e-AADC-164F8A28A991}',
  bookTableCaches: '{46BE6895-7355-4a93-B00E-2C351335B9C9}',
  tableSlicerCache: '{2F2917AC-EB37-4324-AD4E-5DD8C200BD13}',
  pivotCacheId: '{725AE2AE-9491-48be-B2B4-4EB974FC3084}',
};
const EMU_PER_PX = 9525;

export interface CellAnchor {
  column: number;
  columnOffset: number;
  row: number;
  rowOffset: number;
}

/** Entrée de la ressource de dessins d'Univer pour un segment (composant flottant). */
export interface SlicerDrawingItem {
  drawingId: string;
  componentKey?: string;
  sheetTransform: { from: CellAnchor; to: CellAnchor };
  data?: SlicerSpec;
  [key: string]: unknown;
}

type DrawingResource = Record<string, { data: Record<string, SlicerDrawingItem>; order: string[] }>;

/** Segment écrit comme un vrai segment d'Excel. */
export interface NativeSlicer {
  /** Nom du segment (et du cadre dans le dessin), unique dans le classeur. */
  name: string;
  /** Nom du cache (nom défini « Segment_… »). */
  cacheName: string;
  caption: string;
  spec: SlicerSpec;
  /** Feuille où le segment est dessiné. */
  drawSheetId: string;
  from: CellAnchor;
  to: CellAnchor;
  /** Segment de tableau : feuille et coin supérieur gauche du tableau (identifie le tableau écrit). */
  table?: { sheetId: string; startRow: number; startColumn: number; column: number };
  /** Segment de tableau croisé : feuille du tableau croisé et colonne de la source. */
  pivot?: { sheetId: string; col: number };
}

/** Ce que Grillon garde pour lui (partie customXml). */
interface SavedState {
  version: 1;
  /** Segments et chronologies : feuille de dessin (nom), entrée de dessin, nom Excel s'il a été écrit comme segment. */
  slicers: { sheet: string; item: SlicerDrawingItem; native?: string }[];
  /** Tableaux créés seulement pour porter un segment (feuille, plage) : retirés à la réouverture dans Grillon. */
  syntheticTables: { sheet: string; ref: string }[];
}

export interface SlicerPlan {
  native: NativeSlicer[];
  /** Tableaux à créer (sans style) pour les segments de plages ordinaires, par feuille. */
  extraTables: Record<string, UniverTable[]>;
  /** Choix des segments de tableau : feuille|ligne|colonne du tableau → colonne → libellés retenus. */
  tableFilters: Map<string, Map<number, string[]>>;
  /** Colonnes de la source à décrire pour les segments de chaque tableau croisé (par feuille). */
  pivotSlicerCols: Record<string, number[]>;
  saved: SavedState;
}

export const tableKey = (sheetId: string, startRow: number, startColumn: number) => `${sheetId}|${startRow}|${startColumn}`;

/** Nom valide pour un nom défini d'Excel (« Segment_Région »). */
function definedName(base: string, used: Set<string>): string {
  let name = base.replace(/[^\p{L}\p{N}_.]/gu, '_');
  if (!/^[\p{L}_]/u.test(name)) name = `_${name}`;
  let out = name;
  for (let i = 1; used.has(out.toLowerCase()); i++) out = `${name}${i}`;
  used.add(out.toLowerCase());
  return out;
}

function uniqueName(base: string, used: Set<string>): string {
  const name = base.trim() || 'Segment';
  let out = name;
  for (let i = 1; used.has(out.toLowerCase()); i++) out = `${name} ${i}`;
  used.add(out.toLowerCase());
  return out;
}

const sameRange = (a: ReturnType<typeof rangeFromA1>, b: ReturnType<typeof rangeFromA1>) =>
  a.startRow === b.startRow && a.endRow === b.endRow && a.startColumn === b.startColumn && a.endColumn === b.endColumn;
const overlaps = (a: ReturnType<typeof rangeFromA1>, b: ReturnType<typeof rangeFromA1>) =>
  a.startRow <= b.endRow && b.startRow <= a.endRow && a.startColumn <= b.endColumn && b.startColumn <= a.endColumn;

function cellText(wb: IWorkbookData, sheetId: string, row: number, col: number): string {
  const v = wb.sheets[sheetId]?.cellData?.[row]?.[col]?.v;
  return v === undefined || v === null ? '' : String(v);
}

/**
 * Prépare l'écriture des segments : lesquels deviennent de vrais segments d'Excel, les tableaux à créer,
 * les filtres à poser et l'état à garder pour Grillon.
 */
export function planSlicers(
  wb: IWorkbookData,
  drawings: DrawingResource | undefined,
  tables: UniverTableResource,
  autoFilters: Record<string, { ref?: ReturnType<typeof rangeFromA1> } | undefined>,
  hasPivot: (sheetId: string) => boolean,
): SlicerPlan {
  const plan: SlicerPlan = { native: [], extraTables: {}, tableFilters: new Map(), pivotSlicerCols: {}, saved: { version: 1, slicers: [], syntheticTables: [] } };
  const idByName = new Map(wb.sheetOrder.map((id) => [wb.sheets[id]?.name ?? '', id]));
  const names = new Set<string>();
  const cacheNames = new Set<string>();
  const tableNames = new Set(Object.values(tables).flatMap((t) => t.tables.map((x) => x.name.toLowerCase())));
  for (const sheetId of wb.sheetOrder) {
    const res = drawings?.[sheetId];
    if (!res) continue;
    for (const id of res.order ?? Object.keys(res.data ?? {})) {
      const item = res.data?.[id];
      const spec = item?.data;
      if (!item || item.componentKey !== SLICER_COMPONENT_KEY || !spec?.target || !item.sheetTransform) continue;
      const saved: SavedState['slicers'][number] = { sheet: wb.sheets[sheetId]?.name ?? '', item };
      plan.saved.slicers.push(saved);
      if (spec.type !== 'slicer') continue; // chronologie : gardée pour Grillon seulement
      const targetId = idByName.get(spec.target.sheet);
      if (!targetId) continue;
      const base = { caption: spec.title, spec, drawSheetId: sheetId, from: item.sheetTransform.from, to: item.sheetTransform.to };
      if (spec.target.kind === 'table') {
        let r: ReturnType<typeof rangeFromA1>;
        try {
          r = rangeFromA1(spec.target.ref.replace(/\$/g, ''));
        } catch {
          continue;
        }
        if (r.endRow <= r.startRow) continue;
        const existing = [...(tables[targetId]?.tables ?? []), ...(plan.extraTables[targetId] ?? [])].find((t) => sameRange(t.range, r));
        if (!existing) {
          // Un tableau ne peut pas chevaucher un filtre automatique ni un autre tableau.
          const filterRef = autoFilters[targetId]?.ref;
          if (filterRef && overlaps(filterRef, r)) continue;
          if ((tables[targetId]?.tables ?? []).some((t) => overlaps(t.range, r))) continue;
          let name = 'Tableau1';
          for (let i = 1; tableNames.has(name.toLowerCase()); i++) name = `Tableau${i}`;
          tableNames.add(name.toLowerCase());
          const columns = [];
          for (let c = r.startColumn; c <= r.endColumn; c++) {
            columns.push({ id: `col${c}`, displayName: cellText(wb, targetId, r.startRow, c) || `Colonne${c - r.startColumn + 1}`, dataType: 'string', formula: '', showFilterButton: true, meta: {}, style: {} });
          }
          const table: UniverTable = {
            id: `slicer-table-${targetId}-${r.startRow}-${r.startColumn}`,
            name,
            range: r,
            options: { showHeader: true, showFooter: false, showAutoFilter: true },
            filters: { tableColumnFilterList: [] },
            columns,
            meta: { excel: { styleName: '', showRowStripes: false } },
          };
          (plan.extraTables[targetId] ??= []).push(table);
          plan.saved.syntheticTables.push({ sheet: spec.target.sheet, ref: rangeToA1(r) });
        }
        const name = uniqueName(spec.title, names);
        const native: NativeSlicer = { ...base, name, cacheName: definedName(`Segment_${spec.title}`, cacheNames), table: { sheetId: targetId, startRow: r.startRow, startColumn: r.startColumn, column: spec.target.col } };
        plan.native.push(native);
        saved.native = name;
        if (spec.selected) {
          const key = tableKey(targetId, r.startRow, r.startColumn);
          const byCol = plan.tableFilters.get(key) ?? new Map<number, string[]>();
          const before = byCol.get(spec.target.col);
          byCol.set(spec.target.col, before ? before.filter((l) => spec.selected!.includes(l)) : [...spec.selected]);
          plan.tableFilters.set(key, byCol);
        }
      } else {
        if (!hasPivot(targetId)) continue;
        const name = uniqueName(spec.title, names);
        plan.native.push({ ...base, name, cacheName: definedName(`Segment_${spec.title}`, cacheNames), pivot: { sheetId: targetId, col: spec.target.col } });
        saved.native = name;
        const cols = (plan.pivotSlicerCols[targetId] ??= []);
        if (!cols.includes(spec.target.col)) cols.push(spec.target.col);
      }
    }
  }
  return plan;
}

/** Critères d'un filtre de tableau (dans <autoFilter>) pour les choix des segments. */
export function tableFilterXml(byCol: Map<number, string[]> | undefined): string {
  if (!byCol?.size) return '';
  return [...byCol]
    .sort(([a], [b]) => a - b)
    .map(([col, labels]) => {
      const blank = labels.includes(EMPTY_LABEL);
      const vals = labels.filter((l) => l !== EMPTY_LABEL).map((l) => `<filter${attrs({ val: l })}/>`).join('');
      return `<filterColumn colId="${col}"><filters${blank ? ' blank="1"' : ''}>${vals}</filters></filterColumn>`;
    })
    .join('');
}

/** Cadre du segment dans le dessin (avec un texte de remplacement pour les logiciels qui ne les connaissent pas). */
export function slicerFrameXml(name: string, id: number, kind: 'table' | 'pivot'): string {
  const [prefix, ns] = kind === 'table' ? ['sle15', 'http://schemas.microsoft.com/office/drawing/2012/slicer'] : ['a14', 'http://schemas.microsoft.com/office/drawing/2010/main'];
  return (
    `<mc:AlternateContent xmlns:mc="${NS_MC}"><mc:Choice xmlns:${prefix}="${ns}" Requires="${prefix}">` +
    `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr${attrs({ id, name })}/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
    '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>' +
    `<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2010/slicer"><sle:slicer xmlns:sle="http://schemas.microsoft.com/office/drawing/2010/slicer"${attrs({ name })}/></a:graphicData></a:graphic>` +
    '</xdr:graphicFrame></mc:Choice><mc:Fallback>' +
    `<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr${attrs({ id, name })}/><xdr:cNvSpPr><a:spLocks noTextEdit="1"/></xdr:cNvSpPr></xdr:nvSpPr>` +
    '<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="2524125"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:prstClr val="white"/></a:solidFill><a:ln w="1"><a:solidFill><a:prstClr val="green"/></a:solidFill></a:ln></xdr:spPr>' +
    '<xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip"/><a:lstStyle/><a:p><a:r><a:rPr lang="fr-FR" sz="1100"/><a:t>Ce segment n’est pas pris en charge par votre version d’Excel.</a:t></a:r></a:p></xdr:txBody></xdr:sp>' +
    '</mc:Fallback></mc:AlternateContent>'
  );
}

/** Informations sur un tableau croisé écrit (pour relier ses segments). */
export interface WrittenPivot {
  name: string;
  /** Identifiant x14 du cache (pivotCacheId). */
  pivotCacheId: number;
  /** Éléments partagés de chaque champ (dans l'ordre du cache), libellés tels que Grillon les affiche. */
  items: (string[] | undefined)[];
  fieldNames: string[];
  /** Libellés masqués (filtres) par champ. */
  hidden: Map<number, Set<string>>;
}

interface SheetEntryWithId {
  name: string;
  path: string;
  sheetId: number;
}

async function sheetsWithIds(zip: JSZip): Promise<SheetEntryWithId[]> {
  const wbXml = (await zip.file('xl/workbook.xml')?.async('string')) ?? '';
  const rels = new Map((await readRelationships(zip, 'xl/workbook.xml')).map((r) => [r.Id, r]));
  const out: SheetEntryWithId[] = [];
  for (const m of wbXml.matchAll(/<(?:\w+:)?sheet\b([^>]*)\/?>/g)) {
    const a = m[1];
    const get = (k: string) => new RegExp(`(?:^|\\s)${k}="([^"]*)"`).exec(a)?.[1];
    const rel = rels.get(get('r:id') ?? '');
    if (rel) out.push({ name: unescapeXml(get('name') ?? ''), path: resolveRelTarget('xl/workbook.xml', rel.Target), sheetId: Number(get('sheetId')) });
  }
  return out;
}

function unescapeXml(s: string): string {
  return s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
}

/** Ajoute une extension à l'<extLst> d'une feuille ou du classeur (créée au besoin, toujours en dernier). */
function appendExt(xml: string, rootTag: string, ext: string): string {
  const close = `</${rootTag}>`;
  const at = xml.lastIndexOf(close);
  const existing = extractAll(xml, 'extLst').filter((e) => xml.indexOf(close, e.end) === at && !xml.slice(e.end, at).includes('<'));
  const top = existing[existing.length - 1];
  if (top) return xml.slice(0, top.end - '</extLst>'.length) + ext + xml.slice(top.end - '</extLst>'.length);
  return xml.slice(0, at) + `<extLst>${ext}</extLst>` + xml.slice(at);
}

/** Ajoute des noms définis (#N/A) pour les caches de segments, comme Excel. */
function addDefinedNames(wbXml: string, names: string[]): string {
  const fragment = names.map((n) => `<definedName${attrs({ name: n })}>#N/A</definedName>`).join('');
  const el = extractElement(wbXml, 'definedNames');
  if (el && !el.xml.endsWith('/>')) return wbXml.slice(0, el.end - '</definedNames>'.length) + fragment + wbXml.slice(el.end - '</definedNames>'.length);
  if (el) return wbXml.slice(0, el.start) + `<definedNames>${fragment}</definedNames>` + wbXml.slice(el.end);
  const m = /<calcPr\b|<oleSize\b|<customWorkbookViews\b|<pivotCaches\b|<extLst\b|<\/workbook>/.exec(wbXml);
  return m ? wbXml.slice(0, m.index) + `<definedNames>${fragment}</definedNames>` + wbXml.slice(m.index) : wbXml;
}

/**
 * Écrit les segments d'Excel (caches, parties de feuille, extensions) et l'état de Grillon.
 * À appeler après les feuilles, les tableaux, les dessins et les tableaux croisés.
 */
export async function writeSlicers(zip: JSZip, plan: SlicerPlan, sheetPathOf: (sheetId: string) => string | undefined, pivots: Map<string, WrittenPivot>): Promise<void> {
  if (plan.native.length) await writeNativeSlicers(zip, plan, sheetPathOf, pivots);
  if (plan.saved.slicers.length) await writeSavedState(zip, plan.saved);
}

async function writeNativeSlicers(zip: JSZip, plan: SlicerPlan, sheetPathOf: (sheetId: string) => string | undefined, pivots: Map<string, WrittenPivot>): Promise<void> {
  const sheets = await sheetsWithIds(zip);
  // Tableaux écrits : (feuille, coin) → id et noms de colonnes.
  const tablesByCorner = new Map<string, { id: number; columns: string[] }>();
  for (const sheet of sheets) {
    for (const rel of (await readRelationships(zip, sheet.path)).filter((r) => r.Type.endsWith('/table'))) {
      const xml = await zip.file(resolveRelTarget(sheet.path, rel.Target))?.async('string');
      if (!xml) continue;
      const id = Number(/<table\b[^>]*\sid="(\d+)"/.exec(xml)?.[1]);
      const ref = /<table\b[^>]*\sref="([^"]+)"/.exec(xml)?.[1];
      if (!id || !ref) continue;
      const r = rangeFromA1(ref);
      const columns = [...xml.matchAll(/<tableColumn\b[^>]*\sname="([^"]*)"/g)].map((m) => unescapeXml(m[1]));
      tablesByCorner.set(`${sheet.path}|${r.startRow}|${r.startColumn}`, { id, columns });
    }
  }
  const tableCaches: string[] = [];
  const pivotCaches: string[] = [];
  const written: { slicer: NativeSlicer; kind: 'table' | 'pivot' }[] = [];
  let n = 0;
  for (const s of plan.native) {
    let body: string;
    let sourceName: string;
    let kind: 'table' | 'pivot';
    if (s.table) {
      const path = sheetPathOf(s.table.sheetId);
      const table = path ? tablesByCorner.get(`${path}|${s.table.startRow}|${s.table.startColumn}`) : undefined;
      if (!table || !table.columns[s.table.column]) continue;
      sourceName = table.columns[s.table.column];
      kind = 'table';
      body = `<extLst><x:ext uri="${EXT.tableSlicerCache}" xmlns:x15="${NS_X15}"><x15:tableSlicerCache tableId="${table.id}" column="${s.table.column + 1}"/></x:ext></extLst>`;
    } else if (s.pivot) {
      const path = sheetPathOf(s.pivot.sheetId);
      const pivot = path ? pivots.get(path) : undefined;
      const sheet = sheets.find((x) => x.path === path);
      const items = pivot?.items[s.pivot.col];
      if (!pivot || !sheet || !items) continue;
      sourceName = pivot.fieldNames[s.pivot.col];
      kind = 'pivot';
      const selected = s.spec.selected;
      const hidden = pivot.hidden.get(s.pivot.col);
      const isOn = (label: string) => (selected ? selected.includes(label) : !hidden?.has(label));
      body =
        `<pivotTables><pivotTable${attrs({ tabId: sheet.sheetId, name: pivot.name })}/></pivotTables>` +
        `<data><tabular pivotCacheId="${pivot.pivotCacheId}"><items count="${items.length}">${items.map((label, k) => `<i x="${k}"${isOn(label) ? ' s="1"' : ''}/>`).join('')}</items></tabular></data>`;
    } else continue;
    n++;
    const cachePath = `xl/slicerCaches/slicerCache${n}.xml`;
    zip.file(
      cachePath,
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
        `<slicerCacheDefinition xmlns="${NS_X14}" xmlns:mc="${NS_MC}" mc:Ignorable="x" xmlns:x="${NS_MAIN}"${attrs({ name: s.cacheName, sourceName })}>${body}</slicerCacheDefinition>`,
    );
    await addContentTypeOverride(zip, `/${cachePath}`, CT_SLICER_CACHE);
    const rid = await addRelationship(zip, 'xl/workbook.xml', REL_SLICER_CACHE, `slicerCaches/slicerCache${n}.xml`);
    (kind === 'table' ? tableCaches : pivotCaches).push(`<x14:slicerCache r:id="${rid}"/>`);
    written.push({ slicer: s, kind });
  }
  if (!written.length) return;

  // Parties « segments » de chaque feuille (une par sorte, comme Excel) et extensions de feuille.
  let partIndex = 0;
  const bySheet = new Map<string, { slicer: NativeSlicer; kind: 'table' | 'pivot' }[]>();
  for (const w of written) bySheet.set(w.slicer.drawSheetId, [...(bySheet.get(w.slicer.drawSheetId) ?? []), w]);
  for (const [sheetId, list] of bySheet) {
    const sheetPath = sheetPathOf(sheetId);
    const file = sheetPath ? zip.file(sheetPath) : null;
    if (!sheetPath || !file) continue;
    let xml = await file.async('string');
    for (const kind of ['pivot', 'table'] as const) {
      const ofKind = list.filter((w) => w.kind === kind);
      if (!ofKind.length) continue;
      partIndex++;
      const path = `xl/slicers/slicer${partIndex}.xml`;
      zip.file(
        path,
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
          `<slicers xmlns="${NS_X14}" xmlns:mc="${NS_MC}" mc:Ignorable="x" xmlns:x="${NS_MAIN}">` +
          ofKind.map(({ slicer }) => `<slicer${attrs({ name: slicer.name, cache: slicer.cacheName, caption: slicer.caption, rowHeight: 241300 })}/>`).join('') +
          '</slicers>',
      );
      await addContentTypeOverride(zip, `/${path}`, CT_SLICER);
      const rid = await addRelationship(zip, sheetPath, REL_SLICER, `../slicers/slicer${partIndex}.xml`);
      const ext =
        kind === 'pivot'
          ? `<ext uri="${EXT.sheetPivotSlicers}" xmlns:x14="${NS_X14}"><x14:slicerList><x14:slicer r:id="${rid}"/></x14:slicerList></ext>`
          : `<ext uri="${EXT.sheetTableSlicers}" xmlns:x15="${NS_X15}"><x14:slicerList xmlns:x14="${NS_X14}"><x14:slicer r:id="${rid}"/></x14:slicerList></ext>`;
      if (!/xmlns:r=/.test(xml.slice(0, 2000))) xml = xml.replace(/<worksheet\b/, '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"');
      xml = appendExt(xml, 'worksheet', ext);
    }
    zip.file(sheetPath, xml);
  }

  // Classeur : noms définis et listes de caches.
  const wbFile = zip.file('xl/workbook.xml');
  if (!wbFile) return;
  let wbXml = await wbFile.async('string');
  wbXml = addDefinedNames(wbXml, written.map((w) => w.slicer.cacheName));
  if (!/xmlns:r=/.test(wbXml.slice(0, 1500))) wbXml = wbXml.replace(/<workbook\b/, '<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"');
  if (pivotCaches.length) wbXml = appendExt(wbXml, 'workbook', `<ext uri="${EXT.bookPivotCaches}" xmlns:x14="${NS_X14}"><x14:slicerCaches>${pivotCaches.join('')}</x14:slicerCaches></ext>`);
  if (tableCaches.length) wbXml = appendExt(wbXml, 'workbook', `<ext uri="${EXT.bookTableCaches}" xmlns:x15="${NS_X15}"><x15:slicerCaches xmlns:x14="${NS_X14}">${tableCaches.join('')}</x15:slicerCaches></ext>`);
  zip.file('xl/workbook.xml', wbXml);
}

/** Extension du cache d'un tableau croisé qui porte l'identifiant utilisé par les segments. */
export const pivotCacheIdExt = (id: number) => `<extLst><ext uri="${EXT.pivotCacheId}" xmlns:x14="${NS_X14}"><x14:pivotCacheDefinition pivotCacheId="${id}"/></ext></extLst>`;

async function writeSavedState(zip: JSZip, state: SavedState): Promise<void> {
  let n = 1;
  while (zip.file(`customXml/item${n}.xml`)) n++;
  const item = `customXml/item${n}.xml`;
  const props = `customXml/itemProps${n}.xml`;
  zip.file(item, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<grillon xmlns="${GRILLON_NS}">${escapeXml(JSON.stringify(state))}</grillon>`);
  zip.file(
    props,
    '<?xml version="1.0" encoding="UTF-8" standalone="no"?>\n' +
      `<ds:datastoreItem ds:itemID="{6C1F2A5E-4B7D-4E8A-9C3F-0D2B1A7E5F${String(n).padStart(2, '0')}}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:schemaRefs><ds:schemaRef ds:uri="${GRILLON_NS}"/></ds:schemaRefs></ds:datastoreItem>`,
  );
  await addRelationship(zip, item, REL_CUSTOM_XML_PROPS, `itemProps${n}.xml`);
  await addContentTypeOverride(zip, `/${props}`, CT_CUSTOM_XML_PROPS);
  const types = await zip.file('[Content_Types].xml')?.async('string');
  if (types && !/Extension="xml"/i.test(types)) zip.file('[Content_Types].xml', types.replace(/<Types\b[^>]*>/, (m) => `${m}<Default Extension="xml" ContentType="application/xml"/>`));
  await addRelationship(zip, 'xl/workbook.xml', REL_CUSTOM_XML, `../customXml/item${n}.xml`);
}

// --- Lecture ---

/** Segment lu dans un fichier : où il est dessiné et ce qu'il filtre. */
export interface ReadSlicer {
  /** Feuille (nom) où le segment est dessiné. */
  sheet: string;
  name: string;
  from: CellAnchor;
  to: CellAnchor;
  spec: SlicerSpec;
}

export interface ReadSlicersResult {
  slicers: ReadSlicer[];
  /** Tableaux à retirer (créés par Grillon pour porter un segment) : feuille, plage. */
  syntheticTables: { sheet: string; ref: string }[];
}

async function readSavedState(zip: JSZip): Promise<SavedState | null> {
  for (const rel of (await readRelationships(zip, 'xl/workbook.xml')).filter((r) => r.Type === REL_CUSTOM_XML)) {
    const xml = await zip.file(resolveRelTarget('xl/workbook.xml', rel.Target))?.async('string');
    if (!xml?.includes(GRILLON_NS)) continue;
    const text = /<grillon\b[^>]*>([\s\S]*)<\/grillon>/.exec(xml)?.[1];
    if (!text) continue;
    try {
      const state = JSON.parse(unescapeXml(text)) as SavedState;
      if (state?.version === 1 && Array.isArray(state.slicers)) return state;
    } catch {
      /* partie illisible : ignorée */
    }
  }
  return null;
}

const attr = (tag: string, name: string) => {
  const v = new RegExp(`\\s${name}="([^"]*)"`).exec(tag)?.[1];
  return v === undefined ? undefined : unescapeXml(v);
};

function anchorOf(anchorXml: string, tag: 'from' | 'to'): CellAnchor | null {
  const el = extractElement(anchorXml, tag);
  if (!el) return null;
  const n = (name: string) => Number(new RegExp(`<xdr:${name}>(-?\\d+)</xdr:${name}>`).exec(el.xml)?.[1] ?? 0);
  return { column: n('col'), columnOffset: Math.round(n('colOff') / EMU_PER_PX), row: n('row'), rowOffset: Math.round(n('rowOff') / EMU_PER_PX) };
}

/** Libellé d'un élément partagé de tableau croisé, comme Grillon l'affiche. */
function sharedLabels(fieldXml: string): string[] {
  const out: string[] = [];
  const shared = /<(?:\w+:)?sharedItems\b[^>]*>([\s\S]*?)<\/(?:\w+:)?sharedItems>/.exec(fieldXml)?.[1] ?? '';
  for (const m of shared.matchAll(/<(?:\w+:)?(s|n|m|b|d|e)\b([^>]*?)\/?>/g)) {
    const v = attr(m[2] ?? '', 'v');
    if (m[1] === 'm') out.push(EMPTY_LABEL);
    else if (m[1] === 'b') out.push(v === '1' || v === 'true' ? 'VRAI' : 'FAUX');
    else out.push(v ?? '');
  }
  return out;
}

/** Segments d'Excel (relus depuis les parties du fichier). */
async function readNativeSlicers(zip: JSZip): Promise<ReadSlicer[]> {
  const sheets = await sheetsWithIds(zip);
  const wbPath = 'xl/workbook.xml';
  const wbRels = await readRelationships(zip, wbPath);
  // Caches : nom → description.
  type Cache = { sourceName: string; table?: { id: number; column: number }; pivot?: { tabId: number; name: string; items: { x: number; s: boolean }[] } };
  const caches = new Map<string, Cache>();
  for (const rel of wbRels.filter((r) => r.Type === REL_SLICER_CACHE)) {
    const xml = await zip.file(resolveRelTarget(wbPath, rel.Target))?.async('string');
    if (!xml) continue;
    const head = /<(?:\w+:)?slicerCacheDefinition\b[^>]*>/.exec(xml)?.[0] ?? '';
    const name = attr(head, 'name');
    if (!name) continue;
    const cache: Cache = { sourceName: attr(head, 'sourceName') ?? '' };
    const tsc = /<(?:\w+:)?tableSlicerCache\b[^>]*>/.exec(xml)?.[0];
    if (tsc) cache.table = { id: Number(attr(tsc, 'tableId')), column: Number(attr(tsc, 'column')) };
    const pt = /<(?:\w+:)?pivotTable\b[^>]*>/.exec(xml)?.[0];
    if (pt) {
      cache.pivot = {
        tabId: Number(attr(pt, 'tabId')),
        name: attr(pt, 'name') ?? '',
        items: [...xml.matchAll(/<(?:\w+:)?i\b([^>]*)\/?>/g)].map((m) => ({ x: Number(attr(m[1], 'x')), s: attr(m[1], 's') === '1' || attr(m[1], 's') === 'true' })),
      };
    }
    caches.set(name, cache);
  }
  if (!caches.size) return [];
  // Tableaux : id → feuille, plage (sans la ligne de totaux), colonnes (id → position), filtre.
  const tables = new Map<number, { sheet: string; range: ReturnType<typeof rangeFromA1>; columnIds: number[]; filters: Map<number, string[]> }>();
  for (const sheet of sheets) {
    for (const rel of (await readRelationships(zip, sheet.path)).filter((r) => r.Type.endsWith('/table'))) {
      const xml = await zip.file(resolveRelTarget(sheet.path, rel.Target))?.async('string');
      const head = xml && /<(?:\w+:)?table\b[^>]*>/.exec(xml)?.[0];
      if (!xml || !head) continue;
      const ref = attr(head, 'ref');
      if (!ref) continue;
      const range = rangeFromA1(ref);
      range.endRow -= Number(attr(head, 'totalsRowCount') ?? 0) || 0;
      const filters = new Map<number, string[]>();
      for (const fc of extractAll(xml, 'filterColumn')) {
        const col = Number(attr(fc.xml, 'colId'));
        const labels = [...fc.xml.matchAll(/<(?:\w+:)?filter\b([^>]*)\/?>/g)].map((m) => attr(m[1], 'val') ?? '');
        if (/<(?:\w+:)?filters\b[^>]*\sblank="(?:1|true)"/.test(fc.xml)) labels.push(EMPTY_LABEL);
        if (/<(?:\w+:)?filters\b/.test(fc.xml)) filters.set(col, labels);
      }
      tables.set(Number(attr(head, 'id')), { sheet: sheet.name, range, columnIds: [...xml.matchAll(/<(?:\w+:)?tableColumn\b([^>]*)/g)].map((m) => Number(attr(m[1], 'id'))), filters });
    }
  }
  // Tableaux croisés : (feuille, nom) → libellés de chaque champ du cache.
  const pivotFields = async (sheetPath: string, name: string): Promise<{ names: string[]; labels: string[][] } | null> => {
    for (const rel of (await readRelationships(zip, sheetPath)).filter((r) => r.Type.endsWith('/pivotTable'))) {
      const tablePath = resolveRelTarget(sheetPath, rel.Target);
      const xml = await zip.file(tablePath)?.async('string');
      if (!xml || attr(/<(?:\w+:)?pivotTableDefinition\b[^>]*>/.exec(xml)?.[0] ?? '', 'name') !== name) continue;
      const cacheRel = (await readRelationships(zip, tablePath)).find((r) => r.Type.endsWith('/pivotCacheDefinition'));
      const cacheXml = cacheRel && (await zip.file(resolveRelTarget(tablePath, cacheRel.Target))?.async('string'));
      if (!cacheXml) return null;
      const blocks = [...cacheXml.matchAll(/<(?:\w+:)?cacheField\b[^>]*?(?:\/>|>[\s\S]*?<\/(?:\w+:)?cacheField>)/g)].map((m) => m[0]);
      return { names: blocks.map((b) => attr(/<(?:\w+:)?cacheField\b[^>]*>/.exec(b)?.[0] ?? '', 'name') ?? ''), labels: blocks.map(sharedLabels) };
    }
    return null;
  };

  const out: ReadSlicer[] = [];
  for (const sheet of sheets) {
    const rels = await readRelationships(zip, sheet.path);
    // Segments de la feuille : nom → (cache, légende).
    const defs = new Map<string, { cache: string; caption: string }>();
    for (const rel of rels.filter((r) => r.Type === REL_SLICER)) {
      const xml = await zip.file(resolveRelTarget(sheet.path, rel.Target))?.async('string');
      for (const m of xml?.matchAll(/<(?:\w+:)?slicer\b([^>]*)\/?>/g) ?? []) {
        const name = attr(m[1], 'name');
        if (name) defs.set(name, { cache: attr(m[1], 'cache') ?? '', caption: attr(m[1], 'caption') ?? name });
      }
    }
    if (!defs.size) continue;
    // Positions : cadres sle:slicer du dessin.
    const anchors = new Map<string, { from: CellAnchor; to: CellAnchor }>();
    for (const rel of rels.filter((r) => r.Type.endsWith('/drawing'))) {
      const xml = await zip.file(resolveRelTarget(sheet.path, rel.Target))?.async('string');
      for (const a of xml ? extractAll(xml, 'twoCellAnchor') : []) {
        const name = attr(/<(?:\w+:)?slicer\b[^>]*>/.exec(a.xml)?.[0] ?? '', 'name');
        const from = anchorOf(a.xml, 'from');
        const to = anchorOf(a.xml, 'to');
        if (name && from && to) anchors.set(name, { from, to });
      }
    }
    for (const [name, def] of defs) {
      const cache = caches.get(def.cache);
      const pos = anchors.get(name);
      if (!cache || !pos) continue;
      const base = { id: `slicer-${out.length}-${Date.now().toString(36)}`, type: 'slicer' as const, title: def.caption };
      if (cache.table) {
        const table = tables.get(cache.table.id);
        const col = table?.columnIds.indexOf(cache.table.column) ?? -1;
        if (!table || col < 0) continue;
        out.push({ sheet: sheet.name, name, ...pos, spec: { ...base, target: { kind: 'table', sheet: table.sheet, ref: rangeToA1(table.range), col }, selected: table.filters.get(col) ?? null } });
      } else if (cache.pivot) {
        const pivotSheet = sheets.find((s) => s.sheetId === cache.pivot!.tabId);
        const fields = pivotSheet && (await pivotFields(pivotSheet.path, cache.pivot.name));
        const col = fields?.names.indexOf(cache.sourceName) ?? -1;
        if (!pivotSheet || !fields || col < 0) continue;
        const labels = fields.labels[col] ?? [];
        const chosen = cache.pivot.items.filter((i) => i.s).map((i) => labels[i.x]).filter((l): l is string => l !== undefined);
        const all = cache.pivot.items.length === 0 || chosen.length === cache.pivot.items.length;
        out.push({ sheet: sheet.name, name, ...pos, spec: { ...base, target: { kind: 'pivot', sheet: pivotSheet.name, col }, selected: all ? null : chosen } });
      }
    }
  }
  return out;
}

/**
 * Segments et chronologies d'un fichier : l'état de Grillon s'il est là (mis à jour par ce qu'Excel a pu changer :
 * choix, position, segments supprimés), sinon les segments d'Excel.
 */
export async function readSlicers(zip: JSZip): Promise<ReadSlicersResult> {
  const native = await readNativeSlicers(zip);
  const saved = await readSavedState(zip);
  if (!saved) return { slicers: native, syntheticTables: [] };
  const byName = new Map(native.map((s) => [s.name, s]));
  const slicers: ReadSlicer[] = [];
  for (const entry of saved.slicers) {
    const spec = entry.item.data;
    if (!spec) continue;
    if (!entry.native) {
      slicers.push({ sheet: entry.sheet, name: spec.title, from: entry.item.sheetTransform.from, to: entry.item.sheetTransform.to, spec });
      continue;
    }
    const n = byName.get(entry.native);
    if (!n) continue; // supprimé dans Excel
    byName.delete(entry.native);
    slicers.push({ ...n, spec: { ...spec, title: n.spec.title, selected: n.spec.selected } });
  }
  // Segments ajoutés dans Excel.
  slicers.push(...byName.values());
  return { slicers, syntheticTables: saved.syntheticTables };
}
