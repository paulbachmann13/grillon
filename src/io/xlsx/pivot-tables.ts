/**
 * Tableaux croisés dynamiques ⇄ .xlsx (xl/pivotCache/* et xl/pivotTables/*).
 *
 * À l'export, chaque tableau croisé de Grillon devient un vrai tableau croisé Excel : cache
 * (définition + enregistrements) et définition du tableau (champs en lignes, colonnes, valeurs,
 * filtres). Le cache est marqué « actualiser à l'ouverture » : Excel recalcule la mise en page.
 * À l'import, la définition est relue pour que Grillon puisse l'actualiser et la modifier.
 */
import type JSZip from 'jszip';
import { cellValue } from '../../data/cells';
import type { PivotConfig, PivotDefinition, PivotFunction } from '../../data/pivot';
import type { ICellData, IWorkbookData } from '../../model/univer-types';
import { addContentTypeOverride, addRelationship, escapeXml, listSheets, parseXml, readRelationships, resolveRelTarget } from './ooxml';
import { pivotCacheIdExt, type WrittenPivot } from './slicers';
import { columnToLetters, rangeFromA1 } from './units';

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL = {
  cacheDefinition: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition',
  cacheRecords: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords',
  pivotTable: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable',
};
const CT = {
  cacheDefinition: 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml',
  cacheRecords: 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml',
  pivotTable: 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml',
};

const FN_NAMES: Record<PivotFunction, string> = { sum: 'Somme de', average: 'Moyenne de', count: 'Nombre de', min: 'Min de', max: 'Max de' };
const FN_TO_XML: Record<PivotFunction, string | undefined> = { sum: undefined, average: 'average', count: 'count', min: 'min', max: 'max' };
const XML_TO_FN: Record<string, PivotFunction> = { sum: 'sum', average: 'average', count: 'count', countNums: 'count', min: 'min', max: 'max' };

type Scalar = string | number | boolean | null;

/** Libellé d'un élément, tel que Grillon l'affiche dans les filtres. */
const labelOf = (v: Scalar): string => (v === null ? '(vide)' : typeof v === 'boolean' ? (v ? 'VRAI' : 'FAUX') : String(v));
const ref = (r1: number, c1: number, r2: number, c2: number) => `${columnToLetters(c1)}${r1 + 1}:${columnToLetters(c2)}${r2 + 1}`;

function sortItems(values: Scalar[]): Scalar[] {
  const rank = (v: Scalar) => (v === null ? 3 : typeof v === 'number' ? 0 : typeof v === 'boolean' ? 2 : 1);
  return [...values].sort((a, b) => {
    const d = rank(a) - rank(b);
    if (d) return d;
    if (typeof a === 'number' && typeof b === 'number') return a - b;
    return String(a).localeCompare(String(b), 'fr');
  });
}

interface FieldInfo {
  name: string;
  values: Scalar[];
  /** Éléments distincts (champs placés en lignes, colonnes ou filtres). */
  items?: Scalar[];
}

function itemXml(v: Scalar): string {
  if (v === null) return '<m/>';
  if (typeof v === 'number') return `<n v="${v}"/>`;
  if (typeof v === 'boolean') return `<b v="${v ? 1 : 0}"/>`;
  return `<s v="${escapeXml(v)}"/>`;
}

function sharedItemsXml(f: FieldInfo): string {
  const nums = f.values.filter((v): v is number => typeof v === 'number');
  const hasString = f.values.some((v) => typeof v === 'string' || typeof v === 'boolean');
  const hasBlank = f.values.some((v) => v === null);
  const attrs: string[] = [];
  if (nums.length) {
    if (!hasString) attrs.push('containsSemiMixedTypes="0"', 'containsString="0"');
    attrs.push('containsNumber="1"');
    if (nums.every(Number.isInteger)) attrs.push('containsInteger="1"');
    attrs.push(`minValue="${Math.min(...nums)}"`, `maxValue="${Math.max(...nums)}"`);
  }
  if (hasBlank) attrs.push('containsBlank="1"');
  if (!f.items) return `<sharedItems${attrs.length ? ' ' + attrs.join(' ') : ''}/>`;
  return `<sharedItems${attrs.length ? ' ' + attrs.join(' ') : ''} count="${f.items.length}">${f.items.map(itemXml).join('')}</sharedItems>`;
}

/** Lit la plage source d'un tableau croisé dans le classeur. */
function sourceFields(wb: IWorkbookData, def: PivotDefinition, slicerCols: number[] = []): { fields: FieldInfo[]; rows: number } | null {
  const sheet = Object.values(wb.sheets).find((s) => s.name === def.sourceSheet);
  if (!sheet) return null;
  let r: ReturnType<typeof rangeFromA1>;
  try {
    r = rangeFromA1(def.sourceRange.replace(/\$/g, ''));
  } catch {
    return null;
  }
  const cell = (row: number, col: number) => sheet.cellData?.[row]?.[col] as ICellData | undefined;
  // Source en colonnes entières (A1:F1048576, fréquent dans Excel) : seulement jusqu'à la dernière ligne remplie.
  const lastRow = Math.max(-1, ...Object.keys(sheet.cellData ?? {}).map(Number).filter((row) => Object.keys(sheet.cellData![row] ?? {}).length));
  r = { ...r, endRow: Math.max(r.startRow, Math.min(r.endRow, lastRow)) };
  const fields: FieldInfo[] = [];
  const used = new Set([...def.config.rows, ...def.config.columns, ...def.config.filters.map((f) => f.col), ...slicerCols]);
  for (let c = r.startColumn; c <= r.endColumn; c++) {
    const i = c - r.startColumn;
    const header = cellValue(cell(r.startRow, c));
    const values: Scalar[] = [];
    for (let row = r.startRow + 1; row <= r.endRow; row++) values.push(cellValue(cell(row, c)));
    const field: FieldInfo = { name: header === null ? `Colonne ${i + 1}` : String(header), values };
    if (used.has(i)) {
      const seen = new Map<string, Scalar>();
      for (const v of values) seen.set(`${typeof v}:${String(v)}`, v);
      field.items = sortItems([...seen.values()]);
    }
    fields.push(field);
  }
  // Noms de champs uniques (exigé par Excel).
  const names = new Set<string>();
  for (const f of fields) {
    let n = f.name;
    for (let k = 2; names.has(n); k++) n = `${f.name}${k}`;
    names.add(n);
    f.name = n;
  }
  return { fields, rows: r.endRow - r.startRow };
}

function cacheDefinitionXml(def: PivotDefinition, fields: FieldInfo[], records: number, ext = ''): string {
  const cacheFields = fields.map((f) => `<cacheField name="${escapeXml(f.name)}" numFmtId="0">${sharedItemsXml(f)}</cacheField>`).join('');
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    `<pivotCacheDefinition xmlns="${NS}" xmlns:r="${NS_R}" r:id="rId1" refreshOnLoad="1" createdVersion="6" refreshedVersion="6" minRefreshableVersion="3" recordCount="${records}">` +
    `<cacheSource type="worksheet"><worksheetSource ref="${escapeXml(def.sourceRange.replace(/\$/g, ''))}" sheet="${escapeXml(def.sourceSheet)}"/></cacheSource>` +
    `<cacheFields count="${fields.length}">${cacheFields}</cacheFields>${ext}</pivotCacheDefinition>`
  );
}

function cacheRecordsXml(fields: FieldInfo[], rows: number): string {
  const index = fields.map((f) => (f.items ? new Map(f.items.map((v, i) => [`${typeof v}:${String(v)}`, i])) : null));
  let out = '';
  for (let r = 0; r < rows; r++) {
    out += '<r>';
    fields.forEach((f, c) => {
      const v = f.values[r] ?? null;
      const map = index[c];
      out += map ? `<x v="${map.get(`${typeof v}:${String(v)}`) ?? 0}"/>` : itemXml(v);
    });
    out += '</r>';
  }
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + `<pivotCacheRecords xmlns="${NS}" xmlns:r="${NS_R}" count="${rows}">${out}</pivotCacheRecords>`;
}

export function pivotTableXml(def: PivotDefinition, fields: FieldInfo[], cacheId: number, name: string, slicerCols: number[] = []): string {
  const cfg = def.config;
  const top = def.outRow ?? 2;
  const left = def.outCol ?? 0;
  const location = ref(top, left, top + Math.max(1, def.outRows) - 1, left + Math.max(1, def.outCols) - 1);
  const pivotFields = fields
    .map((f, i) => {
      const axis = cfg.rows.includes(i) ? 'axisRow' : cfg.columns.includes(i) ? 'axisCol' : cfg.filters.some((x) => x.col === i) ? 'axisPage' : undefined;
      const data = cfg.values.some((v) => v.col === i);
      const filter = cfg.filters.find((x) => x.col === i);
      const attrs = `${axis ? ` axis="${axis}"` : ''}${data ? ' dataField="1"' : ''}${filter ? ' multipleItemSelectionAllowed="1"' : ''} showAll="0"`;
      // Champ d'un segment hors des axes : ses éléments sont décrits quand même (comme Excel).
      if (!f.items || (!axis && !slicerCols.includes(i))) return `<pivotField${attrs}/>`;
      const items = f.items.map((v, k) => `<item${filter && !filter.include.includes(labelOf(v)) ? ' h="1"' : ''} x="${k}"/>`).join('');
      return `<pivotField${attrs}><items count="${f.items.length + 1}">${items}<item t="default"/></items></pivotField>`;
    })
    .join('');
  const list = (tag: string, xs: number[]) => (xs.length ? `<${tag} count="${xs.length}">${xs.map((x) => `<field x="${x}"/>`).join('')}</${tag}>` : '');
  const colFields = [...cfg.columns, ...(cfg.values.length > 1 ? [-2] : [])];
  const pages = cfg.filters.length ? `<pageFields count="${cfg.filters.length}">${cfg.filters.map((f) => `<pageField fld="${f.col}" hier="-1"/>`).join('')}</pageFields>` : '';
  const dataFields = cfg.values.length
    ? `<dataFields count="${cfg.values.length}">${cfg.values
        .map((v) => {
          const sub = FN_TO_XML[v.fn];
          return `<dataField name="${escapeXml(`${FN_NAMES[v.fn]} ${fields[v.col]?.name ?? ''}`)}" fld="${v.col}"${sub ? ` subtotal="${sub}"` : ''} baseField="0" baseItem="0"/>`;
        })
        .join('')}</dataFields>`
    : '';
  const pageCounts = cfg.filters.length ? ` rowPageCount="${cfg.filters.length}" colPageCount="1"` : '';
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    `<pivotTableDefinition xmlns="${NS}" name="${escapeXml(name)}" cacheId="${cacheId}" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Valeurs" updatedVersion="6" minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" createdVersion="6" indent="0" outline="1" outlineData="1" multipleFieldFilters="0">` +
    `<location ref="${location}" firstHeaderRow="1" firstDataRow="${colFields.length ? 2 : 1}" firstDataCol="${Math.max(1, cfg.rows.length)}"${pageCounts}/>` +
    `<pivotFields count="${fields.length}">${pivotFields}</pivotFields>` +
    list('rowFields', cfg.rows) +
    list('colFields', colFields) +
    pages +
    dataFields +
    '<pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/>' +
    '</pivotTableDefinition>'
  );
}

/** Ajoute <pivotCaches> à xl/workbook.xml, au bon endroit du schéma. */
function addWorkbookPivotCaches(workbookXml: string, caches: { id: number; rid: string }[]): string {
  const fragment = `<pivotCaches>${caches.map((c) => `<pivotCache cacheId="${c.id}" r:id="${c.rid}"/>`).join('')}</pivotCaches>`;
  let xml = workbookXml;
  if (!/xmlns:r=/.test(xml.slice(0, 1500))) xml = xml.replace(/<workbook\b/, `<workbook xmlns:r="${NS_R}"`);
  const after = [/<calcPr\b[^>]*\/>/, /<\/calcPr>/, /<\/definedNames>/, /<\/sheets>/];
  for (const re of after) {
    const m = re.exec(xml);
    if (m) return xml.slice(0, m.index + m[0].length) + fragment + xml.slice(m.index + m[0].length);
  }
  return xml.replace('</workbook>', `${fragment}</workbook>`);
}

export interface PivotToWrite {
  sheetPath: string;
  def: PivotDefinition;
  /** Colonnes de la source filtrées par des segments. */
  slicerCols?: number[];
}

/** Écrit les tableaux croisés (cache, définition, relations) dans le paquet ; renvoie ce qu'il faut aux segments. */
export async function writePivotTables(zip: JSZip, wb: IWorkbookData, items: PivotToWrite[]): Promise<Map<string, WrittenPivot>> {
  const caches: { id: number; rid: string }[] = [];
  const written = new Map<string, WrittenPivot>();
  let n = 0;
  for (const item of items) {
    const slicerCols = item.slicerCols ?? [];
    const src = sourceFields(wb, item.def, slicerCols);
    if (!src || !item.def.config.values.length) continue;
    n++;
    const cacheId = n;
    const name = item.def.name ?? `TCD${n}`;
    const defPath = `xl/pivotCache/pivotCacheDefinition${n}.xml`;
    const recPath = `xl/pivotCache/pivotCacheRecords${n}.xml`;
    const tablePath = `xl/pivotTables/pivotTable${n}.xml`;
    zip.file(defPath, cacheDefinitionXml(item.def, src.fields, src.rows, slicerCols.length ? pivotCacheIdExt(cacheId) : ''));
    zip.file(recPath, cacheRecordsXml(src.fields, src.rows));
    zip.file(tablePath, pivotTableXml(item.def, src.fields, cacheId, name, slicerCols));
    written.set(item.sheetPath, {
      name,
      pivotCacheId: cacheId,
      items: src.fields.map((f) => f.items?.map(labelOf)),
      fieldNames: src.fields.map((f) => f.name),
      hidden: new Map(item.def.config.filters.map((f) => [f.col, new Set((src.fields[f.col]?.items ?? []).map(labelOf).filter((l) => !f.include.includes(l)))])),
    });
    await addRelationship(zip, defPath, REL.cacheRecords, `pivotCacheRecords${n}.xml`);
    await addRelationship(zip, tablePath, REL.cacheDefinition, `../pivotCache/pivotCacheDefinition${n}.xml`);
    await addRelationship(zip, item.sheetPath, REL.pivotTable, `../pivotTables/pivotTable${n}.xml`);
    const rid = await addRelationship(zip, 'xl/workbook.xml', REL.cacheDefinition, `pivotCache/pivotCacheDefinition${n}.xml`);
    caches.push({ id: cacheId, rid });
    await addContentTypeOverride(zip, `/${defPath}`, CT.cacheDefinition);
    await addContentTypeOverride(zip, `/${recPath}`, CT.cacheRecords);
    await addContentTypeOverride(zip, `/${tablePath}`, CT.pivotTable);
  }
  if (!caches.length) return written;
  const wbFile = zip.file('xl/workbook.xml');
  if (wbFile) zip.file('xl/workbook.xml', addWorkbookPivotCaches(await wbFile.async('string'), caches));
  return written;
}

type Xml = Record<string, unknown>;
const arr = <T>(v: T | T[] | undefined): T[] => (v === undefined || v === '' ? [] : Array.isArray(v) ? v : [v]);

/** Valeurs des éléments partagés, dans l'ordre du document. */
function orderedItems(fieldXml: string): Scalar[] {
  const out: Scalar[] = [];
  for (const m of fieldXml.matchAll(/<(?:\w+:)?(s|n|m|b|d|e)\b([^>]*?)\/?>(?:<\/(?:\w+:)?\1>)?/g)) {
    const kind = m[1];
    const v = /\bv="([^"]*)"/.exec(m[2] ?? '')?.[1];
    if (kind === 'm') out.push(null);
    else if (kind === 'n') out.push(Number(v));
    else if (kind === 'b') out.push(v === '1' || v === 'true');
    else out.push((v ?? '').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&'));
  }
  return out;
}

/** Nom défini ou nom de tableau → feuille et plage (avec la ligne d'en-têtes). */
async function resolveNamedSource(zip: JSZip, name: string): Promise<{ ref: string; sheet: string } | null> {
  const wbXml = (await zip.file('xl/workbook.xml')?.async('string')) ?? '';
  const esc = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const def = new RegExp(`<(?:\\w+:)?definedName\\b[^>]*\\bname="${esc}"[^>]*>([^<]*)<`, 'i').exec(wbXml)?.[1];
  const m = def && /^'?(.+?)'?!\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)$/.exec(def.replace(/&apos;/g, "'").replace(/''/g, "'").trim());
  if (m) return { sheet: m[1], ref: `${m[2]}${m[3]}:${m[4]}${m[5]}` };
  for (const sheet of await listSheets(zip)) {
    for (const rel of (await readRelationships(zip, sheet.path)).filter((r) => r.Type.endsWith('/table'))) {
      const xml = await zip.file(resolveRelTarget(sheet.path, rel.Target))?.async('string');
      const head = xml && /<(?:\w+:)?table\b[^>]*>/.exec(xml)?.[0];
      if (!head) continue;
      const tableName = /\sdisplayName="([^"]*)"/.exec(head)?.[1] ?? /\sname="([^"]*)"/.exec(head)?.[1];
      const ref = /\sref="([^"]*)"/.exec(head)?.[1];
      if (ref && tableName?.toLowerCase() === name.toLowerCase()) {
        // Sans la ligne de totaux.
        const totals = Number(/\stotalsRowCount="(\d+)"/.exec(head)?.[1] ?? 0);
        const r = rangeFromA1(ref);
        return { sheet: sheet.name, ref: `${columnToLetters(r.startColumn)}${r.startRow + 1}:${columnToLetters(r.endColumn)}${r.endRow + 1 - totals}` };
      }
    }
  }
  return null;
}

/** Lit les tableaux croisés de chaque feuille (chemin de la feuille → définitions). */
export async function readPivotTables(zip: JSZip): Promise<Map<string, PivotDefinition[]>> {
  const out = new Map<string, PivotDefinition[]>();
  const wbXml = await zip.file('xl/workbook.xml')?.async('string');
  if (!wbXml || !/pivotCache/.test(wbXml)) return out;
  const wbRels = new Map((await readRelationships(zip, 'xl/workbook.xml')).map((r) => [r.Id, r]));
  const cachePaths = new Map<string, string>();
  for (const m of wbXml.matchAll(/<(?:\w+:)?pivotCache\b[^>]*cacheId="(\d+)"[^>]*r:id="([^"]+)"/g)) {
    const rel = wbRels.get(m[2]!);
    if (rel) cachePaths.set(m[1]!, resolveRelTarget('xl/workbook.xml', rel.Target));
  }
  for (const sheet of await listSheets(zip)) {
    for (const rel of (await readRelationships(zip, sheet.path)).filter((r) => r.Type === REL.pivotTable)) {
      const tablePath = resolveRelTarget(sheet.path, rel.Target);
      const tableXml = await zip.file(tablePath)?.async('string');
      if (!tableXml) continue;
      const table = parseXml<{ pivotTableDefinition?: Xml }>(tableXml).pivotTableDefinition;
      if (!table) continue;
      const cachePath = cachePaths.get(String(table.cacheId));
      const cacheXml = cachePath ? await zip.file(cachePath)?.async('string') : undefined;
      if (!cacheXml) continue;
      const cache = parseXml<{ pivotCacheDefinition?: Xml }>(cacheXml).pivotCacheDefinition;
      const ws = (cache?.cacheSource as Xml | undefined)?.worksheetSource as { ref?: string; sheet?: string; name?: string } | undefined;
      // Source nommée (nom défini ou tableau : « Tableau1 », ce qu'Excel écrit pour un tableau croisé fait d'un tableau).
      const named = !ws?.ref && ws?.name ? await resolveNamedSource(zip, String(ws.name)) : null;
      if (named) Object.assign(ws!, named);
      if (!ws?.ref) continue; // source externe : non prise en charge
      const fieldBlocks = [...cacheXml.matchAll(/<(?:\w+:)?cacheField\b[^>]*?(?:\/>|>[\s\S]*?<\/(?:\w+:)?cacheField>)/g)].map((m) => m[0]);
      const items = fieldBlocks.map((b) => {
        const shared = /<(?:\w+:)?sharedItems\b[^>]*?(?:\/>|>[\s\S]*?<\/(?:\w+:)?sharedItems>)/.exec(b)?.[0] ?? '';
        return orderedItems(shared.replace(/^<(?:\w+:)?sharedItems\b[^>]*>/, ''));
      });
      const pivotFields = arr(((table.pivotFields as Xml | undefined)?.pivotField) as Xml | Xml[] | undefined);
      const fieldsOf = (node: unknown) => arr(((node as Xml | undefined)?.field) as Xml | Xml[] | undefined).map((f) => Number(f.x)).filter((x) => x >= 0);
      const config: PivotConfig = {
        rows: fieldsOf(table.rowFields),
        columns: fieldsOf(table.colFields),
        values: arr(((table.dataFields as Xml | undefined)?.dataField) as Xml | Xml[] | undefined).map((d) => ({ col: Number(d.fld), fn: XML_TO_FN[String(d.subtotal ?? 'sum')] ?? 'sum' })),
        filters: arr(((table.pageFields as Xml | undefined)?.pageField) as Xml | Xml[] | undefined).map((p) => {
          const col = Number(p.fld);
          const shared = items[col] ?? [];
          const hidden = new Set(
            arr(((pivotFields[col]?.items as Xml | undefined)?.item) as Xml | Xml[] | undefined)
              .filter((it) => String(it.h) === '1' || String(it.h) === 'true')
              .map((it) => Number(it.x)),
          );
          return { col, include: shared.filter((_, k) => !hidden.has(k)).map(labelOf) };
        }),
      };
      const loc = rangeFromA1(String(((table.location as Xml | undefined)?.ref) ?? 'A3'));
      const def: PivotDefinition = {
        sourceSheet: String(ws.sheet ?? sheet.name),
        sourceRange: String(ws.ref).replace(/\$/g, ''),
        config,
        outRows: loc.endRow - loc.startRow + 1,
        outCols: loc.endColumn - loc.startColumn + 1,
        outRow: loc.startRow,
        outCol: loc.startColumn,
        name: String(table.name ?? ''),
      };
      out.set(sheet.path, [...(out.get(sheet.path) ?? []), def]);
    }
  }
  return out;
}

