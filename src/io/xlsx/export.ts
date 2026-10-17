/**
 * Export IWorkbookData (modèle Univer) → .xlsx via ExcelJS, puis post-traitement XML
 * pour ce qu'ExcelJS ne sait pas écrire (critères de filtre, tableaux, stopIfTrue).
 */
import ExcelJS from 'exceljs';
import { buildAutoFilterXml, FILTER_RESOURCE, type UniverAutoFilter, type UniverFilterResource } from './auto-filter';
import { chartsOfSheet, DRAWING_RESOURCE, writeSheetCharts, type DrawingResource, type SheetChart } from './charts';
import { DV_RESOURCE, exportDataValidations, type UniverDvResource } from './data-validation';
import { exportNotes, NOTE_RESOURCE, type UniverNoteResource } from './notes';
import { insertWorksheetElement, listSheets, loadZip, zipToBytes } from './ooxml';
import { TABLE_RESOURCE, writeTables, type TableToWrite, type UniverTable, type UniverTableResource } from './tables';
import {
  CellValueType,
  DEFINED_NAME_RESOURCE,
  GRILLON_CUSTOM_KEY,
  type ICellData,
  type IStyleData,
  type IWorkbookData,
  type IWorksheetData,
} from '../../model/univer-types';
import { cssToArgb } from './colors';
import { CF_RESOURCE, exportConditionalFormats, type UniverCfResource } from './conditional-formatting';
import { documentToPlainText, documentToRichText } from './rich-text';
import { styleFromUniver } from './styles';
import { EXCEL_DEFAULT_COL_WIDTH, EXCEL_DEFAULT_ROW_HEIGHT_PT, pxToExcelWidth, pxToPt, rangeToA1 } from './units';
import { toExcelPageSetup, type PageSetup } from '../../data/page-setup';
import { sparklinesOfSheet, writeSparklines, type SheetSparkline } from './sparklines';
import type { PivotDefinition } from '../../data/pivot';
import { writePivotTables, type PivotToWrite } from './pivot-tables';
import { sheetProtectionXml, type SheetProtection } from '../../data/sheet-protection';

/** Niveau de plan stocké dans les métadonnées d'une ligne ou colonne. */
function outlineLevel(custom: Record<string, unknown> | null | undefined | void): number {
  const level = (custom?.[GRILLON_CUSTOM_KEY] as { outline?: number } | undefined)?.outline;
  return typeof level === 'number' && level > 0 ? Math.min(7, Math.round(level)) : 0;
}

/** Résout un style Univer (objet ou identifiant dans workbook.styles). */
function resolveStyle(wb: IWorkbookData, s: ICellData['s']): IStyleData | null {
  if (!s) return null;
  if (typeof s === 'string') return wb.styles?.[s] ?? null;
  return s;
}

/** Les noms de feuille Excel : ≤ 31 caractères, sans []:*?/\ et uniques. */
export function sanitizeSheetName(name: string, used: Set<string>): string {
  let base = (name || 'Feuille').replace(/[[\]:*?/\\]/g, '_').slice(0, 31).replace(/^'+|'+$/g, '') || 'Feuille';
  let candidate = base;
  let i = 2;
  while (used.has(candidate.toLowerCase())) {
    const suffix = ` (${i++})`;
    candidate = base.slice(0, 31 - suffix.length) + suffix;
  }
  used.add(candidate.toLowerCase());
  return candidate;
}

function cellValue(cell: ICellData): ExcelJS.CellValue {
  const { v, t } = cell;
  if (v === null || v === undefined || v === '') return null;
  if (t === CellValueType.BOOLEAN) return v === true || v === 1 || v === '1' || String(v).toUpperCase() === 'TRUE';
  if (t === CellValueType.NUMBER || (t === undefined && typeof v === 'number')) {
    const n = typeof v === 'number' ? v : Number(v);
    return Number.isFinite(n) ? n : String(v);
  }
  if (t === CellValueType.STRING || t === CellValueType.FORCE_STRING) {
    const s = String(v);
    if (/^#(NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|SPILL!|CALC!)$/.test(s)) return { error: s } as ExcelJS.CellErrorValue;
    return s;
  }
  return typeof v === 'boolean' ? v : v;
}

function writeSheet(
  wb: IWorkbookData,
  xlWb: ExcelJS.Workbook,
  sheet: Partial<IWorksheetData>,
  name: string,
  filteredOutRows: number[] = [],
): ExcelJS.Worksheet {
  const defaultColWidth = sheet.defaultColumnWidth ? pxToExcelWidth(sheet.defaultColumnWidth) : EXCEL_DEFAULT_COL_WIDTH;
  const defaultRowHeight = sheet.defaultRowHeight ? pxToPt(sheet.defaultRowHeight) : EXCEL_DEFAULT_ROW_HEIGHT_PT;

  const freeze = sheet.freeze;
  const frozen = !!freeze && (freeze.xSplit > 0 || freeze.ySplit > 0) && freeze.startRow >= 0 && freeze.startColumn >= 0;
  const view: Partial<ExcelJS.WorksheetView> = frozen
    ? { state: 'frozen', xSplit: freeze!.xSplit, ySplit: freeze!.ySplit, topLeftCell: undefined }
    : { state: 'normal' };
  if (sheet.showGridlines === 0) view.showGridLines = false;
  if (sheet.zoomRatio && sheet.zoomRatio !== 1) view.zoomScale = Math.round(sheet.zoomRatio * 100);
  if (sheet.rightToLeft) view.rightToLeft = true;

  const ws = xlWb.addWorksheet(name, {
    properties: {
      defaultColWidth,
      defaultRowHeight,
      tabColor: sheet.tabColor ? ({ argb: cssToArgb(sheet.tabColor) } as never) : undefined,
    } as Partial<ExcelJS.WorksheetProperties>,
    views: [view],
    state: sheet.hidden ? 'hidden' : 'visible',
  });

  // Mise en page (papier, marges, échelle, zone d'impression, en-têtes et pieds de page).
  const page = (sheet.custom as { grillon?: { page?: PageSetup } } | undefined)?.grillon?.page;
  if (page) {
    const x = toExcelPageSetup(page);
    ws.pageSetup = x.pageSetup as unknown as ExcelJS.PageSetup;
    if (x.oddHeader || x.oddFooter) ws.headerFooter = { oddHeader: x.oddHeader, oddFooter: x.oddFooter };
  }

  let maxRowLevel = 0;
  let maxColLevel = 0;

  // Colonnes : largeur, masquage, niveau de plan, style de colonne.
  for (const [k, cd] of Object.entries(sheet.columnData ?? {})) {
    if (!cd) continue;
    const col = ws.getColumn(Number(k) + 1);
    if (cd.w !== undefined) col.width = pxToExcelWidth(cd.w);
    if (cd.hd) col.hidden = true;
    const colLevel = outlineLevel(cd.custom);
    if (colLevel) {
      col.outlineLevel = colLevel;
      maxColLevel = Math.max(maxColLevel, colLevel);
    }
    const st = styleFromUniver(resolveStyle(wb, cd.s));
    if (st) col.style = st as ExcelJS.Style;
  }

  // Lignes : hauteur, masquage, style.
  for (const [k, rd] of Object.entries(sheet.rowData ?? {})) {
    if (!rd) continue;
    const row = ws.getRow(Number(k) + 1);
    if (rd.h !== undefined && rd.h !== sheet.defaultRowHeight) row.height = pxToPt(rd.h);
    if (rd.hd) row.hidden = true;
    const rowLevel = outlineLevel(rd.custom);
    if (rowLevel) {
      row.outlineLevel = rowLevel;
      maxRowLevel = Math.max(maxRowLevel, rowLevel);
    }
    const st = styleFromUniver(resolveStyle(wb, rd.s));
    if (st) (row as unknown as { style: Partial<ExcelJS.Style> }).style = st;
  }
  // Excel ne réapplique pas les filtres à l'ouverture : les lignes filtrées doivent être masquées.
  for (const r of filteredOutRows) ws.getRow(r + 1).hidden = true;
  if (maxRowLevel) ws.properties.outlineLevelRow = maxRowLevel;
  if (maxColLevel) ws.properties.outlineLevelCol = maxColLevel;

  // Cache des styles convertis (beaucoup de cellules partagent un même identifiant).
  const styleCache = new Map<string, Partial<ExcelJS.Style> | null>();
  const toXlStyle = (s: ICellData['s']) => {
    if (!s) return null;
    if (typeof s === 'string') {
      if (!styleCache.has(s)) styleCache.set(s, styleFromUniver(resolveStyle(wb, s)));
      return styleCache.get(s)!;
    }
    return styleFromUniver(s);
  };

  for (const [rk, cols] of Object.entries(sheet.cellData ?? {})) {
    if (!cols) continue;
    const row = ws.getRow(Number(rk) + 1);
    for (const [ck, cell] of Object.entries(cols) as [string, ICellData | undefined][]) {
      if (!cell) continue;
      const xc = row.getCell(Number(ck) + 1);
      const f = typeof cell.f === 'string' && cell.f.startsWith('=') ? cell.f.slice(1) : undefined;
      let v = cellValue(cell);
      if (v === null && cell.p) v = documentToPlainText(cell.p) || null;
      const rich = !f && cell.p ? documentToRichText(cell.p) : null;
      const link = typeof cell.custom?.link === 'string' ? (cell.custom.link as string) : undefined;
      if (f) {
        xc.value = { formula: f, result: v ?? undefined } as ExcelJS.CellFormulaValue;
      } else if (rich) {
        xc.value = { richText: rich };
      } else if (link) {
        xc.value = { text: String(v ?? link), hyperlink: link };
      } else if (v !== null) {
        xc.value = v;
      }
      const st = toXlStyle(cell.s);
      if (st) xc.style = st as ExcelJS.Style;
    }
  }

  for (const m of sheet.mergeData ?? []) {
    if (m.endRow === m.startRow && m.endColumn === m.startColumn) continue;
    ws.mergeCells(rangeToA1(m));
  }
  return ws;
}

/** Lit une ressource de plugin Univer (JSON). */
function readResource<T>(wb: IWorkbookData, name: string): T | undefined {
  const data = wb.resources?.find((r) => r.name === name)?.data;
  if (!data) return undefined;
  try {
    return JSON.parse(data) as T;
  } catch {
    return undefined;
  }
}

interface StoredDefinedName {
  id: string;
  name: string;
  formulaOrRefString: string;
}

/** Ce qu'il reste à écrire directement dans le XML de chaque feuille, après ExcelJS. */
interface SheetPostProcess {
  name: string;
  charts: SheetChart[];
  sparklines: SheetSparkline[];
  pivot?: PivotDefinition;
  protection?: SheetProtection;
  autoFilter?: UniverAutoFilter;
  tables: TableToWrite[];
  stopIfTruePriorities: number[];
}

/** Texte d'une cellule d'en-tête de tableau. */
function headerText(cell: ICellData | undefined): string {
  if (!cell) return '';
  if (cell.v !== undefined && cell.v !== null && cell.v !== '') return String(cell.v);
  return cell.p ? documentToPlainText(cell.p) : '';
}

/**
 * Prépare les tableaux d'une feuille : Excel exige que le nom de chaque colonne soit
 * exactement le texte (unique, non vide) de la cellule d'en-tête.
 */
/** Fonctions de la ligne de totaux d'Excel, par numéro de SUBTOTAL. */
const TOTALS_FUNCTIONS: Record<number, string> = {
  101: 'average',
  102: 'countNums',
  103: 'count',
  104: 'max',
  105: 'min',
  107: 'stdDev',
  109: 'sum',
  110: 'var',
};

/**
 * Univer ne connaît pas la ligne de totaux : on la reconnaît à la ligne juste sous le tableau
 * qui contient des SUBTOTAL(…, Tableau[Colonne]) (créée par Grillon ou venant d'Excel).
 */
function withTotalsRow(sheet: Partial<IWorksheetData>, table: UniverTable): UniverTable {
  if (table.meta?.excel?.totalsRowCount) return table;
  const row = sheet.cellData?.[table.range.endRow + 1];
  if (!row) return table;
  const name = table.name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const re = new RegExp(`^=SUBTOTAL\\((\\d+),\\s*${name}\\[`, 'i');
  let found = false;
  const columns = table.columns.map((col, i) => {
    const cell = row[table.range.startColumn + i];
    const m = typeof cell?.f === 'string' ? re.exec(cell.f) : null;
    if (m) {
      found = true;
      const fn = TOTALS_FUNCTIONS[Number(m[1])];
      return fn ? { ...col, totalsRowFunction: fn } : col;
    }
    if (cell && cell.v !== undefined && cell.v !== null && cell.v !== '' && !cell.f) return { ...col, totalsRowLabel: String(cell.v) };
    return col;
  });
  if (!found) return table;
  return { ...table, columns, meta: { ...(table.meta ?? {}), excel: { ...(table.meta?.excel ?? {}), totalsRowCount: 1 } } };
}

function prepareTables(ws: ExcelJS.Worksheet, sheet: Partial<IWorksheetData>, tables: UniverTable[]): TableToWrite[] {
  return tables.map((original) => {
    const table = withTotalsRow(sheet, original);
    const names: string[] = [];
    const seen = new Set<string>();
    for (let c = table.range.startColumn; c <= table.range.endColumn; c++) {
      const i = c - table.range.startColumn;
      let name = table.options.showHeader !== false ? headerText(sheet.cellData?.[table.range.startRow]?.[c]).trim() : '';
      if (!name) name = table.columns[i]?.displayName || `Colonne${i + 1}`;
      let unique = name;
      for (let k = 2; seen.has(unique.toLowerCase()); k++) unique = `${name}${k}`;
      seen.add(unique.toLowerCase());
      names.push(unique);
      if (table.options.showHeader !== false) {
        const xc = ws.getRow(table.range.startRow + 1).getCell(c + 1);
        if (xc.value !== unique) xc.value = unique;
      }
    }
    return { table, headerNames: names };
  });
}

function buildWorkbook(wb: IWorkbookData): { xlWb: ExcelJS.Workbook; post: SheetPostProcess[] } {
  const xlWb = new ExcelJS.Workbook();
  xlWb.creator = 'Grillon';
  xlWb.created = new Date();
  const used = new Set<string>();
  const cf = readResource<UniverCfResource>(wb, CF_RESOURCE) ?? {};
  const dv = readResource<UniverDvResource>(wb, DV_RESOURCE) ?? {};
  const filters = readResource<UniverFilterResource>(wb, FILTER_RESOURCE) ?? {};
  const tables = readResource<UniverTableResource>(wb, TABLE_RESOURCE) ?? {};
  const drawings = readResource<DrawingResource>(wb, DRAWING_RESOURCE);
  const notes = readResource<UniverNoteResource>(wb, NOTE_RESOURCE) ?? {};
  const post: SheetPostProcess[] = [];
  for (const id of wb.sheetOrder) {
    const sheet = wb.sheets[id];
    if (!sheet) continue;
    const name = sanitizeSheetName(sheet.name ?? '', used);
    const filter = filters[id];
    const ws = writeSheet(wb, xlWb, sheet, name, filter?.cachedFilteredOut);
    const p: SheetPostProcess = { name, charts: chartsOfSheet(drawings, id), sparklines: sparklinesOfSheet(drawings, id), tables: [], stopIfTruePriorities: [], pivot: (sheet.custom as { grillon?: { pivot?: PivotDefinition } } | undefined)?.grillon?.pivot, protection: (sheet.custom as { grillon?: { protection?: SheetProtection } } | undefined)?.grillon?.protection };
    if (cf[id]?.length) {
      exportConditionalFormats(ws, cf[id]);
      cf[id].forEach((entry, i) => entry.stopIfTrue && p.stopIfTruePriorities.push(i + 1));
    }
    if (dv[id]?.length) exportDataValidations(ws, dv[id]);
    exportNotes(ws, notes[id]);
    if (filter?.ref) {
      ws.autoFilter = rangeToA1(filter.ref);
      if (filter.filterColumns?.length) p.autoFilter = filter;
    }
    if (tables[id]?.tables?.length) p.tables = prepareTables(ws, sheet, tables[id].tables);
    post.push(p);
  }

  const res = wb.resources?.find((r) => r.name === DEFINED_NAME_RESOURCE);
  if (res?.data) {
    try {
      const parsed = JSON.parse(res.data) as Record<string, StoredDefinedName> | Record<string, Record<string, StoredDefinedName>>;
      // Selon la version, Univer range les noms par classeur ({unitId: {id: nom}}) ou à plat.
      const flat: StoredDefinedName[] = [];
      for (const v of Object.values(parsed)) {
        if (v && typeof v === 'object' && 'name' in v && 'formulaOrRefString' in v) flat.push(v as StoredDefinedName);
        else if (v && typeof v === 'object') flat.push(...(Object.values(v) as StoredDefinedName[]));
      }
      for (const n of flat) {
        for (const part of n.formulaOrRefString.replace(/^=/, '').split(',')) {
          const ref = part.trim();
          if (!/!\$?[A-Z]+\$?\d+(:\$?[A-Z]+\$?\d+)?$/i.test(ref)) continue;
          try {
            xlWb.definedNames.add(ref, n.name);
          } catch {
            /* nom non représentable par ExcelJS : ignoré */
          }
        }
      }
    } catch {
      /* ressource illisible : ignorée */
    }
  }
  return { xlWb, post };
}

/** Classeur ExcelJS équivalent (sans les éléments écrits en post-traitement XML). */
export function univerToWorkbook(wb: IWorkbookData): ExcelJS.Workbook {
  return buildWorkbook(wb).xlWb;
}

/** Ajoute stopIfTrue="1" aux règles de MFC dont la priorité est listée. */
function addStopIfTrue(sheetXml: string, priorities: number[]): string {
  const set = new Set(priorities);
  return sheetXml.replace(/<cfRule\b[^>]*>/g, (tag) => {
    const p = /\bpriority="(\d+)"/.exec(tag);
    if (!p || !set.has(Number(p[1])) || /stopIfTrue=/.test(tag)) return tag;
    return tag.replace(/<cfRule\b/, '<cfRule stopIfTrue="1"');
  });
}

/** Écrit dans le XML ce qu'ExcelJS ne sait pas écrire (critères de filtre, tableaux, stopIfTrue). */
async function postProcess(wb: IWorkbookData, bytes: Uint8Array, post: SheetPostProcess[]): Promise<Uint8Array> {
  const needed = (p: SheetPostProcess) => p.autoFilter || p.tables.length || p.stopIfTruePriorities.length || p.charts.length || p.sparklines.length || !!p.pivot || !!p.protection;
  if (!post.some(needed)) return bytes;
  const zip = await loadZip(bytes);
  const sheets = await listSheets(zip);
  let tableId = 0;
  const counters = { drawing: 0, chart: 0 };
  const usedTableNames = new Set<string>();
  const pivots: PivotToWrite[] = [];
  for (let i = 0; i < sheets.length && i < post.length; i++) {
    const p = post[i];
    if (!needed(p)) continue;
    const file = zip.file(sheets[i].path);
    if (!file) continue;
    let xml = await file.async('string');
    if (p.protection) xml = insertWorksheetElement(xml, 'sheetProtection', sheetProtectionXml(p.protection));
    if (p.autoFilter) xml = insertWorksheetElement(xml, 'autoFilter', buildAutoFilterXml(p.autoFilter));
    if (p.stopIfTruePriorities.length) xml = addStopIfTrue(xml, p.stopIfTruePriorities);
    if (p.tables.length) xml = await writeTables(zip, sheets[i].path, xml, p.tables, () => ++tableId, usedTableNames);
    if (p.charts.length) xml = await writeSheetCharts(zip, wb, p.name, sheets[i].path, xml, p.charts, counters);
    if (p.sparklines.length) xml = writeSparklines(xml, p.sparklines);
    zip.file(sheets[i].path, xml);
    if (p.pivot) pivots.push({ sheetPath: sheets[i].path, def: p.pivot });
  }
  // Les tableaux croisés touchent aussi le classeur (caches) : écrits après les feuilles.
  if (pivots.length) await writePivotTables(zip, wb, pivots);
  return zipToBytes(zip);
}

export async function exportXlsx(wb: IWorkbookData): Promise<Uint8Array> {
  const { xlWb, post } = buildWorkbook(wb);
  const buf = await xlWb.xlsx.writeBuffer();
  return postProcess(wb, new Uint8Array(buf as ArrayBuffer), post);
}
