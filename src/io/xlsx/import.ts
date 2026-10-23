/**
 * Import .xlsx → IWorkbookData (modèle Univer).
 *
 * Lecture via ExcelJS (MIT). Ne dépend ni du DOM ni d'Univer à l'exécution :
 * fonctionne dans Node (tests), dans le navigateur et dans un Web Worker.
 */
import ExcelJS from 'exceljs';
import {
  CellValueType,
  DEFAULT_COLUMN_COUNT,
  DEFAULT_ROW_COUNT,
  DEFINED_NAME_RESOURCE,
  GRILLON_CUSTOM_KEY,
  LOCALE_FR,
  type ICellData,
  type IColumnData,
  type IRange,
  type IRowData,
  type IStyleData,
  type IWorkbookData,
  type IWorksheetData,
} from '../../model/univer-types';
import { FILTER_RESOURCE, readAutoFilter, type UniverAutoFilter, type UniverFilterResource } from './auto-filter';
import { addSlicerToResource, chartsToDrawingResource, DRAWING_RESOURCE, readSheetCharts, type DrawingResource, type SheetChart } from './charts';
import { readSlicers, type ReadSlicersResult } from './slicers';
import { parseThemeColors } from './colors';
import { parseTheme, sameTheme, DEFAULT_THEME, THEME_CUSTOM_KEY } from '../../data/theme';
import { CF_RESOURCE, importConditionalFormats, type UniverCfResource } from './conditional-formatting';
import { DV_RESOURCE, readDataValidations, type UniverDvResource, type UniverDvRule } from './data-validation';
import { importNotes, NOTE_RESOURCE, readNotes, type SheetNotes, type UniverNoteResource } from './notes';
import { normalizePackage, stripForExcelJs } from './normalize';
import { listSheets, loadZip, removeWorksheetElement, zipToBytes } from './ooxml';
import { readTables, TABLE_RESOURCE, type UniverTable, type UniverTableResource } from './tables';
import { richTextToDocument } from './rich-text';
import { styleToUniver } from './styles';
import { dateToSerial, EXCEL_DEFAULT_COL_WIDTH, EXCEL_DEFAULT_ROW_HEIGHT_PT, excelWidthToPx, fromA1, ptToPx, rangeFromA1, rangeToA1 } from './units';
import { fromExcelPageSetup, type ExcelPageSetupLike, type PageSetup } from '../../data/page-setup';
import { readBreaks, withBreaks } from '../../data/page-breaks';
import { addSparklinesToResource, readSparklines, type SheetSparkline } from './sparklines';
import type { PivotDefinition } from '../../data/pivot';
import { readPivotTables } from './pivot-tables';
import { readSheetProtection, type SheetProtection } from '../../data/sheet-protection';
import { cellsToRects } from '../../data/cell-locks';

/** Police par défaut d'Excel (les cellules sans police explicite l'utilisent). */
export const EXCEL_DEFAULT_FONT = { ff: 'Calibri', fs: 11 } as const;

export interface ImportOptions {
  /** Nom du classeur (sinon « Classeur »). */
  name?: string;
  /** Identifiant du classeur (sinon généré). */
  id?: string;
}

/** Registre des styles : dédoublonne les IStyleData et renvoie un identifiant. */
class StyleRegistry {
  readonly styles: Record<string, IStyleData> = {};
  private readonly byKey = new Map<string, string>();
  private readonly cache = new WeakMap<object, string | null>();

  constructor(readonly theme: string[]) {}

  /** ExcelJS partage souvent le même objet style entre cellules : on met le résultat en cache. */
  idFor(style: Partial<ExcelJS.Style> | undefined): string | null {
    if (!style || typeof style !== 'object') return null;
    const cached = this.cache.get(style);
    if (cached !== undefined) return cached;
    const u = styleToUniver(style, this.theme);
    let id: string | null = null;
    if (u) {
      const key = JSON.stringify(u);
      id = this.byKey.get(key) ?? null;
      if (!id) {
        id = `s${this.byKey.size + 1}`;
        this.byKey.set(key, id);
        this.styles[id] = u;
      }
    }
    this.cache.set(style, id);
    return id;
  }
}

function newId(prefix: string): string {
  return `${prefix}-${Math.random().toString(36).slice(2, 10)}`;
}

/** Valeur ExcelJS → (v, t) Univer. */
function convertValue(value: ExcelJS.CellValue): Pick<ICellData, 'v' | 't'> | null {
  if (value === null || value === undefined) return null;
  if (typeof value === 'number') return { v: value, t: CellValueType.NUMBER };
  if (typeof value === 'string') return { v: value, t: CellValueType.STRING };
  if (typeof value === 'boolean') return { v: value ? 1 : 0, t: CellValueType.BOOLEAN };
  if (value instanceof Date) return { v: dateToSerial(value), t: CellValueType.NUMBER };
  if (typeof value === 'object') {
    if ('richText' in value && Array.isArray(value.richText)) {
      return { v: value.richText.map((r) => r.text).join(''), t: CellValueType.STRING };
    }
    if ('error' in value && value.error) return { v: String(value.error), t: CellValueType.STRING };
    if ('text' in value && 'hyperlink' in value) {
      const text = value.text as unknown;
      if (typeof text === 'string') return { v: text, t: CellValueType.STRING };
      // Nombre avec un lien (numéro de dossier pointant vers une page…).
      if (typeof text === 'number') return { v: text, t: CellValueType.NUMBER };
      if (text && typeof text === 'object' && 'richText' in text) {
        return { v: (text as ExcelJS.CellRichTextValue).richText.map((r) => r.text).join(''), t: CellValueType.STRING };
      }
    }
  }
  return null;
}

function convertCell(cell: ExcelJS.Cell, registry: StyleRegistry): ICellData | null {
  const out: ICellData = {};
  const value = cell.value;
  const isFormula = cell.type === ExcelJS.ValueType.Formula;
  if (isFormula) {
    const f = cell.formula;
    if (f) out.f = `=${f}`;
    const result = (value as ExcelJS.CellFormulaValue).result;
    const conv = convertValue(result as ExcelJS.CellValue);
    if (conv) Object.assign(out, conv);
  } else {
    const conv = convertValue(value);
    if (conv) Object.assign(out, conv);
    // Texte enrichi : on garde les styles de chaque morceau dans un document Univer.
    const rich = value && typeof value === 'object' && 'richText' in value ? value.richText : undefined;
    if (rich?.length && rich.some((r) => r.font && Object.keys(r.font).length)) out.p = richTextToDocument(rich, registry.theme);
  }
  const hyperlink = cell.hyperlink;
  if (hyperlink) out.custom = { link: hyperlink };
  const s = registry.idFor(cell.style);
  if (s) out.s = s;
  if (out.v === undefined && out.f === undefined && out.s === undefined) return null;
  return out;
}

/** Cellules couvertes par une fusion, sauf la cellule en haut à gauche (« r,c »). */
function mergedFollowers(merges: IRange[]): Set<string> {
  const out = new Set<string>();
  for (const m of merges) {
    if ((m.endRow - m.startRow + 1) * (m.endColumn - m.startColumn + 1) > 100_000) continue;
    for (let r = m.startRow; r <= m.endRow; r++) for (let c = m.startColumn; c <= m.endColumn; c++) if (r !== m.startRow || c !== m.startColumn) out.add(`${r},${c}`);
  }
  return out;
}

function convertWorksheet(ws: ExcelJS.Worksheet, registry: StyleRegistry, sheetMerges?: IRange[]): IWorksheetData {
  // Fusions lues par Grillon (ExcelJS met des dizaines de secondes sur des milliers de fusions).
  const merges = sheetMerges ?? ((ws.model as unknown as { merges?: string[] }).merges ?? []).map(rangeFromA1);
  const followers = mergedFollowers(merges);
  const id = newId('sheet');
  const cellData: Record<number, Record<number, ICellData>> = {};
  const rowData: Record<number, Partial<IRowData>> = {};
  const columnData: Record<number, Partial<IColumnData>> = {};
  let maxRow = 0;
  let maxCol = 0;
  // Cellules déverrouillées (Format de cellule → Protection → Verrouillée décochée).
  const unlocked: [number, number][] = [];

  const props = ws.properties ?? ({} as ExcelJS.WorksheetProperties);
  const defaultColWidthPx = excelWidthToPx(props.defaultColWidth ?? EXCEL_DEFAULT_COL_WIDTH);
  const defaultRowHeightPx = ptToPx(props.defaultRowHeight ?? EXCEL_DEFAULT_ROW_HEIGHT_PT);

  ws.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    const r = rowNumber - 1;
    const rd: Partial<IRowData> = {};
    if (row.height && Math.abs(ptToPx(row.height) - defaultRowHeightPx) > 0) rd.h = ptToPx(row.height);
    if (row.hidden) rd.hd = 1;
    if (row.outlineLevel) rd.custom = { [GRILLON_CUSTOM_KEY]: { outline: row.outlineLevel } };
    // Style de ligne entière (rare) : seulement s'il est défini sur la ligne elle-même.
    const rowStyle = (row as unknown as { _style?: Partial<ExcelJS.Style> })._style ?? undefined;
    if (rowStyle && Object.keys(rowStyle).length) {
      const s = registry.idFor(rowStyle);
      if (s) rd.s = s;
    }
    if (Object.keys(rd).length) rowData[r] = rd;

    const cells = (row as unknown as { _cells: (ExcelJS.Cell | undefined)[] })._cells ?? [];
    for (let c = 0; c < cells.length; c++) {
      const cell = cells[c];
      if (!cell) continue;
      if (cell.style?.protection?.locked === false) unlocked.push([r, c]);
      // Les cellules fusionnées « esclaves » sont ignorées : seule la cellule maître porte la valeur.
      if (followers.has(`${r},${c}`)) {
        const s = registry.idFor(cell.style);
        if (!s) continue;
        (cellData[r] ??= {})[c] = { s };
        continue;
      }
      const data = convertCell(cell, registry);
      if (!data) continue;
      (cellData[r] ??= {})[c] = data;
      if (r > maxRow) maxRow = r;
      if (c > maxCol) maxCol = c;
    }
  });

  const columns = (ws as unknown as { _columns?: ExcelJS.Column[] })._columns ?? [];
  columns.forEach((col, i) => {
    if (!col) return;
    const cd: Partial<IColumnData> = {};
    const w = col.width;
    if (w !== undefined && w !== null) {
      const px = excelWidthToPx(w);
      if (px !== defaultColWidthPx) cd.w = px;
    }
    if (col.hidden) cd.hd = 1;
    if (col.outlineLevel) cd.custom = { [GRILLON_CUSTOM_KEY]: { outline: col.outlineLevel } };
    const colStyle = (col as unknown as { style?: Partial<ExcelJS.Style> }).style;
    if (colStyle && Object.keys(colStyle).length) {
      const s = registry.idFor(colStyle);
      if (s) cd.s = s;
    }
    if (Object.keys(cd).length) columnData[i] = cd;
  });

  for (const m of merges) {
    maxRow = Math.max(maxRow, m.endRow);
    maxCol = Math.max(maxCol, m.endColumn);
  }

  const view = ws.views?.[0] as Partial<ExcelJS.WorksheetViewFrozen & ExcelJS.WorksheetViewCommon> | undefined;
  let freeze = { xSplit: 0, ySplit: 0, startRow: -1, startColumn: -1 };
  if (view && view.state === 'frozen' && ((view.xSplit ?? 0) > 0 || (view.ySplit ?? 0) > 0)) {
    const xSplit = view.xSplit ?? 0;
    const ySplit = view.ySplit ?? 0;
    freeze = { xSplit, ySplit, startRow: ySplit, startColumn: xSplit };
  }

  const tabArgb = (props.tabColor as { argb?: string } | undefined)?.argb;
  const hf = ws.headerFooter as { oddHeader?: string; oddFooter?: string } | undefined;
  const page = fromExcelPageSetup(ws.pageSetup as unknown as ExcelPageSetupLike, hf?.oddHeader, hf?.oddFooter);

  return {
    id,
    name: ws.name,
    tabColor: tabArgb ? `#${tabArgb.slice(-6)}` : '',
    hidden: ws.state === 'hidden' || ws.state === 'veryHidden' ? 1 : 0,
    freeze,
    rowCount: Math.max(DEFAULT_ROW_COUNT, maxRow + 1 + 100),
    columnCount: Math.max(DEFAULT_COLUMN_COUNT, maxCol + 1 + 5),
    zoomRatio: view?.zoomScale ? view.zoomScale / 100 : 1,
    scrollTop: 0,
    scrollLeft: 0,
    defaultColumnWidth: defaultColWidthPx,
    defaultRowHeight: defaultRowHeightPx,
    mergeData: merges,
    cellData,
    rowData,
    columnData,
    rowHeader: { width: 46, hidden: 0 },
    columnHeader: { height: 20, hidden: 0 },
    showGridlines: view?.showGridLines === false ? 0 : 1,
    rightToLeft: view?.rightToLeft ? 1 : 0,
    ...(page || unlocked.length ? { custom: { [GRILLON_CUSTOM_KEY]: { ...(page ? { page } : {}), ...(unlocked.length ? { unlocked: cellsToRects(unlocked) } : {}) } } } : {}),
  } as IWorksheetData;
}

interface DefinedNameModel {
  name: string;
  ranges: string[];
  localSheetId?: number;
}

/** Éléments lus directement dans le XML d'une feuille (ce qu'ExcelJS ignore ou lit mal). */
export interface SheetExtras {
  dataValidations: UniverDvRule[];
  autoFilter: UniverAutoFilter | null;
  tables: UniverTable[];
  cfStopIfTrue: Set<number>;
  charts: SheetChart[];
  sparklines: SheetSparkline[];
  pivots: PivotDefinition[];
  protection: SheetProtection | null;
  rowBreaks: number[];
  colBreaks: number[];
  /** Fusions de cellules (retirées du XML donné à ExcelJS). */
  merges?: IRange[];
  /** Notes (commentaires), lues par Grillon. */
  notes?: SheetNotes;
  /** Feuille graphique (sans cellules). */
  chartSheet?: boolean;
}

function readMerges(sheetXml: string): IRange[] {
  const out: IRange[] = [];
  for (const m of sheetXml.matchAll(/<(?:\w+:)?mergeCell\b[^>]*\bref="([^"]+)"/g)) {
    try {
      out.push(rangeFromA1(m[1]));
    } catch {
      /* fusion illisible : ignorée */
    }
  }
  return out;
}

function readCfStopIfTrue(sheetXml: string): Set<number> {
  const out = new Set<number>();
  for (const m of sheetXml.matchAll(/<(?:\w+:)?cfRule\b[^>]*>/g)) {
    if (!/\bstopIfTrue="(1|true)"/.test(m[0])) continue;
    const p = /\bpriority="(\d+)"/.exec(m[0]);
    if (p) out.add(Number(p[1]));
  }
  return out;
}

/**
 * Lit les extras de chaque feuille et retire du XML les validations de données avant de
 * passer le fichier à ExcelJS (qui déplierait chaque plage cellule par cellule).
 */
async function preprocess(data: ArrayBuffer | Uint8Array): Promise<{ bytes: ArrayBuffer | Uint8Array; extras: Map<string, SheetExtras>; slicers: ReadSlicersResult }> {
  const extras = new Map<string, SheetExtras>();
  const zip = await loadZip(data);
  let modified = await normalizePackage(zip);
  const pivotsByPath = await readPivotTables(zip);
  const slicers = await readSlicers(zip);
  for (const sheet of await listSheets(zip)) {
    const file = zip.file(sheet.path);
    if (!file) continue;
    const xml = await file.async('string');
    extras.set(sheet.name, {
      dataValidations: readDataValidations(xml, ''),
      autoFilter: readAutoFilter(xml),
      tables: await readTables(zip, sheet.path),
      cfStopIfTrue: readCfStopIfTrue(xml),
      charts: await readSheetCharts(zip, sheet.path),
      sparklines: readSparklines(xml),
      pivots: pivotsByPath.get(sheet.path) ?? [],
      protection: readSheetProtection(xml),
      rowBreaks: readBreaks(xml, 'rowBreaks'),
      merges: readMerges(xml),
      notes: await readNotes(zip, sheet.path),
      chartSheet: /<(?:\w+:)?chartsheet\b/.test(xml.slice(0, 500)),
      colBreaks: readBreaks(xml, 'colBreaks'),
    });
    const stripped = removeWorksheetElement(xml, 'dataValidations');
    if (stripped.removed) {
      zip.file(sheet.path, stripped.xml);
      modified = true;
    }
  }
  if (await stripForExcelJs(zip, (await listSheets(zip)).map((s) => s.path))) modified = true;
  return { bytes: modified ? await zipToBytes(zip) : data, extras, slicers };
}

export async function importXlsx(data: ArrayBuffer | Uint8Array, options: ImportOptions = {}): Promise<IWorkbookData> {
  const { bytes, extras, slicers } = await preprocess(data);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(bytes as ArrayBuffer);
  return workbookToUniver(wb, options, extras, slicers);
}

export function workbookToUniver(
  wb: ExcelJS.Workbook,
  options: ImportOptions = {},
  extras = new Map<string, SheetExtras>(),
  slicers: ReadSlicersResult = { slicers: [], syntheticTables: [] },
): IWorkbookData {
  const themes = (wb as unknown as { _themes?: Record<string, string> })._themes;
  const theme = parseThemeColors(themes?.theme1);
  const registry = new StyleRegistry(theme);
  // Thème du fichier : gardé (s'il n'est pas celui d'Office par défaut) pour pouvoir le changer et le réécrire.
  const fileTheme = parseTheme(themes?.theme1, theme);
  const custom = fileTheme && !sameTheme(fileTheme, DEFAULT_THEME) ? { grillon: { [THEME_CUSTOM_KEY]: fileTheme } } : undefined;

  const sheets: Record<string, Partial<IWorksheetData>> = {};
  const sheetOrder: string[] = [];
  const cf: UniverCfResource = {};
  const dv: UniverDvResource = {};
  const filters: UniverFilterResource = {};
  const tables: UniverTableResource = {};
  const drawings: DrawingResource = {};
  const notes: UniverNoteResource = {};
  const unitId = options.id ?? newId('wb');
  wb.eachSheet((ws) => {
    const sheet = convertWorksheet(ws, registry, extras.get(ws.name)?.merges);
    sheets[sheet.id] = sheet;
    sheetOrder.push(sheet.id);
    const extra = extras.get(ws.name);
    const sheetNotes = extra?.notes ?? importNotes(ws);
    if (Object.keys(sheetNotes).length) notes[sheet.id] = sheetNotes;
    const rules = importConditionalFormats(ws, sheet.id, theme, extra?.cfStopIfTrue);
    if (rules.length) cf[sheet.id] = rules;
    if (extra?.dataValidations.length) {
      dv[sheet.id] = extra.dataValidations.map((r) => ({ ...r, ranges: r.ranges.map((rg) => ({ ...rg, sheetId: sheet.id })) }));
    }
    if (extra?.autoFilter) {
      filters[sheet.id] = extra.autoFilter;
      // Les lignes masquées par le filtre sont recalculées par Univer : on ne les garde pas
      // comme lignes masquées « à la main », sinon elles le resteraient après « Effacer le filtre ».
      if (extra.autoFilter.filterColumns?.length) {
        const { startRow, endRow } = extra.autoFilter.ref;
        for (let r = startRow + 1; r <= endRow; r++) {
          const rd = sheet.rowData?.[r];
          if (rd?.hd) {
            delete rd.hd;
            if (!Object.keys(rd).length) delete sheet.rowData[r];
          }
        }
      }
    }
    // Tableaux créés par Grillon pour porter un segment : ce n'étaient que des plages.
    const synthetic = new Set(slicers.syntheticTables.filter((t) => t.sheet === ws.name).map((t) => t.ref));
    const sheetTables = extra?.tables.filter((t) => !synthetic.has(rangeToA1(t.range))) ?? [];
    if (sheetTables.length) tables[sheet.id] = { tables: sheetTables, tableFilteredOutRows: [] };
    if (extra?.charts.length) drawings[sheet.id] = chartsToDrawingResource(unitId, sheet.id, sheet, extra.charts);
    if (extra?.rowBreaks.length || extra?.colBreaks.length) {
      const custom = (sheet.custom ?? {}) as { grillon?: { page?: PageSetup } & Record<string, unknown> };
      const page = withBreaks(custom.grillon?.page ?? {}, { rowBreaks: extra.rowBreaks, colBreaks: extra.colBreaks });
      sheet.custom = { ...custom, grillon: { ...(custom.grillon ?? {}), page } } as never;
    }
    if (extra?.protection) {
      const custom = (sheet.custom ?? {}) as { grillon?: Record<string, unknown> };
      sheet.custom = { ...custom, grillon: { ...(custom.grillon ?? {}), protection: extra.protection } } as never;
    }
    if (extra?.pivots.length) {
      // Grillon gère un tableau croisé par feuille : le premier est relu, les autres restent en valeurs.
      const custom = (sheet.custom ?? {}) as { grillon?: Record<string, unknown> };
      sheet.custom = { ...custom, grillon: { ...(custom.grillon ?? {}), pivot: extra.pivots[0] } } as never;
    }
    if (extra?.sparklines.length) drawings[sheet.id] = addSparklinesToResource(unitId, sheet.id, sheet, extra.sparklines, drawings[sheet.id]);
    for (const s of slicers.slicers.filter((x) => x.sheet === ws.name)) drawings[sheet.id] = addSlicerToResource(drawings[sheet.id], unitId, sheet.id, sheet, s);
  });

  // Feuilles graphiques (« Déplacer le graphique → nouvelle feuille ») : ExcelJS les ignore ; elles deviennent des
  // feuilles qui portent leur graphique, à leur place dans l'ordre des onglets.
  const known = new Set(sheetOrder.map((id) => sheets[id]!.name));
  for (const [name, extra] of extras) {
    if (known.has(name) || !extra.chartSheet) continue;
    const id = newId('sheet');
    const sheet: Partial<IWorksheetData> = { id, name, cellData: {}, rowCount: DEFAULT_ROW_COUNT, columnCount: DEFAULT_COLUMN_COUNT };
    sheets[id] = sheet;
    sheetOrder.push(id);
    if (extra.charts.length) drawings[id] = chartsToDrawingResource(unitId, id, sheet, extra.charts);
  }
  const position = new Map([...extras.keys()].map((name, i) => [name, i]));
  sheetOrder.sort((a, b) => (position.get(sheets[a]!.name ?? '') ?? 1e9) - (position.get(sheets[b]!.name ?? '') ?? 1e9));

  if (!sheetOrder.length) {
    const id = newId('sheet');
    sheets[id] = { id, name: 'Feuil1', cellData: {} };
    sheetOrder.push(id);
  }

  const resources: IWorkbookData['resources'] = [];
  const names = (wb.definedNames.model ?? []) as DefinedNameModel[];
  if (names.length) {
    const map: Record<string, { id: string; name: string; formulaOrRefString: string; localSheetId?: string }> = {};
    for (const n of names) {
      // Les noms internes d'Excel (_xlnm.Print_Area, filtres…) ne sont pas des noms utilisateur.
      if (!n.name || n.name.startsWith('_xlnm.')) continue;
      // Noms des caches de segments (« Segment_… » = #N/A) : internes à Excel.
      if (!n.ranges.length || n.ranges.every((r) => /^#N\/A$/i.test(r.replace(/^=/, '')))) continue;
      const id = newId('name');
      const refs = n.ranges.map(normalizeSheetRef);
      map[id] = { id, name: n.name, formulaOrRefString: refs.join(',') };
    }
    if (Object.keys(map).length) resources.push({ name: DEFINED_NAME_RESOURCE, data: JSON.stringify(map) });
  }

  if (Object.keys(cf).length) resources.push({ name: CF_RESOURCE, data: JSON.stringify(cf) });
  if (Object.keys(dv).length) resources.push({ name: DV_RESOURCE, data: JSON.stringify(dv) });
  if (Object.keys(filters).length) resources.push({ name: FILTER_RESOURCE, data: JSON.stringify(filters) });
  if (Object.keys(tables).length) resources.push({ name: TABLE_RESOURCE, data: JSON.stringify(tables) });
  if (Object.keys(notes).length) resources.push({ name: NOTE_RESOURCE, data: JSON.stringify(notes) });
  if (Object.keys(drawings).length) resources.push({ name: DRAWING_RESOURCE, data: JSON.stringify(drawings) });

  return {
    id: unitId,
    name: options.name ?? 'Classeur',
    appVersion: '1.0.0',
    locale: LOCALE_FR as IWorkbookData['locale'],
    styles: registry.styles,
    defaultStyle: { ...EXCEL_DEFAULT_FONT },
    sheetOrder,
    sheets,
    resources,
    ...(custom ? { custom } : {}),
  };
}

/** ExcelJS renvoie parfois « 'Feuil 1'!$A$1:$B$2 » ou « Feuil1!A1 » : on garde ce format Excel tel quel. */
function normalizeSheetRef(ref: string): string {
  const bang = ref.lastIndexOf('!');
  if (bang < 0) return ref;
  const sheet = ref.slice(0, bang);
  const cells = ref.slice(bang + 1);
  // Validation légère : on s'assure que la partie cellule est bien une référence A1.
  try {
    cells.split(':').forEach((p) => fromA1(p));
  } catch {
    return ref;
  }
  return `${sheet}!${cells}`;
}
