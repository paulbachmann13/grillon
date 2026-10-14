/**
 * Export IWorkbookData (modèle Univer) → .xlsx via ExcelJS.
 */
import ExcelJS from 'exceljs';
import {
  CellValueType,
  DEFINED_NAME_RESOURCE,
  type ICellData,
  type IStyleData,
  type IWorkbookData,
  type IWorksheetData,
} from '../../model/univer-types';
import { cssToArgb } from './colors';
import { styleFromUniver } from './styles';
import { EXCEL_DEFAULT_COL_WIDTH, EXCEL_DEFAULT_ROW_HEIGHT_PT, pxToExcelWidth, pxToPt, rangeToA1 } from './units';

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

/** Univer garde parfois du texte enrichi (p) sans v : on en tire le texte brut. */
function plainTextFromDoc(cell: ICellData): string | undefined {
  const body = cell.p?.body?.dataStream;
  if (!body) return undefined;
  return body.replace(/\r\n$/, '').replace(/[\r\n]+$/, '').replace(/\r/g, '\n');
}

function writeSheet(wb: IWorkbookData, xlWb: ExcelJS.Workbook, sheet: Partial<IWorksheetData>, name: string): void {
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

  // Colonnes : largeur, masquage, style de colonne.
  for (const [k, cd] of Object.entries(sheet.columnData ?? {})) {
    if (!cd) continue;
    const col = ws.getColumn(Number(k) + 1);
    if (cd.w !== undefined) col.width = pxToExcelWidth(cd.w);
    if (cd.hd) col.hidden = true;
    const st = styleFromUniver(resolveStyle(wb, cd.s));
    if (st) col.style = st as ExcelJS.Style;
  }

  // Lignes : hauteur, masquage, style.
  for (const [k, rd] of Object.entries(sheet.rowData ?? {})) {
    if (!rd) continue;
    const row = ws.getRow(Number(k) + 1);
    if (rd.h !== undefined && rd.h !== sheet.defaultRowHeight) row.height = pxToPt(rd.h);
    if (rd.hd) row.hidden = true;
    const st = styleFromUniver(resolveStyle(wb, rd.s));
    if (st) (row as unknown as { style: Partial<ExcelJS.Style> }).style = st;
  }

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
      if (v === null && cell.p) v = plainTextFromDoc(cell) ?? null;
      const link = typeof cell.custom?.link === 'string' ? (cell.custom.link as string) : undefined;
      if (f) {
        xc.value = { formula: f, result: v ?? undefined } as ExcelJS.CellFormulaValue;
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
}

interface StoredDefinedName {
  id: string;
  name: string;
  formulaOrRefString: string;
}

export function univerToWorkbook(wb: IWorkbookData): ExcelJS.Workbook {
  const xlWb = new ExcelJS.Workbook();
  xlWb.creator = 'Grillon';
  xlWb.created = new Date();
  const used = new Set<string>();
  const renamed = new Map<string, string>();
  for (const id of wb.sheetOrder) {
    const sheet = wb.sheets[id];
    if (!sheet) continue;
    const name = sanitizeSheetName(sheet.name ?? '', used);
    if (sheet.name && name !== sheet.name) renamed.set(sheet.name, name);
    writeSheet(wb, xlWb, sheet, name);
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
  return xlWb;
}

export async function exportXlsx(wb: IWorkbookData): Promise<Uint8Array> {
  const xlWb = univerToWorkbook(wb);
  const buf = await xlWb.xlsx.writeBuffer();
  return new Uint8Array(buf as ArrayBuffer);
}
