/**
 * Mini-graphiques (sparklines) ⇄ extension Excel 2010 (x14:sparklineGroups dans <extLst>).
 * Dans Grillon, un mini-graphique est un « DOM flottant » ancré sur une cellule
 * (ressource SHEET_DRAWING_PLUGIN, composant GrillonSparkline).
 */
import type { SparklineSpec, SparklineType } from '../../charts/sparkline';
import { SPARK_COLOR, SPARK_HIGHLIGHT } from '../../charts/sparkline';
import type { IWorksheetData } from '../../model/univer-types';
import { type CellAnchor, DRAWING_RESOURCE, type DrawingResource } from './charts';
import { escapeXml, extractAll, getWorksheetElement, insertWorksheetElement, parseXml } from './ooxml';

export const SPARKLINE_COMPONENT_KEY = 'GrillonSparkline';
const DRAWING_DOM_TYPE = 8;
const EXT_URI = '{05C60535-1F16-4fd2-B633-F4F36F0B64E0}';

export interface SheetSparkline {
  spec: SparklineSpec;
  row: number;
  col: number;
}

type Xml = Record<string, unknown>;
const arr = <T>(v: T | T[] | undefined): T[] => (v === undefined ? [] : Array.isArray(v) ? v : [v]);

const rgb = (color: string | undefined, fallback: string) => `FF${(color ?? fallback).replace('#', '').toUpperCase()}`;
const fromRgb = (node: unknown): string | undefined => {
  const v = (node as { rgb?: string } | undefined)?.rgb;
  return v ? `#${v.slice(-6).toUpperCase()}` : undefined;
};

const TYPE_TO_XML: Record<SparklineType, string | undefined> = { line: undefined, column: 'column', winloss: 'stacked' };

function groupKey(s: SparklineSpec): string {
  return JSON.stringify({ ...s, dataRef: undefined });
}

function colRowToA1(row: number, col: number): string {
  let name = '';
  for (let n = col + 1; n > 0; n = Math.floor((n - 1) / 26)) name = String.fromCharCode(65 + ((n - 1) % 26)) + name;
  return `${name}${row + 1}`;
}

function a1ToColRow(a1: string): { row: number; col: number } | null {
  const m = /^\$?([A-Z]{1,3})\$?(\d+)$/i.exec(a1.trim());
  if (!m) return null;
  let col = 0;
  for (const ch of m[1]!.toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { row: Number(m[2]) - 1, col: col - 1 };
}

/** Écrit les mini-graphiques dans l'<extLst> d'une feuille. */
export function writeSparklines(sheetXml: string, items: SheetSparkline[]): string {
  if (!items.length) return sheetXml;
  const groups = new Map<string, SheetSparkline[]>();
  for (const it of items) {
    const k = groupKey(it.spec);
    groups.set(k, [...(groups.get(k) ?? []), it]);
  }
  const b = (v: boolean | undefined) => (v ? ' 1' : '');
  const groupXml = [...groups.values()]
    .map((list) => {
      const s = list[0]!.spec;
      const type = TYPE_TO_XML[s.type];
      const flag = (name: string, on: boolean | undefined) => (on ? ` ${name}="1"` : '');
      const attrs =
        (type ? ` type="${type}"` : '') +
        (s.lineWidth ? ` lineWeight="${s.lineWidth}"` : '') +
        ` displayEmptyCellsAs="${s.emptyCells === 'zero' ? 'zero' : s.emptyCells === 'connect' ? 'span' : 'gap'}"` +
        flag('markers', s.markers) +
        flag('high', s.high) +
        flag('low', s.low) +
        flag('first', s.first) +
        flag('last', s.last) +
        flag('negative', s.negative) +
        (s.min !== undefined ? ` minAxisType="custom" manualMin="${s.min}"` : '') +
        (s.max !== undefined ? ` maxAxisType="custom" manualMax="${s.max}"` : '');
      const hl = rgb(s.highlight, SPARK_HIGHLIGHT);
      const colors =
        `<x14:colorSeries rgb="${rgb(s.color, SPARK_COLOR)}"/><x14:colorNegative rgb="${hl}"/><x14:colorAxis rgb="FF000000"/>` +
        `<x14:colorMarkers rgb="${hl}"/><x14:colorFirst rgb="${hl}"/><x14:colorLast rgb="${hl}"/><x14:colorHigh rgb="${hl}"/><x14:colorLow rgb="${hl}"/>`;
      const lines = list.map((it) => `<x14:sparkline><xm:f>${escapeXml(it.spec.dataRef)}</xm:f><xm:sqref>${colRowToA1(it.row, it.col)}</xm:sqref></x14:sparkline>`).join('');
      void b;
      return `<x14:sparklineGroup${attrs}>${colors}<x14:sparklines>${lines}</x14:sparklines></x14:sparklineGroup>`;
    })
    .join('');
  const ext = `<ext uri="${EXT_URI}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">${groupXml}</x14:sparklineGroups></ext>`;
  const existing = getWorksheetElement(sheetXml, 'extLst');
  if (existing) return insertWorksheetElement(sheetXml, 'extLst', existing.replace('</extLst>', `${ext}</extLst>`));
  return insertWorksheetElement(sheetXml, 'extLst', `<extLst>${ext}</extLst>`);
}

/** Lit les mini-graphiques d'une feuille. */
export function readSparklines(sheetXml: string): SheetSparkline[] {
  const out: SheetSparkline[] = [];
  for (const g of extractAll(sheetXml, 'sparklineGroup')) {
    const doc = parseXml<{ sparklineGroup?: Xml }>(g.xml).sparklineGroup;
    if (!doc) continue;
    const type = String(doc.type ?? 'line');
    const spec: Omit<SparklineSpec, 'dataRef'> = { type: type === 'column' ? 'column' : type === 'stacked' ? 'winloss' : 'line' };
    const color = fromRgb(doc.colorSeries);
    if (color && color !== SPARK_COLOR) spec.color = color;
    const highlight = fromRgb(doc.colorHigh) ?? fromRgb(doc.colorNegative);
    if (highlight && highlight !== SPARK_HIGHLIGHT) spec.highlight = highlight;
    const flags = ['markers', 'high', 'low', 'first', 'last', 'negative'] as const;
    for (const f of flags) if (String(doc[f]) === '1' || String(doc[f]) === 'true') spec[f] = true;
    const weight = Number(doc.lineWeight);
    if (Number.isFinite(weight) && weight > 0) spec.lineWidth = weight;
    const blanks = String(doc.displayEmptyCellsAs ?? 'gap');
    if (blanks === 'zero') spec.emptyCells = 'zero';
    else if (blanks === 'span') spec.emptyCells = 'connect';
    if (doc.minAxisType === 'custom' && Number.isFinite(Number(doc.manualMin))) spec.min = Number(doc.manualMin);
    if (doc.maxAxisType === 'custom' && Number.isFinite(Number(doc.manualMax))) spec.max = Number(doc.manualMax);
    for (const sp of arr(((doc.sparklines as Xml | undefined)?.sparkline) as Xml | Xml[] | undefined)) {
      const f = String(sp.f ?? '').replace(/^=/, '');
      const cell = a1ToColRow(String(sp.sqref ?? ''));
      if (f && cell) out.push({ spec: { ...spec, dataRef: f }, row: cell.row, col: cell.col });
    }
  }
  return out;
}

/** Mini-graphiques d'une feuille (depuis la ressource de dessins Univer). */
export function sparklinesOfSheet(resource: DrawingResource | undefined, sheetId: string): SheetSparkline[] {
  const sheet = resource?.[sheetId];
  if (!sheet) return [];
  const out: SheetSparkline[] = [];
  for (const id of sheet.order ?? Object.keys(sheet.data ?? {})) {
    const d = sheet.data?.[id] as { componentKey?: string; sheetTransform?: { from: CellAnchor }; data?: { spec?: SparklineSpec } } | undefined;
    if (d?.componentKey === SPARKLINE_COMPONENT_KEY && d.data?.spec && d.sheetTransform) out.push({ spec: d.data.spec, row: d.sheetTransform.from.row, col: d.sheetTransform.from.column });
  }
  return out;
}

/** Ajoute les mini-graphiques lus à la ressource de dessins d'une feuille. */
export function addSparklinesToResource(unitId: string, sheetId: string, sheet: Partial<IWorksheetData>, items: SheetSparkline[], existing?: DrawingResource[string]): DrawingResource[string] {
  const data = { ...(existing?.data ?? {}) };
  const order = [...(existing?.order ?? [])];
  const colWidth = (c: number) => sheet.columnData?.[c]?.w ?? sheet.defaultColumnWidth ?? 88;
  const rowHeight = (r: number) => sheet.rowData?.[r]?.h ?? sheet.defaultRowHeight ?? 24;
  items.forEach((it, i) => {
    const id = `spark${Date.now().toString(36)}${i}`;
    let left = 46;
    for (let c = 0; c < it.col; c++) left += colWidth(c);
    let top = 20;
    for (let r = 0; r < it.row; r++) top += rowHeight(r);
    const from = { column: it.col, columnOffset: 0, row: it.row, rowOffset: 0 };
    const to = { column: it.col + 1, columnOffset: 0, row: it.row + 1, rowOffset: 0 };
    data[id] = {
      unitId,
      subUnitId: sheetId,
      drawingId: id,
      drawingType: DRAWING_DOM_TYPE,
      componentKey: SPARKLINE_COMPONENT_KEY,
      sheetTransform: { from, to },
      axisAlignSheetTransform: { from, to },
      transform: { left, top, width: colWidth(it.col), height: rowHeight(it.row) },
      data: { sparkId: `spark-${id}`, spec: it.spec } as never,
      allowTransform: false,
    } as never;
    order.push(id);
  });
  return { data, order };
}

export { DRAWING_RESOURCE };
