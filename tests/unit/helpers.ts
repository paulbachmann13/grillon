import { DEFINED_NAME_RESOURCE, type ICellData, type IStyleData, type IWorkbookData, type IWorksheetData } from '../../src/model/univer-types';
import { toA1 } from '../../src/io/xlsx/units';

export function sheetByName(wb: IWorkbookData, name: string): Partial<IWorksheetData> {
  const s = Object.values(wb.sheets).find((x) => x.name === name);
  if (!s) throw new Error(`Feuille introuvable : ${name}`);
  return s;
}

export function styleOf(wb: IWorkbookData, cell: ICellData | undefined): IStyleData {
  if (!cell?.s) return {};
  return (typeof cell.s === 'string' ? wb.styles[cell.s] : cell.s) ?? {};
}

export function cellAt(wb: IWorkbookData, sheet: string, a1: string): ICellData | undefined {
  const s = sheetByName(wb, sheet);
  const m = /^([A-Z]+)(\d+)$/.exec(a1)!;
  let col = 0;
  for (const ch of m[1]) col = col * 26 + ch.charCodeAt(0) - 64;
  return s.cellData?.[Number(m[2]) - 1]?.[col - 1];
}

/** Représentation indépendante des identifiants, pour comparer deux imports. */
export function normalize(wb: IWorkbookData) {
  return wb.sheetOrder.map((id) => {
    const s = wb.sheets[id]!;
    const cells: Record<string, unknown> = {};
    for (const [r, cols] of Object.entries(s.cellData ?? {})) {
      for (const [c, cell] of Object.entries(cols ?? {}) as [string, ICellData | undefined][]) {
        if (!cell) continue;
        cells[toA1(Number(r), Number(c))] = {
          v: cell.v,
          t: cell.t,
          f: cell.f,
          s: styleOf(wb, cell),
          custom: cell.custom,
          p: cell.p ? { text: cell.p.body?.dataStream, runs: (cell.p.body as { textRuns?: unknown[] })?.textRuns } : undefined,
        };
      }
    }
    return {
      name: s.name,
      hidden: s.hidden,
      tabColor: s.tabColor,
      freeze: s.freeze,
      showGridlines: s.showGridlines,
      defaultColumnWidth: s.defaultColumnWidth,
      defaultRowHeight: s.defaultRowHeight,
      merges: [...(s.mergeData ?? [])].sort((a, b) => a.startRow - b.startRow || a.startColumn - b.startColumn),
      columns: s.columnData,
      rows: s.rowData,
      cells,
    };
  });
}

export function definedNames(wb: IWorkbookData): { name: string; ref: string }[] {
  const res = wb.resources?.find((r) => r.name === DEFINED_NAME_RESOURCE);
  if (!res) return [];
  return Object.values(JSON.parse(res.data) as Record<string, { name: string; formulaOrRefString: string }>)
    .map((n) => ({ name: n.name, ref: n.formulaOrRefString }))
    .sort((a, b) => a.name.localeCompare(b.name));
}

/** Mises en forme conditionnelles par nom de feuille, sans identifiants. */
export function conditionalFormats(wb: IWorkbookData): Record<string, unknown[]> {
  const res = wb.resources?.find((r) => r.name === 'SHEET_CONDITIONAL_FORMATTING_PLUGIN');
  if (!res) return {};
  const data = JSON.parse(res.data) as Record<string, { ranges: Record<string, unknown>[]; stopIfTrue: boolean; rule: unknown }[]>;
  const out: Record<string, unknown[]> = {};
  for (const [sheetId, rules] of Object.entries(data)) {
    const name = wb.sheets[sheetId]?.name ?? sheetId;
    out[name] = rules.map((r) => ({
      ranges: r.ranges.map(({ startRow, endRow, startColumn, endColumn }) => ({ startRow, endRow, startColumn, endColumn })),
      stopIfTrue: r.stopIfTrue,
      rule: r.rule,
    }));
  }
  return out;
}
