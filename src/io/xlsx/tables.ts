/**
 * Tableaux structurés (Excel Tables) : parties xl/tables/tableN.xml ⇄ Univer
 * (ressource SHEET_TABLE_PLUGIN = { [sheetId]: { tables, tableFilteredOutRows } }).
 *
 * Univer ne gère pas encore la ligne de totaux : le tableau Univer couvre l'en-tête et les
 * données, et la ligne de totaux reste une ligne de cellules ordinaires (formules SUBTOTAL)
 * juste en dessous. Les informations Excel (style, totaux) sont gardées dans `meta.excel`.
 */
import type JSZip from 'jszip';
import type { IRange } from '../../model/univer-types';
import { addContentTypeOverride, addRelationship, attrs, insertWorksheetElement, parseXml, readRelationships, resolveRelTarget } from './ooxml';
import { rangeFromA1, rangeToA1 } from './units';
import { excelStyleName, galleryStyleFor } from '../../data/table-styles';

export const TABLE_RESOURCE = 'SHEET_TABLE_PLUGIN';
const TABLE_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table';
const TABLE_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml';

export interface UniverTableColumn {
  id: string;
  displayName: string;
  dataType: string;
  formula: string;
  showFilterButton: boolean;
  meta: Record<string, unknown>;
  style: Record<string, unknown>;
  totalsRowFunction?: string;
  totalsRowLabel?: string;
}

export interface UniverTable {
  id: string;
  name: string;
  range: IRange;
  options: { showHeader?: boolean; showFooter?: boolean; showAutoFilter?: boolean; tableStyleId?: string };
  filters: { tableColumnFilterList: unknown[] };
  columns: UniverTableColumn[];
  meta?: { excel?: ExcelTableMeta } & Record<string, unknown>;
}

export interface ExcelTableMeta {
  totalsRowCount?: number;
  styleName?: string;
  showRowStripes?: boolean;
  showColumnStripes?: boolean;
  showFirstColumn?: boolean;
  showLastColumn?: boolean;
}

export type UniverTableResource = Record<string, { tables: UniverTable[]; tableFilteredOutRows: number[] }>;

interface RawTable {
  table?: {
    id?: string;
    name?: string;
    displayName?: string;
    ref?: string;
    headerRowCount?: string;
    totalsRowCount?: string;
    autoFilter?: unknown;
    tableColumns?: { tableColumn?: { id?: string; name?: string; totalsRowFunction?: string; totalsRowLabel?: string }[] };
    tableStyleInfo?: { name?: string; showRowStripes?: string; showColumnStripes?: string; showFirstColumn?: string; showLastColumn?: string };
  };
}

let counter = 0;
const newId = (p: string) => `${p}${Date.now().toString(36)}${(counter = (counter + 1) % 1_000_000).toString(36)}`;
const bool = (v: string | undefined, def = false) => (v === undefined ? def : v === '1' || v === 'true');

/** Lit les tableaux d'une feuille (via les relations de la feuille). */
export async function readTables(zip: JSZip, sheetPath: string): Promise<UniverTable[]> {
  const rels = await readRelationships(zip, sheetPath);
  const out: UniverTable[] = [];
  for (const rel of rels.filter((r) => r.Type === TABLE_REL_TYPE)) {
    const file = zip.file(resolveRelTarget(sheetPath, rel.Target));
    if (!file) continue;
    const doc = parseXml<RawTable>(await file.async('string'));
    const t = doc.table;
    if (!t?.ref) continue;
    const full = rangeFromA1(t.ref);
    const totals = Number(t.totalsRowCount ?? 0) || 0;
    const header = t.headerRowCount === undefined ? 1 : Number(t.headerRowCount);
    const range = { ...full, endRow: full.endRow - totals };
    const style = t.tableStyleInfo ?? {};
    out.push({
      id: newId('tbl'),
      name: t.displayName ?? t.name ?? `Tableau${out.length + 1}`,
      range,
      options: { showHeader: header !== 0, showFooter: false, showAutoFilter: t.autoFilter !== undefined, tableStyleId: galleryStyleFor(style.name) },
      filters: { tableColumnFilterList: [] },
      columns: (t.tableColumns?.tableColumn ?? []).map((c, i) => {
        const col: UniverTableColumn = {
          id: newId('col'),
          displayName: c.name ?? `Colonne${i + 1}`,
          dataType: 'string',
          formula: '',
          showFilterButton: true,
          meta: {},
          style: {},
        };
        if (c.totalsRowFunction) col.totalsRowFunction = c.totalsRowFunction;
        if (c.totalsRowLabel) col.totalsRowLabel = c.totalsRowLabel;
        return col;
      }),
      meta: {
        excel: {
          totalsRowCount: totals || undefined,
          styleName: style.name,
          showRowStripes: bool(style.showRowStripes, true),
          showColumnStripes: bool(style.showColumnStripes),
          showFirstColumn: bool(style.showFirstColumn),
          showLastColumn: bool(style.showLastColumn),
        },
      },
    });
  }
  return out;
}

/** Nom de tableau valide pour Excel (lettre ou _ en tête, sans espace). */
export function sanitizeTableName(name: string, used: Set<string>): string {
  let base = name.replace(/[^\p{L}\p{N}_.]/gu, '_');
  if (!/^[\p{L}_]/u.test(base)) base = `_${base}`;
  let candidate = base;
  for (let i = 2; used.has(candidate.toLowerCase()); i++) candidate = `${base}${i}`;
  used.add(candidate.toLowerCase());
  return candidate;
}

export interface TableToWrite {
  table: UniverTable;
  /** Noms d'en-tête effectivement écrits dans les cellules (uniques, non vides). */
  headerNames: string[];
  /** Critères de filtre (<filterColumn>) posés par les segments. */
  filterXml?: string;
}

/**
 * Ajoute les parties tableau d'une feuille au paquet et renvoie le XML de feuille modifié.
 * `nextTableId` est un compteur global au classeur (les id de tableau doivent être uniques).
 */
export async function writeTables(
  zip: JSZip,
  sheetPath: string,
  sheetXml: string,
  tables: TableToWrite[],
  nextTableId: () => number,
  usedNames: Set<string>,
): Promise<string> {
  if (!tables.length) return sheetXml;
  const relIds: string[] = [];
  for (const { table, headerNames, filterXml } of tables) {
    const id = nextTableId();
    const excel = table.meta?.excel ?? {};
    const totals = excel.totalsRowCount ?? 0;
    const full = { ...table.range, endRow: table.range.endRow + totals };
    const name = sanitizeTableName(table.name, usedNames);
    const header = table.options.showHeader !== false;
    const columns = headerNames
      .map((colName, i) => {
        const col = table.columns[i];
        return `<tableColumn${attrs({
          id: i + 1,
          name: colName,
          totalsRowLabel: totals ? col?.totalsRowLabel : undefined,
          totalsRowFunction: totals ? col?.totalsRowFunction : undefined,
        })}/>`;
      })
      .join('');
    const filterRange = { ...table.range, startRow: table.range.startRow };
    const xml =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
      `<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"${attrs({
        id,
        name,
        displayName: name,
        ref: rangeToA1(full),
        headerRowCount: header ? undefined : 0,
        totalsRowCount: totals || undefined,
      })}>` +
      (header && table.options.showAutoFilter !== false ? `<autoFilter ref="${rangeToA1(filterRange)}"${filterXml ? `>${filterXml}</autoFilter>` : '/>'}` : '') +
      `<tableColumns count="${headerNames.length}">${columns}</tableColumns>` +
      `<tableStyleInfo${attrs({
        // Nom vide : tableau sans style (créé pour porter un segment).
        name: excelStyleName(table.options.tableStyleId, excel.styleName) || undefined,
        showFirstColumn: excel.showFirstColumn ? 1 : 0,
        showLastColumn: excel.showLastColumn ? 1 : 0,
        showRowStripes: excel.showRowStripes === false ? 0 : 1,
        showColumnStripes: excel.showColumnStripes ? 1 : 0,
      })}/>` +
      `</table>`;
    const partPath = `xl/tables/table${id}.xml`;
    zip.file(partPath, xml);
    await addContentTypeOverride(zip, `/${partPath}`, TABLE_CONTENT_TYPE);
    relIds.push(await addRelationship(zip, sheetPath, TABLE_REL_TYPE, `../tables/table${id}.xml`));
  }
  let out = sheetXml;
  if (!/xmlns:r=/.test(out.slice(0, 2000))) {
    out = out.replace(/<worksheet\b/, '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"');
  }
  const parts = relIds.map((rid) => `<tablePart r:id="${rid}"/>`).join('');
  return insertWorksheetElement(out, 'tableParts', `<tableParts count="${relIds.length}">${parts}</tableParts>`);
}
