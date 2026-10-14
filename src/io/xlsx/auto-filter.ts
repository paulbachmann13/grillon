/**
 * Filtre automatique : <autoFilter> OOXML ⇄ Univer (ressource SHEET_FILTER_PLUGIN = { [sheetId]: IAutoFilter }).
 * ExcelJS n'écrit que la plage du filtre : les critères sont lus et écrits ici, dans le XML.
 */
import type { IRange } from '../../model/univer-types';
import { attrs, getWorksheetElement, parseXml } from './ooxml';
import { rangeFromA1, rangeToA1 } from './units';

export const FILTER_RESOURCE = 'SHEET_FILTER_PLUGIN';

const CUSTOM_OPERATORS = ['equal', 'greaterThan', 'greaterThanOrEqual', 'lessThan', 'lessThanOrEqual', 'notEqual'];

export interface UniverFilterColumn {
  colId: number;
  filters?: { blank?: true; filters?: string[] };
  customFilters?: { and?: 1; customFilters: { val: string | number; operator?: string }[] };
  colorFilters?: unknown;
}

export interface UniverAutoFilter {
  ref: IRange;
  filterColumns?: UniverFilterColumn[];
  cachedFilteredOut?: number[];
}

export type UniverFilterResource = Record<string, UniverAutoFilter>;

interface RawFilterColumn {
  colId?: string;
  filters?: { blank?: string; filter?: { val?: string }[] } | '';
  customFilters?: { and?: string; customFilter?: { operator?: string; val?: string }[] };
}

/** Lit le filtre automatique d'une feuille (null s'il n'y en a pas). */
export function readAutoFilter(sheetXml: string): UniverAutoFilter | null {
  const el = getWorksheetElement(sheetXml, 'autoFilter');
  if (!el) return null;
  const doc = parseXml<{ autoFilter?: { ref?: string; filterColumn?: RawFilterColumn[] } }>(el);
  const af = doc.autoFilter;
  if (!af?.ref) return null;
  const columns: UniverFilterColumn[] = [];
  for (const fc of af.filterColumn ?? []) {
    const colId = Number(fc.colId);
    if (!Number.isFinite(colId)) continue;
    const col: UniverFilterColumn = { colId };
    if (fc.filters !== undefined) {
      const f = typeof fc.filters === 'object' ? fc.filters : {};
      col.filters = {};
      if (f.blank === '1' || f.blank === 'true') col.filters.blank = true;
      const values = (f.filter ?? []).map((x) => x.val ?? '').filter((v) => v !== '');
      if (values.length) col.filters.filters = values;
    } else if (fc.customFilters) {
      const items = (fc.customFilters.customFilter ?? [])
        .slice(0, 2)
        .map((c) => {
          const operator = c.operator && CUSTOM_OPERATORS.includes(c.operator) ? c.operator : undefined;
          const num = Number(c.val);
          const val = c.val !== undefined && c.val !== '' && Number.isFinite(num) && !/^0\d/.test(c.val) ? num : (c.val ?? '');
          return operator && operator !== 'equal' ? { val, operator } : { val };
        });
      if (!items.length) continue;
      col.customFilters = { customFilters: items };
      if (fc.customFilters.and === '1' || fc.customFilters.and === 'true') col.customFilters.and = 1;
    } else {
      // colorFilter, top10, dynamicFilter : non pris en charge par Univer pour l'instant.
      continue;
    }
    columns.push(col);
  }
  const out: UniverAutoFilter = { ref: rangeFromA1(af.ref) };
  if (columns.length) out.filterColumns = columns;
  return out;
}

/** Construit l'élément <autoFilter> complet (avec critères). */
export function buildAutoFilterXml(filter: UniverAutoFilter): string {
  const ref = rangeToA1(filter.ref);
  const cols = (filter.filterColumns ?? [])
    .map((c) => {
      if (c.filters) {
        const items = (c.filters.filters ?? []).map((v) => `<filter${attrs({ val: v })}/>`).join('');
        return `<filterColumn colId="${c.colId}"><filters${attrs({ blank: c.filters.blank ? 1 : undefined })}>${items}</filters></filterColumn>`;
      }
      if (c.customFilters?.customFilters?.length) {
        const items = c.customFilters.customFilters
          .map((cf) => `<customFilter${attrs({ operator: cf.operator && cf.operator !== 'equal' ? cf.operator : undefined, val: String(cf.val) })}/>`)
          .join('');
        return `<filterColumn colId="${c.colId}"><customFilters${attrs({ and: c.customFilters.and ? 1 : undefined })}>${items}</customFilters></filterColumn>`;
      }
      return '';
    })
    .join('');
  return cols ? `<autoFilter ref="${ref}">${cols}</autoFilter>` : `<autoFilter ref="${ref}"/>`;
}
