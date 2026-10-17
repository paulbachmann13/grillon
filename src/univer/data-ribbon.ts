/**
 * Onglet Données : outils qu'Univer open source ne fournit pas (Supprimer les doublons,
 * Convertir, Remplissage instantané, Plan/Grouper, Sous-total, Filtre avancé, Tableau,
 * Ligne de totaux, Importer un CSV). La logique est dans src/data (testée à part).
 */
import { IconManager, IEditorBridgeService } from '@univerjs/preset-sheets-core';
import { CommandType, ICommandService, type FUniver } from '@univerjs/presets';
import type { FWorksheet } from '@univerjs/sheets/facade';
import { createElement, type ComponentType } from 'react';
import type { DocumentController } from '../app/document';
import { advancedFilter } from '../data/advanced-filter';
import { cellText, cellValue, isEmpty, type Grid } from '../data/cells';
import { flashFill } from '../data/flash-fill';
import { detailBlock } from '../data/outline';
import { removeDuplicates } from '../data/remove-duplicates';
import { computeSubtotals, type SubtotalFunction } from '../data/subtotals';
import { textToColumns } from '../data/text-to-columns';
import { t } from '../i18n';
import { columnToLetters, rangeFromA1 } from '../io/xlsx/units';
import type { ICellData } from '../model/univer-types';
import { GRILLON_CUSTOM_KEY } from '../model/univer-types';
import { alertDialog, box, checkbox, field, group, openDialog, radio, select, textInput } from '../ui/dialog';
import { currentRegionRect } from './region';
import { confirmAsync } from '../app/confirm';
import { pickTableStyle, registerTableStyles } from './table-styles-ui';

const DATA = 'ribbon.data';
const ORGANIZATION = 'ribbon.data.organization';
const OTHERS = 'ribbon.data.others';

/** Efface la valeur d'une cellule (en gardant sa mise en forme). */
const CLEAR_VALUE: ICellData = { v: null, f: null, p: null } as unknown as ICellData;
/** Efface tout (valeur et mise en forme). */
const CLEAR_ALL: ICellData = { v: null, f: null, p: null, s: null } as unknown as ICellData;

interface Area {
  ws: FWorksheet;
  row: number;
  col: number;
  rows: number;
  cols: number;
}

function sheetOf(api: FUniver): FWorksheet | null {
  return (api.getActiveWorkbook()?.getActiveSheet() as unknown as FWorksheet) ?? null;
}

/** Sélection active. */
function selection(api: FUniver): Area | null {
  const ws = sheetOf(api);
  const r = api.getActiveWorkbook()?.getActiveRange();
  if (!ws || !r) return null;
  return { ws, row: r.getRow(), col: r.getColumn(), rows: r.getLastRow() - r.getRow() + 1, cols: r.getLastColumn() - r.getColumn() + 1 };
}

/** Sélection, ou « zone de données en cours » si une seule cellule est sélectionnée (comme Excel). */
function currentRegion(api: FUniver): Area | null {
  const sel = selection(api);
  if (!sel) return null;
  if (sel.rows > 1 || sel.cols > 1) return sel;
  const r = currentRegionRect(sel.ws, { r1: sel.row, c1: sel.col, r2: sel.row, c2: sel.col });
  const area = { ws: sel.ws, row: r.r1, col: r.c1, rows: r.r2 - r.r1 + 1, cols: r.c2 - r.c1 + 1 };
  return area.rows === 1 && area.cols === 1 && isEmpty(sel.ws.getRange(sel.row, sel.col).getCellData()) ? null : area;
}

function readGrid(a: Area): Grid {
  return a.ws.getRange(a.row, a.col, a.rows, a.cols).getCellDataGrid() as Grid;
}

function writeGrid(ws: FWorksheet, row: number, col: number, grid: (ICellData | null | undefined)[][], blank: ICellData = CLEAR_VALUE): void {
  if (!grid.length) return;
  const width = Math.max(...grid.map((r) => r.length));
  const values = grid.map((r) => Array.from({ length: width }, (_, i) => r[i] ?? blank));
  ws.getRange(row, col, values.length, width).setValues(values as ICellData[][]);
}

function headerLabels(grid: Grid, a: Area, hasHeader: boolean): string[] {
  return Array.from({ length: a.cols }, (_, i) => {
    const h = hasHeader ? cellText(grid[0]?.[i]).trim() : '';
    return h || t('data.column', { letter: columnToLetters(a.col + i) });
  });
}

/** La première ligne ressemble-t-elle à un en-tête (textes, et la ligne suivante contient autre chose) ? */
function looksLikeHeader(grid: Grid): boolean {
  if (grid.length < 2) return false;
  const first = grid[0].map(cellValue);
  if (!first.every((v) => v === null || typeof v === 'string')) return false;
  return grid[1].some((c) => typeof cellValue(c) !== 'string') || first.some((v, i) => v !== null && cellText(grid[1][i]) !== v);
}

function noData(): void {
  void alertDialog(t('data.removeDuplicates'), t('data.noData'));
}

// --- Supprimer les doublons ---

async function removeDuplicatesTool(api: FUniver): Promise<void> {
  const area = currentRegion(api);
  if (!area) return noData();
  const grid = readGrid(area);
  const hasHeader = checkbox(t('data.hasHeaders'), looksLikeHeader(grid));
  const all = checkbox(t('data.selectAll'), true);
  const colsBox = group(t('data.columns'));
  let colChecks: HTMLInputElement[] = [];
  const renderColumns = () => {
    colsBox.querySelectorAll('label').forEach((l) => l.remove());
    colChecks = headerLabels(grid, area, hasHeader.input.checked).map((label) => {
      const c = checkbox(label, true);
      colsBox.append(c.el);
      return c.input;
    });
  };
  renderColumns();
  hasHeader.input.addEventListener('change', renderColumns);
  all.input.addEventListener('change', () => colChecks.forEach((c) => (c.checked = all.input.checked)));
  let message = '';
  const done = await openDialog({
    title: t('data.removeDuplicates'),
    testId: 'dialog-remove-duplicates',
    body: box(t('data.removeDuplicatesHelp'), all.el, hasHeader.el, colsBox),
    onOk: () => {
      const cols = colChecks.flatMap((c, i) => (c.checked ? [i] : []));
      if (!cols.length) return t('data.chooseColumn');
      const res = removeDuplicates(grid, cols, hasHeader.input.checked);
      if (res.removed) {
        const rows: Grid = [...res.rows.map((r) => [...r]), ...Array.from({ length: res.removed }, () => [])];
        writeGrid(area.ws, area.row, area.col, rows, CLEAR_ALL);
      }
      message = res.removed ? t('data.duplicatesRemoved', { removed: res.removed, uniques: res.uniques }) : t('data.noDuplicates');
    },
  });
  if (done) await alertDialog(t('data.removeDuplicates'), message);
}

// --- Convertir (texte en colonnes) ---

async function textToColumnsTool(api: FUniver): Promise<void> {
  const sel = selection(api);
  if (!sel) return;
  if (sel.cols !== 1) return alertDialog(t('data.textToColumnsTitle'), t('data.oneColumnOnly'));
  const area = sel.rows === 1 ? (currentRegion(api) ?? sel) : sel;
  const values = readGrid({ ...area, col: sel.col, cols: 1 }).map((r) => cellText(r[0]));
  const tab = checkbox(t('data.tab'), true);
  const semicolon = checkbox(t('data.semicolon'), true);
  const comma = checkbox(t('data.comma'));
  const space = checkbox(t('data.space'));
  const otherInput = textInput('', '|');
  otherInput.maxLength = 1;
  otherInput.style.width = '48px';
  const other = checkbox(t('data.other'));
  other.el.append(otherInput);
  const consecutive = checkbox(t('data.consecutive'));
  await openDialog({
    title: t('data.textToColumnsTitle'),
    testId: 'dialog-text-to-columns',
    body: box(group(t('data.separators'), tab.el, semicolon.el, comma.el, space.el, other.el), consecutive.el),
    onOk: async () => {
      const rows = textToColumns(values, {
        tab: tab.input.checked,
        semicolon: semicolon.input.checked,
        comma: comma.input.checked,
        space: space.input.checked,
        other: other.input.checked ? otherInput.value : undefined,
        consecutiveAsOne: consecutive.input.checked,
        decimal: ',',
      });
      const width = Math.max(1, ...rows.map((r) => r.length));
      if (width > 1) {
        const right = readGrid({ ws: area.ws, row: area.row, col: sel.col + 1, rows: area.rows, cols: width - 1 });
        if (right.some((r) => r.some((c) => !isEmpty(c))) && !(await confirmAsync(t('data.overwrite')))) return false;
      }
      writeGrid(area.ws, area.row, sel.col, rows);
    },
  });
}

// --- Remplissage instantané ---

export function flashFillTool(api: FUniver): number | null {
  const sel = selection(api);
  if (!sel) return null;
  const { ws, row, col } = sel;
  // Zone : les lignes de données autour de la cellule active, colonnes voisines comprises.
  const region = currentRegionRect(ws, { r1: row, c1: col, r2: row, c2: col });
  const startRow = region.r1;
  const endRow = region.r2;
  const startCol = region.c1;
  const endCol = region.c2;
  if (startCol === endCol) {
    void alertDialog(t('data.flashFill'), t('data.flashFillNone'));
    return null;
  }
  const grid = readGrid({ ws, row: startRow, col: startCol, rows: endRow - startRow + 1, cols: endCol - startCol + 1 });
  const target = col - startCol;
  const toRows = (g: Grid) => g.map((r) => ({ sources: r.filter((_, i) => i !== target).map(cellText), target: cellText(r[target]) }));
  let offset = 0;
  let result = flashFill(toRows(grid));
  // La première ligne est peut-être un en-tête : on réessaie sans elle.
  if (!result && grid.length > 2) {
    offset = 1;
    result = flashFill(toRows(grid.slice(1)));
  }
  if (!result) {
    void alertDialog(t('data.flashFill'), t('data.flashFillNone'));
    return null;
  }
  let count = 0;
  const column = result.map((v, i) => {
    const existing = grid[i + offset][target];
    if (!isEmpty(existing) || v === null) return [existing ?? CLEAR_VALUE];
    count++;
    return [{ v, t: 1 } as ICellData];
  });
  writeGrid(ws, startRow + offset, col, column);
  return count;
}

// --- Plan : grouper / dissocier / masquer ---

type Custom = Record<string, unknown> & { [GRILLON_CUSTOM_KEY]?: { outline?: number } & Record<string, unknown> };

function getLevel(ws: FWorksheet, index: number, cols: boolean): number {
  const custom = (cols ? ws.getColumnCustomMetadata(index) : ws.getRowCustomMetadata(index)) as Custom | undefined;
  return custom?.[GRILLON_CUSTOM_KEY]?.outline ?? 0;
}

/**
 * Fixe les niveaux de plan de plusieurs lignes/colonnes en une seule commande (annulable).
 * setRow/ColumnCustomMetadata n'agit que sur des lignes déjà créées et sans annulation :
 * on passe donc par setRowCustom / setColumnCustom.
 */
function setLevels(ws: FWorksheet, levels: Map<number, number>, cols: boolean): void {
  const batch: Record<number, Custom> = {};
  for (const [index, level] of levels) {
    const custom = ((cols ? ws.getColumnCustomMetadata(index) : ws.getRowCustomMetadata(index)) as Custom | undefined) ?? {};
    const grillon = { ...(custom[GRILLON_CUSTOM_KEY] ?? {}) };
    if (level > 0) grillon.outline = level;
    else delete grillon.outline;
    const next: Custom = { ...custom, [GRILLON_CUSTOM_KEY]: grillon };
    if (!Object.keys(grillon).length) delete next[GRILLON_CUSTOM_KEY];
    batch[index] = next;
  }
  if (cols) ws.setColumnCustom(batch);
  else ws.setRowCustom(batch);
}

function changeLevels(api: FUniver, cols: boolean, delta: number): void {
  const sel = selection(api);
  if (!sel) return;
  const from = cols ? sel.col : sel.row;
  const count = cols ? sel.cols : sel.rows;
  const levels = new Map<number, number>();
  for (let i = from; i < from + count; i++) levels.set(i, Math.max(0, Math.min(7, getLevel(sel.ws, i, cols) + delta)));
  setLevels(sel.ws, levels, cols);
}

/** Bloc de détail autour de la ligne active (ou juste au-dessus : ligne de synthèse, comme Excel). */
function findDetail(ws: FWorksheet, index: number): { from: number; to: number } | null {
  const window = 5000;
  const start = Math.max(0, index - window);
  const levels = Array.from({ length: 2 * window + 1 }, (_, i) => getLevel(ws, start + i, false));
  const local = index - start;
  const block = detailBlock(levels, local) ?? (local > 0 ? detailBlock(levels, local - 1) : null);
  return block ? { from: block.from + start, to: block.to + start } : null;
}

function showHideDetail(api: FUniver, show: boolean): void {
  const sel = selection(api);
  if (!sel) return;
  const block = findDetail(sel.ws, sel.row);
  if (!block) return void alertDialog(t('data.outline'), t('data.noGroup'));
  if (show) sel.ws.showRows(block.from, block.to - block.from + 1);
  else sel.ws.hideRows(block.from, block.to - block.from + 1);
}

// --- Sous-total ---

async function subtotalTool(api: FUniver): Promise<void> {
  const area = currentRegion(api);
  if (!area || area.rows < 2) return noData();
  const grid = readGrid(area);
  const labels = headerLabels(grid, area, true);
  const by = select(labels.map((l, i) => [String(i), l]));
  const fns: [SubtotalFunction, string][] = [
    ['sum', t('data.fnSum')],
    ['count', t('data.fnCount')],
    ['counta', t('data.fnCountA')],
    ['average', t('data.fnAverage')],
    ['max', t('data.fnMax')],
    ['min', t('data.fnMin')],
    ['product', t('data.fnProduct')],
  ];
  const fn = select(fns);
  const lastNumeric = labels.map((_, i) => i).reverse().find((i) => grid.slice(1).some((r) => typeof cellValue(r[i]) === 'number')) ?? labels.length - 1;
  const checks = labels.map((l, i) => checkbox(l, i === lastNumeric));
  await openDialog({
    title: t('data.subtotalTitle'),
    testId: 'dialog-subtotal',
    body: box(field(t('data.atEachChange'), by), field(t('data.useFunction'), fn), group(t('data.addSubtotalTo'), ...checks.map((c) => c.el))),
    onOk: () => {
      const columns = checks.flatMap((c, i) => (c.input.checked ? [i] : []));
      if (!columns.length) return t('data.chooseColumn');
      const res = computeSubtotals(grid, {
        groupBy: Number(by.value),
        fn: fn.value as SubtotalFunction,
        columns,
        startRow: area.row,
        startColumn: area.col,
        labels: { total: (g) => t('data.total', { group: g }), grandTotal: t('data.grandTotal') },
      });
      const extra = res.rows.length - grid.length;
      if (extra > 0) area.ws.insertRowsAfter(area.row + area.rows - 1, extra);
      writeGrid(area.ws, area.row, area.col, res.rows.map((r) => Array.from({ length: area.cols }, (_, i) => r[i] ?? null)));
      setLevels(area.ws, new Map(res.outline.map((level, i) => [area.row + i, level])), false);
    },
  });
}

// --- Filtre avancé ---

function parseRef(ref: string): { startRow: number; endRow: number; startColumn: number; endColumn: number } | null {
  const clean = ref.trim().replace(/^.*!/, '').replace(/\$/g, '');
  if (!clean) return null;
  try {
    return rangeFromA1(clean);
  } catch {
    return null;
  }
}

function absA1(a: Area): string {
  return `$${columnToLetters(a.col)}$${a.row + 1}:$${columnToLetters(a.col + a.cols - 1)}$${a.row + a.rows}`;
}

async function advancedFilterTool(api: FUniver): Promise<void> {
  const region = currentRegion(api);
  const ws = sheetOf(api);
  if (!ws) return;
  const inPlace = radio('g-af-action', t('data.filterInPlace'), 'place', true);
  const copy = radio('g-af-action', t('data.copyTo'), 'copy');
  const list = textInput(region ? absA1(region) : '');
  const crit = textInput('');
  const dest = textInput('');
  dest.disabled = true;
  const unique = checkbox(t('data.uniqueOnly'));
  [inPlace.input, copy.input].forEach((r) => r.addEventListener('change', () => (dest.disabled = !copy.input.checked)));
  await openDialog({
    title: t('data.advancedFilter'),
    testId: 'dialog-advanced-filter',
    body: box(group(t('data.action'), inPlace.el, copy.el), field(t('data.listRange'), list), field(t('data.criteriaRange'), crit), field(t('data.copyToRange'), dest), unique.el),
    onOk: () => {
      const l = parseRef(list.value);
      if (!l) return t('data.invalidRange', { ref: list.value });
      const c = parseRef(crit.value);
      if (crit.value.trim() && !c) return t('data.invalidRange', { ref: crit.value });
      const listArea: Area = { ws, row: l.startRow, col: l.startColumn, rows: l.endRow - l.startRow + 1, cols: l.endColumn - l.startColumn + 1 };
      const listGrid = readGrid(listArea);
      const critGrid = c ? readGrid({ ws, row: c.startRow, col: c.startColumn, rows: c.endRow - c.startRow + 1, cols: c.endColumn - c.startColumn + 1 }) : [];
      const { rows } = advancedFilter(listGrid, critGrid, unique.input.checked);
      if (copy.input.checked) {
        const d = parseRef(dest.value);
        if (!d) return t('data.invalidRange', { ref: dest.value });
        writeGrid(ws, d.startRow, d.startColumn, [listGrid[0], ...rows.map((i) => listGrid[i + 1])]);
        return;
      }
      const keep = new Set(rows);
      ws.showRows(listArea.row + 1, listArea.rows - 1);
      // Masque les blocs contigus de lignes non retenues.
      let i = 0;
      while (i < listArea.rows - 1) {
        if (keep.has(i)) {
          i++;
          continue;
        }
        let j = i;
        while (j + 1 < listArea.rows - 1 && !keep.has(j + 1)) j++;
        ws.hideRows(listArea.row + 1 + i, j - i + 1);
        i = j + 1;
      }
    },
  });
}

function showAllRows(api: FUniver): void {
  const ws = sheetOf(api);
  if (!ws) return;
  const last = ws.getLastRow();
  if (last >= 0) ws.showRows(0, last + 1);
}

// --- Tableaux ---

async function formatAsTable(api: FUniver): Promise<void> {
  const area = currentRegion(api);
  if (!area) return noData();
  const existing = (area.ws as unknown as { getTableByCell(r: number, c: number): unknown }).getTableByCell(area.row, area.col);
  if (existing) return void changeTableStyle(api);
  const pick = await pickTableStyle(t('data.tableFormat'), true);
  if (!pick) return;
  const wb = api.getActiveWorkbook() as unknown as { getTableList(): { name: string }[] };
  const used = new Set(wb.getTableList().map((x) => x.name.toLowerCase()));
  let n = 1;
  while (used.has(`tableau${n}`)) n++;
  // Sans ligne d'en-tête, Excel en ajoute une (Colonne1, Colonne2…) au-dessus des données.
  let row = area.row;
  if (!pick.headers) {
    area.ws.insertRowsBefore(area.row, 1);
    writeGrid(area.ws, area.row, area.col, [Array.from({ length: area.cols }, (_, i): ICellData => ({ v: `Colonne${i + 1}`, t: 1 }))]);
    row = area.row;
  }
  const rows = pick.headers ? area.rows : area.rows + 1;
  const ws = area.ws as unknown as { addTable(name: string, range: object, id?: string, options?: object): Promise<boolean> | boolean };
  await ws.addTable(`Tableau${n}`, { startRow: row, endRow: row + rows - 1, startColumn: area.col, endColumn: area.col + area.cols - 1 }, undefined, { showHeader: true, tableStyleId: pick.name });
}

/** Création → Styles de tableau : change le style du tableau sous la cellule active. */
async function changeTableStyle(api: FUniver): Promise<void> {
  const sel = selection(api);
  if (!sel) return;
  const table = (sel.ws as unknown as { getTableByCell(r: number, c: number): { id: string; tableStyleId?: string; options?: { tableStyleId?: string } } | undefined }).getTableByCell(sel.row, sel.col);
  if (!table) return void alertDialog(t('data.tableStyle'), t('data.noTable'));
  const pick = await pickTableStyle(t('data.tableStyle'), false, table.tableStyleId ?? table.options?.tableStyleId);
  if (!pick) return;
  await api.executeCommand('sheet.command.set-table-config', { unitId: api.getActiveWorkbook()!.getId(), tableId: table.id, theme: pick.name });
}

function addTotalsRow(api: FUniver): void {
  const sel = selection(api);
  if (!sel) return;
  const ws = sel.ws as unknown as FWorksheet & { getTableByCell(r: number, c: number): { name: string; range: { startRow: number; endRow: number; startColumn: number; endColumn: number }; columns: { displayName: string }[] } | undefined };
  const table = ws.getTableByCell(sel.row, sel.col);
  if (!table) return void alertDialog(t('data.totalsRow'), t('data.noTable'));
  const { range, name, columns } = table;
  const below = range.endRow + 1;
  const width = range.endColumn - range.startColumn + 1;
  const belowCells = readGrid({ ws, row: below, col: range.startColumn, rows: 1, cols: width })[0] ?? [];
  if (belowCells.some((c) => !isEmpty(c))) ws.insertRowsAfter(range.endRow, 1);
  const data = readGrid({ ws, row: range.startRow + 1, col: range.startColumn, rows: range.endRow - range.startRow, cols: width });
  const row = columns.map((col, i): ICellData | null => {
    if (i === 0) return { v: t('data.tableTotalsLabel'), t: 1, s: { bl: 1 } };
    const numeric = data.some((r) => typeof cellValue(r[i]) === 'number');
    return numeric ? { f: `=SUBTOTAL(109,${name}[${col.displayName}])`, s: { bl: 1 } } : null;
  });
  writeGrid(ws, below, range.startColumn, [row]);
}

// --- Installation dans le ruban ---

function icon(path: string): ComponentType<Record<string, unknown>> {
  return (props) =>
    createElement('svg', { viewBox: '0 0 16 16', width: '1em', height: '1em', fill: 'none', stroke: 'currentColor', strokeWidth: 1.3, strokeLinecap: 'round', strokeLinejoin: 'round', ...props }, createElement('path', { d: path }));
}

export function installDataRibbon(api: FUniver, doc: DocumentController): void {
  const injector = (api as unknown as { _injector: { get<T>(id: unknown): T } })._injector;
  const commands = injector.get<ICommandService>(ICommandService);
  injector.get<IconManager>(IconManager).register({
    GrillonDedupIcon: icon('M2 3h8M2 6.5h8M2 10h5M11 9l3 3m0-3l-3 3'),
    GrillonSplitIcon: icon('M2 3h12v10H2zM8 3v10M5 6l-1.5 2L5 10M11 6l1.5 2L11 10'),
    GrillonFlashIcon: icon('M9 1.5L3.5 9H8l-1 5.5L12.5 7H8z'),
    GrillonOutlineIcon: icon('M2 3h2M2 8h2M2 13h2M6 3h8M8 8h6M8 13h6M3 4.5v7'),
    GrillonSubtotalIcon: icon('M3 3h10L8 8l5 5H3'),
    GrillonAdvFilterIcon: icon('M2 3h12L9.5 8.5V13l-3-1.5v-3z'),
    GrillonTableIcon: icon('M2 3h12v10H2zM2 6.5h12M2 10h12M6 3v10M10 3v10'),
    GrillonCsvIcon: icon('M4 2h6l3 3v9H4zM10 2v3h3M6 9h5M6 11.5h5'),
  } as never);

  const register = (id: string, run: () => unknown) => {
    if (!commands.hasCommand(id)) commands.registerCommand({ id, type: CommandType.COMMAND, handler: () => (void run(), true) });
  };
  const button = (id: string, icon: string, title: string, run: () => unknown, path: string[] = [DATA, OTHERS]) => {
    register(id, run);
    api.createMenu({ id, icon, title, tooltip: title, action: id }).appendTo(path);
  };
  const dropdown = (id: string, iconName: string, title: string, items: [string, string, () => unknown][], path: string[] = [DATA, OTHERS]) => {
    api.createSubmenu({ id, icon: iconName, title, tooltip: title }).appendTo(path);
    items.forEach(([itemId, itemTitle, run], order) => {
      register(itemId, run);
      api.createMenu({ id: itemId, title: itemTitle, action: itemId, order }).appendTo([...path, id]);
    });
  };

  button('grillon.data.import-csv', 'GrillonCsvIcon', t('data.importCsv'), () => doc.importCsv(), [DATA, 'ribbon.data.formulas']);
  dropdown(
    'grillon.data.advanced',
    'GrillonAdvFilterIcon',
    t('data.advancedFilter'),
    [
      ['grillon.data.advanced-filter', `${t('data.advancedFilter')}…`, () => advancedFilterTool(api)],
      ['grillon.data.show-all', t('data.showAll'), () => showAllRows(api)],
    ],
    [DATA, ORGANIZATION],
  );
  button('grillon.data.text-to-columns', 'GrillonSplitIcon', t('data.textToColumns'), () => textToColumnsTool(api));
  button('grillon.data.flash-fill', 'GrillonFlashIcon', t('data.flashFill'), () => flashFillTool(api));
  button('grillon.data.remove-duplicates', 'GrillonDedupIcon', t('data.removeDuplicates'), () => removeDuplicatesTool(api));
  registerTableStyles(api);
  dropdown('grillon.data.table', 'GrillonTableIcon', t('data.tableFormat'), [
    ['grillon.data.table.create', `${t('data.tableFormat')}…`, () => formatAsTable(api)],
    ['grillon.data.table.style', `${t('data.tableStyle')}…`, () => changeTableStyle(api)],
    ['grillon.data.table.totals', t('data.totalsRow'), () => addTotalsRow(api)],
  ]);
  dropdown('grillon.data.outline', 'GrillonOutlineIcon', t('data.outline'), [
    ['grillon.data.group-rows', t('data.groupRows'), () => changeLevels(api, false, 1)],
    ['grillon.data.ungroup-rows', t('data.ungroupRows'), () => changeLevels(api, false, -1)],
    ['grillon.data.group-cols', t('data.groupCols'), () => changeLevels(api, true, 1)],
    ['grillon.data.ungroup-cols', t('data.ungroupCols'), () => changeLevels(api, true, -1)],
    ['grillon.data.hide-detail', t('data.hideDetail'), () => showHideDetail(api, false)],
    ['grillon.data.show-detail', t('data.showDetail'), () => showHideDetail(api, true)],
  ]);
  // Accueil → Styles : « Mettre sous forme de tableau », comme Excel.
  api.createMenu({ id: 'grillon.home.format-as-table', icon: 'GrillonTableIcon', title: t('data.tableFormat'), tooltip: t('data.tableFormat'), action: 'grillon.data.table.create' }).appendTo(['ribbon.start', 'ribbon.start.others']);
  button('grillon.data.subtotal', 'GrillonSubtotalIcon', t('data.subtotal'), () => subtotalTool(api));

  // Ctrl+E : remplissage instantané (comme Excel), sauf pendant la saisie dans une cellule
  // ou dans un champ. La grille garde le focus sur un éditeur caché ; on demande donc à
  // Univer si l'éditeur de cellule est réellement ouvert.
  const editorBridge = injector.get<IEditorBridgeService>(IEditorBridgeService);
  window.addEventListener(
    'keydown',
    (e) => {
      if (!(e.ctrlKey || e.metaKey) || e.shiftKey || e.altKey || e.key.toLowerCase() !== 'e') return;
      const active = document.activeElement as HTMLElement | null;
      const typing = !!active && (active.tagName === 'INPUT' || active.tagName === 'TEXTAREA' || active.tagName === 'SELECT' || active.isContentEditable);
      // Champ de saisie hors de la grille (boîte de dialogue, barre de formule…) : on laisse faire.
      if (typing && active?.id !== '__editor___INTERNAL_EDITOR__DOCS_NORMAL') return;
      if (editorBridge.isVisible().visible) return;
      e.preventDefault();
      e.stopPropagation();
      void api.executeCommand('grillon.data.flash-fill');
    },
    { capture: true },
  );
}
