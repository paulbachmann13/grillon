/**
 * Accueil → Styles de cellule (galerie comme Excel) et boîte « Format de cellule » (Ctrl+1 / ⌘1) :
 * Nombre, Alignement, Police, Bordure, Remplissage. Une seule étape d'annulation par application.
 */
import { IconManager } from '@univerjs/preset-sheets-core';
import { CommandType, ICommandService, type FUniver } from '@univerjs/presets';
import { createElement, type ComponentType } from 'react';
import { formatNumberPattern } from '../charts/format';
import { lockedState, setUnlocked } from '../data/cell-locks';
import { CELL_STYLES, NO_BORDER, rangeStyles, type BorderPart } from '../data/cell-styles';
import { t } from '../i18n';
import type { IStyleData } from '../model/univer-types';
import { box, checkbox, field, group, openDialog, select, textInput } from '../ui/dialog';
import { protectionOf, storeUnlocked, unlockedOf } from './protection';

/** Au-delà, on applique propriété par propriété (sélection de colonnes entières…). */
const MAX_MATRIX_CELLS = 200_000;

type Api = FUniver;

function selection(api: Api) {
  const wb = api.getActiveWorkbook();
  const range = wb?.getActiveRange();
  const ws = wb?.getActiveSheet();
  if (!wb || !range || !ws) return null;
  return { wb, ws, range, rows: range.getLastRow() - range.getRow() + 1, cols: range.getLastColumn() - range.getColumn() + 1 };
}

/** Applique un style (et des bordures) à la sélection, en une seule commande quand c'est possible. */
export async function applyCellStyle(api: Api, style: IStyleData, borders?: Partial<Record<BorderPart, { s: number; cl: { rgb: string } } | null>>): Promise<void> {
  const sel = selection(api);
  if (!sel) return;
  if (sel.rows * sel.cols <= MAX_MATRIX_CELLS) {
    const value = rangeStyles(sel.rows, sel.cols, style, borders).map((row) => row.map((s) => ({ s })));
    await api.executeCommand('sheet.command.set-range-values', { value, range: sel.range.getRange(), unitId: sel.wb.getId(), subUnitId: sel.ws.getSheetId() });
    return;
  }
  // Très grande sélection : une commande par propriété (Univer les applique par plage).
  for (const [type, value] of Object.entries(style)) {
    if (type === 'bd') continue;
    await api.executeCommand('sheet.command.set-style', { style: { type, value } });
  }
}

// --- Galerie des styles de cellule ---

function previewCss(style: IStyleData | null): string {
  if (!style) return '';
  const css: string[] = [];
  if (style.bg?.rgb) css.push(`background:${style.bg.rgb}`);
  if (style.cl?.rgb) css.push(`color:${style.cl.rgb}`);
  if (style.bl) css.push('font-weight:700');
  if (style.it) css.push('font-style:italic');
  if (style.fs) css.push(`font-size:${Math.min(16, style.fs)}px`);
  const b = style.bd as Record<string, { s: number; cl?: { rgb?: string } }> | undefined;
  for (const [k, side] of [['t', 'top'], ['b', 'bottom'], ['l', 'left'], ['r', 'right']] as const) {
    const e = b?.[k];
    if (e?.s) css.push(`border-${side}:${e.s === 13 ? 3 : e.s === 8 ? 2 : e.s === 7 ? 3 : 1}px ${e.s === 7 ? 'double' : 'solid'} ${e.cl?.rgb ?? '#000'}`);
  }
  return css.join(';');
}

export async function cellStylesDialog(api: Api): Promise<void> {
  const body = document.createElement('div');
  body.className = 'g-cell-styles';
  body.dataset.testid = 'cell-styles';
  const groups = [...new Set(CELL_STYLES.map((s) => s.group))];
  for (const g of groups) {
    const h = document.createElement('h3');
    h.textContent = g;
    const row = document.createElement('div');
    row.className = 'g-cell-style-row';
    for (const entry of CELL_STYLES.filter((s) => s.group === g)) {
      const b = document.createElement('button');
      b.type = 'button';
      b.className = 'g-cell-style';
      b.textContent = entry.label;
      b.setAttribute('style', previewCss(entry.style));
      b.addEventListener('click', () => {
        if (entry.style) void applyCellStyle(api, entry.style);
        else void api.executeCommand('sheet.command.clear-selection-format');
        // Un clic applique le style et ferme la galerie, comme dans Excel.
        (b.closest('.g-dialog')?.querySelector('.g-dialog-actions button:not(.primary)') as HTMLButtonElement | null)?.click();
      });
      row.append(b);
    }
    body.append(h, row);
  }
  await openDialog({
    title: t('format.stylesTitle'),
    testId: 'dialog-cell-styles',
    wide: true,
    body: box(body),
    okLabel: t('dialog.close'),
    onOk: () => true,
  });
}

// --- Boîte « Format de cellule » ---

const NUMBER_CATEGORIES: [id: string, label: string][] = [
  ['general', 'Standard'],
  ['number', 'Nombre'],
  ['currency', 'Monétaire'],
  ['accounting', 'Comptabilité'],
  ['date', 'Date'],
  ['time', 'Heure'],
  ['percent', 'Pourcentage'],
  ['fraction', 'Fraction'],
  ['scientific', 'Scientifique'],
  ['text', 'Texte'],
  ['special', 'Spécial'],
  ['custom', 'Personnalisée'],
];

const DATE_FORMATS: [string, string][] = [
  ['dd/mm/yyyy', '14/03/2026'],
  ['d mmmm yyyy', '14 mars 2026'],
  ['dddd d mmmm yyyy', 'samedi 14 mars 2026'],
  ['dd/mm/yy', '14/03/26'],
  ['d-mmm', '14-mars'],
  ['mmm-yy', 'mars-26'],
  ['mmmm yyyy', 'mars 2026'],
  ['yyyy-mm-dd', '2026-03-14'],
  ['dd/mm/yyyy hh:mm', '14/03/2026 13:30'],
];
const TIME_FORMATS: [string, string][] = [
  ['hh:mm', '13:30'],
  ['hh:mm:ss', '13:30:55'],
  ['h:mm AM/PM', '1:30 PM'],
  ['[h]:mm:ss', '37:30:55'],
  ['mm:ss.0', '30:55.2'],
];
const SPECIAL_FORMATS: [string, string][] = [
  ['00000', 'Code postal'],
  ['00" "00" "00" "00" "00', 'Numéro de téléphone'],
  ['0" "00" "00" "00" "000" "000" "00', 'Numéro de sécurité sociale'],
];
const CUSTOM_FORMATS = ['0', '0.00', '#,##0', '#,##0.00', '#,##0.00 "€"', '0%', '0.00%', '0.00E+00', '# ?/?', '@', 'dd/mm/yyyy', 'hh:mm', '#,##0;[Red]-#,##0'];

function buildPattern(category: string, decimals: number, thousands: boolean, symbol: string, negativeRed: boolean, pick: string): string | null {
  const dec = decimals > 0 ? `.${'0'.repeat(decimals)}` : '';
  const base = `${thousands ? '#,##0' : '0'}${dec}`;
  switch (category) {
    case 'general':
      return 'General';
    case 'number':
      return negativeRed ? `${base};[Red]-${base}` : base;
    case 'currency': {
      const s = symbol ? ` "${symbol}"` : '';
      const p = `#,##0${dec}${s}`;
      return negativeRed ? `${p};[Red]-${p}` : p;
    }
    case 'accounting':
      return `_-* #,##0${dec}${symbol ? ` "${symbol}"` : ''}_-;-* #,##0${dec}${symbol ? ` "${symbol}"` : ''}_-;_-* "-"??${symbol ? ` "${symbol}"` : ''}_-;_-@_-`;
    case 'percent':
      return `0${dec}%`;
    case 'scientific':
      return `0${dec}E+00`;
    case 'fraction':
      return pick || '# ?/?';
    case 'text':
      return '@';
    case 'date':
    case 'time':
    case 'special':
    case 'custom':
      return pick || null;
    default:
      return null;
  }
}

function categoryOf(pattern: string | undefined): string {
  if (!pattern || pattern === 'General') return 'general';
  if (pattern === '@') return 'text';
  if (/E\+0/i.test(pattern)) return 'scientific';
  if (pattern.includes('%')) return 'percent';
  if (/\?\/\?/.test(pattern)) return 'fraction';
  if (/^_-\*/.test(pattern)) return 'accounting';
  if (/[€$£]/.test(pattern)) return 'currency';
  if (/[dmy]/i.test(pattern.replace(/"[^"]*"/g, '')) && /[dy]/i.test(pattern)) return 'date';
  if (/h|ss/i.test(pattern.replace(/"[^"]*"/g, ''))) return 'time';
  if (/^[#0,.;\-[\]Red]+$/.test(pattern)) return 'number';
  return 'custom';
}

const FONTS = ['Calibri', 'Aptos', 'Aptos Narrow', 'Arial', 'Cambria', 'Candara', 'Consolas', 'Courier New', 'Georgia', 'Helvetica', 'Segoe UI', 'Tahoma', 'Times New Roman', 'Trebuchet MS', 'Verdana'];

export async function formatCellsDialog(api: Api, startTab = 'number'): Promise<void> {
  const sel = selection(api);
  if (!sel) return;
  const cell = sel.ws.getRange(sel.range.getRow(), sel.range.getColumn());
  const current = (cell.getCellStyleData() ?? {}) as IStyleData;
  const sample = cell.getCellDataGrid()[0]?.[0]?.v;
  const sampleNumber = typeof sample === 'number' ? sample : Number.isFinite(Number(sample)) && sample !== '' && sample !== undefined && sample !== null ? Number(sample) : 1234.5678;
  const patch: IStyleData = {};
  const changed = new Set<string>();
  const mark = (key: keyof IStyleData, value: unknown) => {
    (patch as Record<string, unknown>)[key] = value;
    changed.add(key);
  };

  // --- Nombre ---
  const currentPattern = current.n?.pattern ?? (cell.getNumberFormat() || 'General');
  const category = select(NUMBER_CATEGORIES, categoryOf(currentPattern));
  const decimals = Object.assign(document.createElement('input'), { type: 'number', min: '0', max: '15', value: String((/\.(0+)/.exec(currentPattern)?.[1] ?? '').length || 2) });
  const thousands = checkbox(t('format.thousands'), /#,##0/.test(currentPattern));
  const symbol = select([['€', '€'], ['$', '$'], ['£', '£'], ['CHF', 'CHF'], ['', t('format.none')]], /\$/.test(currentPattern) ? '$' : '€');
  const negativeRed = checkbox(t('format.negativeRed'), /\[Red\]/i.test(currentPattern));
  const list = document.createElement('select');
  list.size = 7;
  const custom = textInput(currentPattern);
  const preview = document.createElement('p');
  preview.className = 'g-format-preview';
  const options = document.createElement('div');
  options.className = 'g-format-options';
  const fillList = (items: [string, string][]) => {
    list.replaceChildren(...items.map(([v, label]) => Object.assign(document.createElement('option'), { value: v, textContent: label })));
    list.value = items.find(([v]) => v === currentPattern)?.[0] ?? items[0]?.[0] ?? '';
  };
  const pattern = () => buildPattern(category.value, Number(decimals.value) || 0, thousands.input.checked, symbol.value, negativeRed.input.checked, category.value === 'custom' ? custom.value : list.value);
  const refreshNumber = () => {
    const c = category.value;
    const show = (els: HTMLElement[]) => options.replaceChildren(...els);
    const decField = field(t('format.decimals'), decimals);
    if (c === 'number') show([decField, thousands.el, negativeRed.el]);
    else if (c === 'currency') show([decField, field(t('format.symbol'), symbol), negativeRed.el]);
    else if (c === 'accounting') show([decField, field(t('format.symbol'), symbol)]);
    else if (c === 'percent' || c === 'scientific') show([decField]);
    else if (c === 'date') (fillList(DATE_FORMATS), show([field(t('format.type'), list)]));
    else if (c === 'time') (fillList(TIME_FORMATS), show([field(t('format.type'), list)]));
    else if (c === 'fraction') (fillList([['# ?/?', 'Un chiffre (1/4)'], ['# ??/??', 'Deux chiffres (21/25)'], ['# ?/2', 'Demis (1/2)'], ['# ?/4', 'Quarts (2/4)'], ['# ?/8', 'Huitièmes']]), show([field(t('format.type'), list)]));
    else if (c === 'special') (fillList(SPECIAL_FORMATS), show([field(t('format.type'), list)]));
    else if (c === 'custom') {
      const presets = select(CUSTOM_FORMATS.map((f) => [f, f] as [string, string]), CUSTOM_FORMATS.includes(custom.value) ? custom.value : CUSTOM_FORMATS[0]);
      presets.addEventListener('change', () => {
        custom.value = presets.value;
        update();
      });
      show([field(t('format.code'), custom), field(t('format.presets'), presets)]);
    } else show([]);
    update();
  };
  const update = () => {
    const p = pattern();
    if (p) mark('n', { pattern: p });
    const sampleText = category.value === 'date' || category.value === 'time' ? (list.selectedOptions[0]?.textContent ?? '') : category.value === 'text' ? String(sample ?? '') : formatNumberPattern(sampleNumber, p ?? undefined);
    preview.textContent = `${t('format.sample')} : ${sampleText}`;
  };
  for (const el of [category]) el.addEventListener('change', refreshNumber);
  for (const el of [decimals, custom]) el.addEventListener('input', update);
  for (const el of [thousands.input, negativeRed.input, symbol, list]) el.addEventListener('change', update);
  const numberPanel = box(field(t('format.category'), category), options, preview);
  refreshNumber();
  changed.delete('n'); // le format n'est appliqué que si l'utilisateur l'a modifié
  delete patch.n;
  const touchNumber = () => changed.add('n');
  numberPanel.addEventListener('input', touchNumber);
  numberPanel.addEventListener('change', touchNumber);

  // --- Alignement ---
  const hAlign = select([['0', 'Standard'], ['1', 'Gauche (Retrait)'], ['2', 'Centré'], ['3', 'Droite (Retrait)'], ['4', 'Justifié']], String(current.ht ?? 0));
  const vAlign = select([['1', 'Haut'], ['2', 'Centré'], ['3', 'Bas']], String(current.vt || 3));
  const indent = Object.assign(document.createElement('input'), { type: 'number', min: '0', max: '15', value: String(Math.round((current.pd?.l ?? 0) / 9)) });
  const angle = Object.assign(document.createElement('input'), { type: 'number', min: '-90', max: '90', value: String(current.tr?.a ?? 0) });
  const vertical = checkbox(t('format.verticalText'), !!current.tr?.v);
  const wrap = checkbox(t('format.wrap'), current.tb === 3);
  const merge = checkbox(t('format.merge'), false);
  hAlign.addEventListener('change', () => mark('ht', Number(hAlign.value)));
  vAlign.addEventListener('change', () => mark('vt', Number(vAlign.value)));
  indent.addEventListener('input', () => mark('pd', { ...(current.pd ?? {}), l: Math.max(0, Number(indent.value) || 0) * 9 }));
  const rotation = () => mark('tr', { a: vertical.input.checked ? 0 : Math.max(-90, Math.min(90, Number(angle.value) || 0)), v: vertical.input.checked ? 1 : 0 });
  angle.addEventListener('input', rotation);
  vertical.input.addEventListener('change', rotation);
  wrap.input.addEventListener('change', () => mark('tb', wrap.input.checked ? 3 : 1));
  const alignPanel = box(
    group(t('format.textAlign'), field(t('format.horizontal'), hAlign), field(t('format.vertical'), vAlign), field(t('format.indent'), indent)),
    group(t('format.orientation'), field(t('format.degrees'), angle), vertical.el),
    group(t('format.textControl'), wrap.el, merge.el),
  );

  // --- Police ---
  const fontName = textInput(current.ff ?? 'Calibri');
  fontName.setAttribute('list', 'g-font-list');
  const fontList = document.createElement('datalist');
  fontList.id = 'g-font-list';
  fontList.append(...FONTS.map((f) => Object.assign(document.createElement('option'), { value: f })));
  const fontStyle = select([['0', 'Normal'], ['1', 'Italique'], ['2', 'Gras'], ['3', 'Gras italique']], String((current.bl ? 2 : 0) + (current.it ? 1 : 0)));
  const fontSize = Object.assign(document.createElement('input'), { type: 'number', min: '1', max: '409', step: '0.5', value: String(current.fs ?? 11) });
  const underline = select([['0', 'Aucun'], ['1', 'Simple'], ['2', 'Double']], String(current.ul?.s ? (current.ul as { t?: number }).t === 2 ? 2 : 1 : 0));
  const color = Object.assign(document.createElement('input'), { type: 'color', value: current.cl?.rgb && /^#[0-9a-f]{6}$/i.test(current.cl.rgb) ? current.cl.rgb : '#000000' });
  const strike = checkbox(t('format.strike'), !!current.st?.s);
  const script = select([['1', 'Normal'], ['3', 'Exposant'], ['2', 'Indice']], String(current.va ?? 1));
  fontName.addEventListener('input', () => fontName.value.trim() && mark('ff', fontName.value.trim()));
  fontStyle.addEventListener('change', () => {
    const v = Number(fontStyle.value);
    mark('bl', v >= 2 ? 1 : 0);
    mark('it', v % 2 ? 1 : 0);
  });
  fontSize.addEventListener('input', () => Number(fontSize.value) > 0 && mark('fs', Number(fontSize.value)));
  underline.addEventListener('change', () => mark('ul', { s: underline.value === '0' ? 0 : 1, ...(underline.value === '2' ? { t: 2 } : {}) }));
  color.addEventListener('input', () => mark('cl', { rgb: color.value }));
  strike.input.addEventListener('change', () => mark('st', { s: strike.input.checked ? 1 : 0 }));
  script.addEventListener('change', () => mark('va', Number(script.value)));
  const fontPanel = box(field(t('format.font'), fontName), fontList, field(t('format.fontStyle'), fontStyle), field(t('format.size'), fontSize), field(t('format.underline'), underline), field(t('format.color'), color), strike.el, field(t('format.effects'), script));

  // --- Bordure ---
  const borderStyle = select([['1', 'Fin'], ['8', 'Moyen'], ['13', 'Épais'], ['4', 'Tirets'], ['3', 'Pointillés'], ['7', 'Double']], '1');
  const borderColor = Object.assign(document.createElement('input'), { type: 'color', value: '#000000' });
  const parts: Partial<Record<BorderPart, boolean | null>> = {};
  const edgeBoxes: Record<BorderPart, ReturnType<typeof checkbox>> = {
    t: checkbox(t('format.borderTop'), false),
    b: checkbox(t('format.borderBottom'), false),
    l: checkbox(t('format.borderLeft'), false),
    r: checkbox(t('format.borderRight'), false),
    h: checkbox(t('format.borderInsideH'), false),
    v: checkbox(t('format.borderInsideV'), false),
  };
  for (const [k, c] of Object.entries(edgeBoxes) as [BorderPart, ReturnType<typeof checkbox>][]) c.input.addEventListener('change', () => (parts[k] = c.input.checked));
  const preset = (on: BorderPart[], off: BorderPart[]) => () => {
    for (const k of on) (edgeBoxes[k].input.checked = true), (parts[k] = true);
    for (const k of off) (edgeBoxes[k].input.checked = false), (parts[k] = false);
  };
  const presetButton = (label: string, fn: () => void) => Object.assign(document.createElement('button'), { type: 'button', textContent: label, onclick: fn });
  const presets = document.createElement('div');
  presets.className = 'g-row';
  presets.append(
    presetButton(t('format.borderNone'), preset([], ['t', 'b', 'l', 'r', 'h', 'v'])),
    presetButton(t('format.borderOutline'), preset(['t', 'b', 'l', 'r'], [])),
    presetButton(t('format.borderInside'), preset(['h', 'v'], [])),
  );
  const edges = document.createElement('div');
  edges.className = 'g-csv-settings';
  edges.append(...Object.values(edgeBoxes).map((c) => c.el));
  const borderPanel = box(group(t('format.presets'), presets), group(t('format.borderEdges'), edges), field(t('format.lineStyle'), borderStyle), field(t('format.color'), borderColor));

  // --- Remplissage ---
  const fill = Object.assign(document.createElement('input'), { type: 'color', value: current.bg?.rgb && /^#[0-9a-f]{6}$/i.test(current.bg.rgb) ? current.bg.rgb : '#FFFFFF' });
  const noFill = presetButton(t('format.noFill'), () => {
    mark('bg', { rgb: null });
    fill.value = '#FFFFFF';
  });
  fill.addEventListener('input', () => mark('bg', { rgb: fill.value }));
  const swatches = document.createElement('div');
  swatches.className = 'g-swatches';
  for (const c of ['#FFFFFF', '#F2F2F2', '#D9D9D9', '#FFF2CC', '#FCE4D6', '#E2EFDA', '#DDEBF7', '#EDE2F6', '#FFC7CE', '#FFEB9C', '#C6EFCE', '#BDD7EE', '#FFC000', '#ED7D31', '#70AD47', '#4472C4', '#7030A0', '#C00000']) {
    const b = presetButton('', () => {
      fill.value = c;
      mark('bg', { rgb: c });
    });
    b.style.background = c;
    b.title = c;
    swatches.append(b);
  }
  const fillPanel = box(field(t('format.fillColor'), fill), swatches, noFill);

  // --- Protection ---
  const sheetRange = sel.range.getRange();
  const initialLock = lockedState(unlockedOf(sel.ws), sheetRange);
  const locked = checkbox(t('format.locked'), initialLock !== false);
  locked.input.indeterminate = initialLock === null;
  locked.input.dataset.testid = 'format-locked';
  let lockChanged = false;
  locked.input.addEventListener('change', () => (lockChanged = true));
  const sheetProtected = !!protectionOf(sel.ws);
  locked.input.disabled = sheetProtected;
  const protectionPanel = box(locked.el, sheetProtected ? t('format.lockedProtected') : t('format.lockedHelp'));

  // --- Onglets ---
  const tabs: [string, string, HTMLElement][] = [
    ['number', t('format.tabNumber'), numberPanel],
    ['alignment', t('format.tabAlignment'), alignPanel],
    ['font', t('format.tabFont'), fontPanel],
    ['border', t('format.tabBorder'), borderPanel],
    ['fill', t('format.tabFill'), fillPanel],
    ['protection', t('format.tabProtection'), protectionPanel],
  ];
  const bar = document.createElement('div');
  bar.className = 'g-chart-tabs';
  bar.setAttribute('role', 'tablist');
  const host = document.createElement('div');
  host.className = 'g-chart-panel';
  const show = (id: string) => {
    for (const b of bar.querySelectorAll('button')) {
      b.classList.toggle('selected', b.dataset.tab === id);
      b.setAttribute('aria-selected', String(b.dataset.tab === id));
    }
    host.replaceChildren(tabs.find((x) => x[0] === id)![2]);
  };
  for (const [id, label] of tabs) {
    const b = presetButton(label, () => show(id));
    b.setAttribute('role', 'tab');
    b.dataset.tab = id;
    bar.append(b);
  }
  show(startTab);

  const ok = await openDialog({ title: t('format.title'), testId: 'dialog-format-cells', wide: true, body: box(bar, host), onOk: () => true });
  if (!ok) return;
  if (lockChanged && !sheetProtected) storeUnlocked(sel.ws, setUnlocked(unlockedOf(sel.ws), sheetRange, !locked.input.checked));
  const style: IStyleData = {};
  for (const k of changed) (style as Record<string, unknown>)[k] = (patch as Record<string, unknown>)[k];
  if (changed.has('n')) {
    const p = pattern();
    if (p) style.n = { pattern: p };
    else delete style.n;
  }
  const line = { s: Number(borderStyle.value), cl: { rgb: borderColor.value } };
  const borders: Partial<Record<BorderPart, typeof line | null>> = {};
  for (const [k, on] of Object.entries(parts) as [BorderPart, boolean][]) borders[k] = on ? line : NO_BORDER;
  const hasBorders = Object.keys(borders).length > 0;
  if (Object.keys(style).length || hasBorders) await applyCellStyle(api, style, hasBorders ? borders : undefined);
  if (merge.input.checked && sel.rows * sel.cols > 1) await api.executeCommand('sheet.command.add-worksheet-merge-all');
}

function icon(path: string): ComponentType<Record<string, unknown>> {
  return (props) =>
    createElement('svg', { viewBox: '0 0 16 16', width: '1em', height: '1em', fill: 'none', stroke: 'currentColor', strokeWidth: 1.3, strokeLinecap: 'round', strokeLinejoin: 'round', ...props }, createElement('path', { d: path }));
}

export function installCellFormat(api: Api): void {
  const injector = (api as unknown as { _injector: { get<T>(id: unknown): T } })._injector;
  const commands = injector.get<ICommandService>(ICommandService);
  injector.get<IconManager>(IconManager).register({
    GrillonCellStylesIcon: icon('M2 3h12v10H2zM2 7h12M6 3v10'),
    GrillonFormatCellsIcon: icon('M2 2h12v12H2zM5 5h6M5 8h6M5 11h3'),
  } as never);
  const register = (id: string, run: () => unknown) => {
    if (!commands.hasCommand(id)) commands.registerCommand({ id, type: CommandType.COMMAND, handler: () => (void run(), true) });
  };
  register('grillon.cell-styles', () => cellStylesDialog(api));
  register('grillon.format-cells', () => formatCellsDialog(api));
  api.createMenu({ id: 'grillon.cell-styles', icon: 'GrillonCellStylesIcon', title: t('format.stylesTitle'), tooltip: t('format.stylesTitle'), action: 'grillon.cell-styles' }).appendTo(['ribbon.start', 'ribbon.start.others']);

  // Ctrl+1 / ⌘1 (touche physique « 1 », valable sur AZERTY : Ctrl+&).
  window.addEventListener(
    'keydown',
    (e) => {
      if (!(e.ctrlKey || e.metaKey) || e.shiftKey || e.altKey || (e.code !== 'Digit1' && e.code !== 'Numpad1')) return;
      if (document.querySelector('.g-dialog')) return;
      const active = document.activeElement as HTMLElement | null;
      if (active && (active.tagName === 'INPUT' || active.tagName === 'TEXTAREA')) return;
      e.preventDefault();
      e.stopPropagation();
      void formatCellsDialog(api);
    },
    { capture: true },
  );
}
