/**
 * Boîte « Mise en page » (comme Excel) : papier, orientation, échelle, marges, centrage,
 * zone d'impression, titres répétés, quadrillage, en-têtes et pieds de page.
 */
import type { FUniver } from '@univerjs/presets';
import { DEFAULT_PAGE, MARGIN_PRESETS, marginPresetOf, PAPER_SIZES, parseRowRange, type Margins, type PageSetup, type PaperSize } from '../data/page-setup';
import { t } from '../i18n';
import { rangeFromA1 } from '../io/xlsx/units';
import { box, checkbox, field, group, openDialog, radio, select, textInput } from '../ui/dialog';
import { getPageSetup, savePageSetup } from './page-setup-store';

const CM_PER_INCH = 2.54;
const cm = (inches: number) => String(Math.round(inches * CM_PER_INCH * 100) / 100).replace('.', ',');
const toInches = (text: string) => Number(text.replace(',', '.')) / CM_PER_INCH;

function numberField(value: string): HTMLInputElement {
  const i = textInput(value);
  i.inputMode = 'decimal';
  i.style.width = '5.5em';
  return i;
}

export async function pageSetupDialog(api: FUniver): Promise<boolean> {
  const page = getPageSetup(api);
  const paper = select(Object.entries(PAPER_SIZES).map(([id, p]) => [id, p.label] as [string, string]), page.paper ?? DEFAULT_PAGE.paper);
  const orientation = select([['portrait', t('print.portrait')], ['landscape', t('print.landscape')]], page.orientation ?? 'portrait');

  const useScale = radio('g-ps-scale', t('pageSetup.scale'), 'scale', page.fitWidth === undefined);
  const scale = numberField(String(page.scale ?? 100));
  scale.style.width = '4.5em';
  const useFit = radio('g-ps-scale', t('pageSetup.fit'), 'fit', page.fitWidth !== undefined);
  const fitWidth = numberField(String(page.fitWidth || 1));
  fitWidth.style.width = '3.5em';
  const scaleRow = document.createElement('div');
  scaleRow.className = 'g-row';
  scaleRow.append(useScale.el, scale, Object.assign(document.createElement('span'), { textContent: '%' }));
  const fitRow = document.createElement('div');
  fitRow.className = 'g-row';
  fitRow.append(useFit.el, fitWidth, Object.assign(document.createElement('span'), { textContent: t('pageSetup.pagesWide') }));

  const margins: Margins = { ...(page.margins ?? MARGIN_PRESETS.normal) };
  const preset = select(
    [['normal', t('pageSetup.marginNormal')], ['narrow', t('pageSetup.marginNarrow')], ['wide', t('pageSetup.marginWide')], ['custom', t('pageSetup.marginCustom')]],
    marginPresetOf(margins),
  );
  const marginKeys: [keyof Margins, string][] = [['top', t('pageSetup.top')], ['bottom', t('pageSetup.bottom')], ['left', t('pageSetup.left')], ['right', t('pageSetup.right')]];
  const marginInputs = marginKeys.map(([key, label]) => {
    const input = numberField(cm(margins[key]));
    input.addEventListener('input', () => {
      const v = toInches(input.value);
      if (Number.isFinite(v) && v >= 0) margins[key] = v;
      preset.value = marginPresetOf(margins);
    });
    return { key, input, el: field(`${label} (cm)`, input) };
  });
  preset.addEventListener('change', () => {
    if (preset.value === 'custom') return;
    Object.assign(margins, MARGIN_PRESETS[preset.value as 'normal' | 'narrow' | 'wide']);
    for (const m of marginInputs) m.input.value = cm(margins[m.key]);
  });
  const marginGrid = document.createElement('div');
  marginGrid.className = 'g-csv-settings';
  marginGrid.append(...marginInputs.map((m) => m.el));

  const centerH = checkbox(t('pageSetup.centerH'), !!page.centerH);
  const centerV = checkbox(t('pageSetup.centerV'), !!page.centerV);
  const area = textInput(page.printArea ?? '', 'A1:F40');
  const titles = textInput(page.titleRows ?? '', '1:1');
  const gridlines = checkbox(t('pageSetup.gridlines'), !!page.gridlines);
  const headings = checkbox(t('pageSetup.headings'), !!page.headings);
  const hfField = (label: string, value: string | undefined) => {
    const input = textInput(value ?? '');
    return { input, el: field(label, input) };
  };
  const header = { left: hfField(t('pageSetup.left2'), page.header?.left), center: hfField(t('pageSetup.center'), page.header?.center), right: hfField(t('pageSetup.right2'), page.header?.right) };
  const footer = { left: hfField(t('pageSetup.left2'), page.footer?.left), center: hfField(t('pageSetup.center'), page.footer?.center), right: hfField(t('pageSetup.right2'), page.footer?.right) };
  const hfGrid = (parts: typeof header) => {
    const g = document.createElement('div');
    g.className = 'g-csv-settings';
    g.append(parts.left.el, parts.center.el, parts.right.el);
    return g;
  };

  return openDialog({
    title: t('pageSetup.title'),
    testId: 'dialog-page-setup',
    wide: true,
    okLabel: t('dialog.ok'),
    body: box(
      group(t('pageSetup.page'), field(t('pageSetup.paper'), paper), field(t('print.orientation'), orientation), scaleRow, fitRow),
      group(t('pageSetup.margins'), field(t('pageSetup.preset'), preset), marginGrid, centerH.el, centerV.el),
      group(t('pageSetup.sheet'), field(t('pageSetup.area'), area), field(t('pageSetup.titles'), titles), gridlines.el, headings.el),
      group(t('pageSetup.header'), hfGrid(header)),
      group(t('pageSetup.footer'), hfGrid(footer)),
      t('pageSetup.codes'),
    ),
    onOk: () => {
      const areaText = area.value.trim();
      if (areaText) {
        try {
          rangeFromA1(areaText.replace(/\$/g, ''));
        } catch {
          return t('pageSetup.badArea');
        }
      }
      const titleText = titles.value.trim();
      if (titleText && !parseRowRange(titleText)) return t('pageSetup.badTitles');
      const next: PageSetup = {};
      if (paper.value !== 'A4') next.paper = paper.value as PaperSize;
      if (orientation.value === 'landscape') next.orientation = 'landscape';
      if (useFit.input.checked) {
        next.fitWidth = Math.max(1, Math.round(Number(fitWidth.value) || 1));
        next.fitHeight = 0;
      } else {
        const s = Math.round(Number(scale.value.replace(',', '.')));
        if (Number.isFinite(s) && s >= 10 && s <= 400 && s !== 100) next.scale = s;
      }
      if (marginPresetOf(margins) !== 'normal') next.margins = { ...margins };
      if (centerH.input.checked) next.centerH = true;
      if (centerV.input.checked) next.centerV = true;
      if (areaText) next.printArea = areaText.replace(/\$/g, '');
      if (titleText) next.titleRows = titleText.replace(/\$/g, '');
      if (gridlines.input.checked) next.gridlines = true;
      if (headings.input.checked) next.headings = true;
      const pick = (p: typeof header) => {
        const out: Record<string, string> = {};
        for (const k of ['left', 'center', 'right'] as const) if (p[k].input.value) out[k] = p[k].input.value;
        return Object.keys(out).length ? out : undefined;
      };
      const h = pick(header);
      const f = pick(footer);
      if (h) next.header = h;
      if (f) next.footer = f;
      savePageSetup(api, next);
    },
  });
}
