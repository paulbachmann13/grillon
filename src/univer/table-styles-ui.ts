/**
 * Galerie des styles de tableau (Mettre sous forme de tableau, Style du tableau), comme Excel.
 * Les styles sont enregistrés comme thèmes de plage Univer sous leur nom Excel.
 */
import { RangeThemeStyle, SheetRangeThemeModel } from '@univerjs/preset-sheets-core';
import type { FUniver } from '@univerjs/presets';
import { TABLE_STYLES, type TableStyleItem, type TableStyleJson } from '../data/table-styles';
import { t } from '../i18n';
import { box, checkbox, openDialog } from '../ui/dialog';

let registered = false;

/** Enregistre les styles d'Excel auprès d'Univer (une fois). */
export function registerTableStyles(api: FUniver): void {
  if (registered) return;
  const injector = (api as unknown as { _injector: { get<T>(id: unknown): T } })._injector;
  const model = injector.get<SheetRangeThemeModel>(SheetRangeThemeModel);
  for (const s of TABLE_STYLES) model.registerDefaultRangeTheme(new RangeThemeStyle(s.name, s.style as never));
  registered = true;
}

const css = (item: TableStyleItem | undefined) => {
  if (!item) return '';
  const parts: string[] = [];
  if (item.bg) parts.push(`background:${item.bg.rgb}`);
  if (item.cl) parts.push(`color:${item.cl.rgb}`);
  for (const [k, side] of [['t', 'top'], ['b', 'bottom']] as const) if (item.bd?.[k]) parts.push(`border-${side}:1px solid ${item.bd[k]!.cl.rgb}`);
  return parts.join(';');
};

/** Aperçu miniature : une ligne d'en-tête et quatre lignes à bandes. */
function preview(style: TableStyleJson): HTMLElement {
  const grid = document.createElement('div');
  grid.className = 'g-table-preview';
  for (let r = 0; r < 5; r++) {
    const item = r === 0 ? style.headerRowStyle : r % 2 === 1 ? style.firstRowStyle : style.secondRowStyle;
    const last = r === 4 ? css(style.lastRowStyle) : '';
    for (let c = 0; c < 4; c++) {
      const cell = document.createElement('span');
      cell.setAttribute('style', [css(style.wholeStyle), css(item), last].filter(Boolean).join(';'));
      cell.textContent = r === 0 ? '▬' : '';
      grid.append(cell);
    }
  }
  return grid;
}

/** Ouvre la galerie ; renvoie le nom Excel du style choisi (et l'option « en-têtes »), ou null. */
export async function pickTableStyle(title: string, askHeaders: boolean, current?: string): Promise<{ name: string; headers: boolean } | null> {
  let chosen: string | null = null;
  const headers = checkbox(t('data.tableHasHeaders'), true);
  const body = document.createElement('div');
  body.className = 'g-cell-styles';
  body.dataset.testid = 'table-styles';
  for (const family of ['light', 'medium', 'dark'] as const) {
    const h = document.createElement('h3');
    h.textContent = t(`data.tableFamily.${family}`);
    const row = document.createElement('div');
    row.className = 'g-table-style-row';
    for (const entry of TABLE_STYLES.filter((s) => s.family === family)) {
      const b = document.createElement('button');
      b.type = 'button';
      b.className = 'g-table-style';
      b.title = entry.label;
      b.setAttribute('aria-label', entry.label);
      b.classList.toggle('selected', entry.name === current);
      b.append(preview(entry.style));
      b.addEventListener('click', () => {
        chosen = entry.name;
        (b.closest('.g-dialog')?.querySelector('.g-dialog-actions button:not(.primary)') as HTMLButtonElement | null)?.click();
      });
      row.append(b);
    }
    body.append(h, row);
  }
  await openDialog({
    title,
    testId: 'dialog-table-styles',
    wide: true,
    body: askHeaders ? box(headers.el, body) : box(body),
    okLabel: t('dialog.close'),
    onOk: () => true,
  });
  return chosen ? { name: chosen, headers: headers.input.checked } : null;
}
