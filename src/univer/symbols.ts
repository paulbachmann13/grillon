/**
 * Insertion de symboles (Insérer → Symbole) : une grille de caractères courants par catégorie,
 * ajoutés au texte de la cellule active.
 */
import { IconManager } from '@univerjs/preset-sheets-core';
import { CommandType, ICommandService, type FUniver } from '@univerjs/presets';
import { createElement, type ComponentType } from 'react';
import { t } from '../i18n';
import { box, openDialog, select } from '../ui/dialog';

export const SYMBOL_GROUPS: { id: string; label: string; chars: string }[] = [
  { id: 'common', label: 'Courants', chars: '€£$¥¢©®™°±×÷½¼¾²³µ‰§¶†‡•…«»“”‘’–—·¿¡№' },
  { id: 'currency', label: 'Devises', chars: '€£$¥¢₣₤₱₹₽₩₪₫₺₿฿₡₦₵' },
  { id: 'math', label: 'Mathématiques', chars: '±×÷≠≈≤≥∞√∑∏∫∂∆∇∈∉⊂⊃∪∩∧∨¬∀∃∅‰°′″⁰¹²³⁴⁵⁶⁷⁸⁹½⅓¼¾' },
  { id: 'arrows', label: 'Flèches', chars: '←↑→↓↔↕↖↗↘↙⇐⇑⇒⇓⇔⇕↩↪↺↻▲▼◀▶△▽◁▷' },
  { id: 'greek', label: 'Lettres grecques', chars: 'αβγδεζηθικλμνξοπρστυφχψωΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡΣΤΥΦΧΨΩ' },
  { id: 'latin', label: 'Lettres accentuées', chars: 'àâäæçéèêëîïôöœùûüÿÀÂÄÆÇÉÈÊËÎÏÔÖŒÙÛÜŸñÑßøØåÅ' },
  { id: 'shapes', label: 'Formes et signes', chars: '✓✔✗✘★☆♥♦♣♠☺☹☀☁☂☎✉✂✈♪♫⚠⚡●○■□◆◇◉✪✿❀' },
];

function insertText(api: FUniver, text: string): void {
  const wb = api.getActiveWorkbook();
  const range = wb?.getActiveRange();
  const ws = wb?.getActiveSheet();
  if (!wb || !range || !ws) return;
  // On travaille sur la cellule active : le symbole est ajouté à son texte (les formules sont complétées).
  const cell = ws.getRange(range.getRow(), range.getColumn());
  const raw = cell.getCellDataGrid()[0]?.[0];
  const formula = cell.getFormula();
  if (formula) cell.setValue(`${formula}${text}`);
  else cell.setValue(`${raw?.v === undefined || raw.v === null ? '' : String(raw.v)}${text}`);
}

export async function symbolDialog(api: FUniver): Promise<void> {
  const group = select(SYMBOL_GROUPS.map((g) => [g.id, g.label] as [string, string]), 'common');
  const grid = document.createElement('div');
  grid.className = 'g-symbols';
  grid.dataset.testid = 'symbols';
  const draw = () => {
    const chars = [...(SYMBOL_GROUPS.find((g) => g.id === group.value)?.chars ?? '')];
    grid.replaceChildren(
      ...chars.map((ch) => {
        const b = document.createElement('button');
        b.type = 'button';
        b.textContent = ch;
        b.title = `U+${ch.codePointAt(0)!.toString(16).toUpperCase().padStart(4, '0')}`;
        b.addEventListener('click', () => insertText(api, ch));
        return b;
      }),
    );
  };
  group.addEventListener('change', draw);
  draw();
  const help = document.createElement('p');
  help.className = 'g-muted';
  help.textContent = t('symbols.help');
  await openDialog({
    title: t('symbols.title'),
    testId: 'dialog-symbols',
    body: box(group, grid, help),
    okLabel: t('dialog.close'),
    cancelLabel: t('dialog.cancel'),
    onOk: () => true,
  });
}

function icon(path: string): ComponentType<Record<string, unknown>> {
  return (props) =>
    createElement('svg', { viewBox: '0 0 16 16', width: '1em', height: '1em', fill: 'none', stroke: 'currentColor', strokeWidth: 1.3, strokeLinecap: 'round', strokeLinejoin: 'round', ...props }, createElement('path', { d: path }));
}

export function installSymbols(api: FUniver): void {
  const injector = (api as unknown as { _injector: { get<T>(id: unknown): T } })._injector;
  const commands = injector.get<ICommandService>(ICommandService);
  injector.get<IconManager>(IconManager).register({ GrillonSymbolIcon: icon('M3 13h3v-1.5C4 10.5 3 9 3 7a5 5 0 0110 0c0 2-1 3.5-3 4.5V13h3') } as never);
  const id = 'grillon.insert.symbol';
  if (!commands.hasCommand(id)) commands.registerCommand({ id, type: CommandType.COMMAND, handler: () => (void symbolDialog(api), true) });
  api.createMenu({ id, icon: 'GrillonSymbolIcon', title: t('symbols.title'), tooltip: t('symbols.title'), action: id }).appendTo(['ribbon.insert', 'ribbon.insert.media']);
}
