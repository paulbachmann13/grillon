/**
 * Onglet Vue : mode sombre (suit le système par défaut, choix mémorisé sur l'appareil)
 * et impression / export PDF.
 */
import { IconManager } from '@univerjs/preset-sheets-core';
import { CommandType, ICommandService, type FUniver } from '@univerjs/presets';
import { createElement, type ComponentType } from 'react';
import { t } from '../i18n';
import { insertPageBreak, pageBreakPreviewEnabled, removePageBreak, resetPageBreaks, setPageBreakPreview } from './page-breaks';
import { setTextOverflow, textOverflowEnabled } from './text-overflow';

type ThemeChoice = 'system' | 'light' | 'dark';
const STORAGE_KEY = 'grillon-theme';

function readChoice(): ThemeChoice {
  try {
    const v = localStorage.getItem(STORAGE_KEY);
    return v === 'light' || v === 'dark' ? v : 'system';
  } catch {
    return 'system';
  }
}

function saveChoice(choice: ThemeChoice): void {
  try {
    if (choice === 'system') localStorage.removeItem(STORAGE_KEY);
    else localStorage.setItem(STORAGE_KEY, choice);
  } catch {
    /* stockage indisponible : le choix vaut pour la session */
  }
}

/** Choix d'apparence (Automatique, Clair, Sombre) : lu et changé aussi depuis l'écran Paramètres. */
export type { ThemeChoice };
export const themeChoice = (): ThemeChoice => readChoice();
export function setThemeChoice(api: FUniver, choice: ThemeChoice): void {
  saveChoice(choice);
  applyTheme(api, choice);
}

const media = () => matchMedia('(prefers-color-scheme: dark)');

export function applyTheme(api: FUniver, choice: ThemeChoice = readChoice()): void {
  const dark = choice === 'dark' || (choice === 'system' && media().matches);
  if (choice === 'system') delete document.documentElement.dataset.theme;
  else document.documentElement.dataset.theme = choice;
  (api as unknown as { toggleDarkMode(d: boolean): void }).toggleDarkMode(dark);
  document.querySelector('meta[name="theme-color"]')?.setAttribute('content', dark ? '#1b1f23' : '#1f7a4d');
}

function icon(path: string): ComponentType<Record<string, unknown>> {
  return (props) =>
    createElement('svg', { viewBox: '0 0 16 16', width: '1em', height: '1em', fill: 'none', stroke: 'currentColor', strokeWidth: 1.3, strokeLinecap: 'round', strokeLinejoin: 'round', ...props }, createElement('path', { d: path }));
}

export function installViewRibbon(api: FUniver, print: () => unknown, pageSetup: () => unknown, changeTheme: (mode: 'theme' | 'colors' | 'fonts') => unknown): void {
  const injector = (api as unknown as { _injector: { get<T>(id: unknown): T } })._injector;
  const commands = injector.get<ICommandService>(ICommandService);
  injector.get<IconManager>(IconManager).register({
    GrillonDarkIcon: icon('M13.5 9.5A5.5 5.5 0 016.5 2.5a5.5 5.5 0 107 7z'),
    GrillonOverflowIcon: icon('M2 3v10M2 8h9M8 5l3 3-3 3M14 3v10'),
    GrillonPageIcon: icon('M4 2h6l3 3v9H4zM10 2v3h3M6.5 8h4M6.5 10.5h4'),
    GrillonThemeIcon: icon('M8 2a6 6 0 100 12c1 0 1.5-.7 1.2-1.5-.3-.8.2-1.5 1-1.5H12a2 2 0 002-2A6 6 0 008 2zM5 7.5h.01M7 5h.01M10 5.5h.01'),
    GrillonColorsIcon: icon('M2 12V6h3v6zM6.5 12V3h3v9zM11 12V8h3v4z'),
    GrillonFontsIcon: icon('M3 13L7 3l4 10M4.5 9.5h5M12 13V8M12 8l1.5 0'),
    GrillonBreaksIcon: icon('M3 2.5h10M3 13.5h10M2 8h2M6 8h4M12 8h2M5 5.5h6M5 10.5h6'),
    GrillonBreakPreviewIcon: icon('M2.5 2.5h11v11h-11zM2.5 8h11M8 2.5v11'),
    GrillonPrintIcon: icon('M4 6V2h8v4M4 12H2.5V7h11v5H12M4 10h8v4H4z'),
  } as never);
  applyTheme(api);
  media().addEventListener('change', () => readChoice() === 'system' && applyTheme(api));

  const register = (id: string, run: () => unknown) => {
    if (!commands.hasCommand(id)) commands.registerCommand({ id, type: CommandType.COMMAND, handler: () => (void run(), true) });
  };
  register('grillon.view.dark', () => {
    const dark = document.documentElement.dataset.theme === 'dark' || (!document.documentElement.dataset.theme && media().matches);
    const next: ThemeChoice = dark ? 'light' : 'dark';
    saveChoice(next);
    applyTheme(api, next);
  });
  register('grillon.view.print', print);
  register('grillon.view.break-insert', () => insertPageBreak(api));
  register('grillon.view.break-remove', () => removePageBreak(api));
  register('grillon.view.break-reset', () => resetPageBreaks(api));
  register('grillon.view.break-preview', () => setPageBreakPreview(!pageBreakPreviewEnabled()));
  for (const mode of ['theme', 'colors', 'fonts'] as const) register(`grillon.view.theme-${mode}`, () => changeTheme(mode));
  register('grillon.view.page-setup', pageSetup);
  register('grillon.view.text-overflow', () => setTextOverflow(api, !textOverflowEnabled()));
  api.createMenu({ id: 'grillon.view.dark', icon: 'GrillonDarkIcon', title: t('view.darkMode'), tooltip: t('view.darkMode'), action: 'grillon.view.dark' }).appendTo(['ribbon.view', 'ribbon.view.display']);
  api.createMenu({ id: 'grillon.view.text-overflow', icon: 'GrillonOverflowIcon', title: t('view.textOverflow'), tooltip: t('view.textOverflowTip'), action: 'grillon.view.text-overflow' }).appendTo(['ribbon.view', 'ribbon.view.display']);
  for (const [mode, iconName, label] of [['theme', 'GrillonThemeIcon', t('theme.themes')], ['colors', 'GrillonColorsIcon', t('theme.colors')], ['fonts', 'GrillonFontsIcon', t('theme.fonts')]] as const) {
    const id = `grillon.view.theme-${mode}`;
    api.createMenu({ id, icon: iconName, title: label, tooltip: label, action: id }).appendTo(['ribbon.view', 'ribbon.view.others']);
  }
  api.createMenu({ id: 'grillon.view.break-preview', icon: 'GrillonBreakPreviewIcon', title: t('pageBreaks.preview'), tooltip: t('pageBreaks.preview'), action: 'grillon.view.break-preview' }).appendTo(['ribbon.view', 'ribbon.view.display']);
  api.createSubmenu({ id: 'grillon.view.breaks', icon: 'GrillonBreaksIcon', title: t('pageBreaks.menu'), tooltip: t('pageBreaks.menu') }).appendTo(['ribbon.view', 'ribbon.view.others']);
  for (const [k, order] of [['insert', 0], ['remove', 1], ['reset', 2]] as const) {
    const id = `grillon.view.break-${k}`;
    api.createMenu({ id, title: t(`pageBreaks.${k}`), action: id, order }).appendTo(['ribbon.view', 'ribbon.view.others', 'grillon.view.breaks']);
  }
  api.createMenu({ id: 'grillon.view.page-setup', icon: 'GrillonPageIcon', title: t('view.pageSetup'), tooltip: t('view.pageSetup'), action: 'grillon.view.page-setup' }).appendTo(['ribbon.view', 'ribbon.view.others']);
  api.createMenu({ id: 'grillon.view.print', icon: 'GrillonPrintIcon', title: t('view.print'), tooltip: t('view.print'), action: 'grillon.view.print' }).appendTo(['ribbon.view', 'ribbon.view.others']);
}
