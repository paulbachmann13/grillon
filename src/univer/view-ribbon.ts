/**
 * Onglet Vue : mode sombre (suit le système par défaut, choix mémorisé sur l'appareil)
 * et impression / export PDF.
 */
import { IconManager } from '@univerjs/preset-sheets-core';
import { CommandType, ICommandService, type FUniver } from '@univerjs/presets';
import { createElement, type ComponentType } from 'react';
import { t } from '../i18n';

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

export function installViewRibbon(api: FUniver, print: () => unknown): void {
  const injector = (api as unknown as { _injector: { get<T>(id: unknown): T } })._injector;
  const commands = injector.get<ICommandService>(ICommandService);
  injector.get<IconManager>(IconManager).register({
    GrillonDarkIcon: icon('M13.5 9.5A5.5 5.5 0 016.5 2.5a5.5 5.5 0 107 7z'),
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
  api.createMenu({ id: 'grillon.view.dark', icon: 'GrillonDarkIcon', title: t('view.darkMode'), tooltip: t('view.darkMode'), action: 'grillon.view.dark' }).appendTo(['ribbon.view', 'ribbon.view.display']);
  api.createMenu({ id: 'grillon.view.print', icon: 'GrillonPrintIcon', title: t('view.print'), tooltip: t('view.print'), action: 'grillon.view.print' }).appendTo(['ribbon.view', 'ribbon.view.others']);
}
