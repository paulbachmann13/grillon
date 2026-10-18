/**
 * Écran d'accueil, comme celui d'Excel : Nouveau, Ouvrir, fichiers récents, Paramètres et numéro de version.
 * Il apparaît au démarrage quand il n'y a rien à reprendre, et à la demande (Fichier → Accueil).
 */
import { folderOf, formatWhen, type RecentFile } from '../data/recent';
import { t } from '../i18n';
import { isTauri } from '../io/tauri-files';
import { forgetRecent, loadRecent } from '../storage/recent';
import { setThemeChoice, themeChoice, type ThemeChoice } from '../univer/view-ribbon';
import { setTextOverflow, textOverflowEnabled } from '../univer/text-overflow';
import type { FUniver } from '@univerjs/presets';
import { autoUpdateEnabled, setAutoUpdate } from './settings';
import { appVersion } from './version';

export type StartPane = 'home' | 'new' | 'open' | 'settings';

export interface StartActions {
  newBlank(): void;
  openFile(): void | Promise<void>;
  importCsv(): void | Promise<void>;
  openRecent(file: RecentFile): void | Promise<void>;
  checkUpdates(): void | Promise<void>;
  /** Travail non enregistré de la dernière session (nom, date de la dernière sauvegarde automatique). */
  draft(): Promise<{ name: string; savedAt: number } | null>;
  resumeDraft(): void | Promise<void>;
  /** Renvoie vrai si le brouillon a bien été supprimé (l'utilisateur confirme). */
  discardDraft(): Promise<boolean>;
}

const el = <K extends keyof HTMLElementTagNameMap>(tag: K, className?: string, text?: string): HTMLElementTagNameMap[K] => {
  const e = document.createElement(tag);
  if (className) e.className = className;
  if (text !== undefined) e.textContent = text;
  return e;
};

const PLUS = '<svg viewBox="0 0 24 24" width="34" height="34" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"><path d="M12 5v14M5 12h14"/></svg>';
const FOLDER = '<svg viewBox="0 0 24 24" width="34" height="34" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><path d="M3 7a2 2 0 012-2h4l2 2h8a2 2 0 012 2v8a2 2 0 01-2 2H5a2 2 0 01-2-2z"/></svg>';
const TABLE = '<svg viewBox="0 0 24 24" width="34" height="34" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><path d="M4 4h16v16H4zM4 10h16M10 4v16"/></svg>';
const RESUME = '<svg viewBox="0 0 24 24" width="22" height="22" fill="none" stroke="#b54708" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"><path d="M4 12a8 8 0 108-8M4 4v5h5"/></svg>';
const FILE = '<svg viewBox="0 0 24 24" width="22" height="22" fill="none" stroke="#1f7a4d" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><path d="M6 3h8l4 4v14H6zM14 3v4h4M9 12h6M9 16h6"/></svg>';

export class StartScreen {
  readonly root = el('div', 'g-start');
  private readonly nav = el('nav', 'g-start-nav');
  private readonly main = el('section', 'g-start-main');
  private pane: StartPane = 'home';
  /** Un classeur est ouvert derrière l'écran : on propose d'y revenir. */
  private returnable = false;
  /** Dossier personnel (application de bureau), pour raccourcir les chemins en « ~/… ». */
  home: string | undefined;

  constructor(
    private readonly api: FUniver,
    private readonly actions: StartActions,
  ) {
    this.root.dataset.testid = 'start-screen';
    this.root.setAttribute('role', 'dialog');
    this.root.setAttribute('aria-label', t('home.title'));
    this.root.hidden = true;
    this.root.append(this.nav, this.main);
    document.getElementById('app')?.append(this.root);
    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape' && !this.root.hidden && this.returnable) this.hide();
    });
  }

  get visible(): boolean {
    return !this.root.hidden;
  }

  show(pane: StartPane = 'home', returnable = false): void {
    this.pane = pane;
    this.returnable = returnable;
    this.root.hidden = false;
    this.render();
  }

  hide(): void {
    this.root.hidden = true;
  }

  private go(pane: StartPane): void {
    this.pane = pane;
    this.render();
  }

  private render(): void {
    this.renderNav();
    this.main.replaceChildren();
    if (this.returnable) {
      const back = el('button', 'g-start-back', `← ${t('home.back')}`);
      back.type = 'button';
      back.dataset.testid = 'start-back';
      back.addEventListener('click', () => this.hide());
      this.main.append(back);
    }
    if (this.pane === 'settings') this.main.append(this.settingsPane());
    else if (this.pane === 'new') this.main.append(el('h1', undefined, t('home.nav.new')), this.tiles(false));
    else if (this.pane === 'open') this.main.append(el('h1', undefined, t('home.nav.open')), this.openTile(), this.recentList());
    else this.main.append(el('h1', undefined, t('home.hello')), el('p', 'g-start-sub', t('home.tagline')), this.tiles(true), this.recentList());
    (this.main.querySelector('[data-autofocus]') as HTMLElement | null)?.focus();
  }

  private renderNav(): void {
    this.nav.replaceChildren();
    const brand = el('div', 'g-start-brand');
    const logo = el('img');
    logo.src = `${import.meta.env.BASE_URL}icon.svg`;
    logo.alt = '';
    logo.width = 34;
    logo.height = 34;
    brand.append(logo, el('span', undefined, 'Grillon'));
    this.nav.append(brand);
    const item = (pane: StartPane, label: string) => {
      const b = el('button', pane === this.pane ? 'g-start-nav-item active' : 'g-start-nav-item', label);
      b.type = 'button';
      b.dataset.testid = `start-nav-${pane}`;
      b.addEventListener('click', () => this.go(pane));
      return b;
    };
    this.nav.append(item('home', t('home.nav.home')), item('new', t('home.nav.new')), item('open', t('home.nav.open')));
    this.nav.append(el('div', 'g-start-spacer'), item('settings', t('home.nav.settings')));
    const version = el('div', 'g-start-version', t('home.version', { version: appVersion() }));
    version.dataset.testid = 'start-version';
    this.nav.append(version);
  }

  private tile(icon: string, title: string, hint: string, testId: string, run: () => void | Promise<void>, focus = false): HTMLButtonElement {
    const b = el('button', 'g-start-tile');
    b.type = 'button';
    b.dataset.testid = testId;
    if (focus) b.dataset.autofocus = '';
    const i = el('span', 'g-start-tile-icon');
    i.innerHTML = icon;
    const text = el('span', 'g-start-tile-text');
    text.append(el('strong', undefined, title), el('small', undefined, hint));
    b.append(i, text);
    b.addEventListener('click', () => void run());
    return b;
  }

  private tiles(withOpen: boolean): HTMLElement {
    const row = el('div', 'g-start-tiles');
    row.append(
      this.tile(PLUS, t('home.newBlank'), t('home.newBlankHint'), 'start-new', () => this.actions.newBlank(), true),
      ...(withOpen ? [this.tile(FOLDER, t('home.openFile'), t('home.openFileHint'), 'start-open', () => this.actions.openFile())] : []),
      this.tile(TABLE, t('home.importCsv'), t('home.importCsvHint'), 'start-csv', () => this.actions.importCsv()),
    );
    return row;
  }

  private openTile(): HTMLElement {
    const row = el('div', 'g-start-tiles');
    row.append(this.tile(FOLDER, t('home.browse'), t('home.openFileHint'), 'start-browse', () => this.actions.openFile(), true));
    return row;
  }

  private recentList(): HTMLElement {
    const box = el('div', 'g-start-recent');
    box.append(el('h2', undefined, t('home.recent')));
    const drafts = el('div', 'g-start-drafts');
    box.append(drafts);
    void this.actions.draft().then((draft) => {
      if (!draft) return;
      const row = el('div', 'g-start-recent-row g-start-draft');
      row.dataset.testid = 'start-draft';
      const open = el('button', 'g-start-recent-open');
      open.type = 'button';
      const icon = el('span', 'g-start-recent-icon');
      icon.innerHTML = RESUME;
      const text = el('span', 'g-start-recent-text');
      text.append(el('strong', undefined, t('home.resume', { name: draft.name })), el('small', undefined, t('home.resumeHint')));
      open.append(icon, text, el('span', 'g-start-recent-when', formatWhen(draft.savedAt)));
      open.addEventListener('click', () => void this.actions.resumeDraft());
      const discard = el('button', 'g-start-recent-forget', '✕');
      discard.type = 'button';
      discard.title = t('home.discard');
      discard.setAttribute('aria-label', t('home.discard'));
      discard.dataset.testid = 'start-discard-draft';
      discard.addEventListener('click', async () => {
        if (await this.actions.discardDraft()) row.remove();
      });
      row.append(open, discard);
      drafts.append(row);
    });
    const list = loadRecent();
    if (!list.length) {
      const empty = el('p', 'g-start-empty', t('home.noRecent'));
      empty.dataset.testid = 'start-no-recent';
      box.append(empty);
      return box;
    }
    for (const file of list) {
      const row = el('div', 'g-start-recent-row');
      row.dataset.testid = 'start-recent';
      const open = el('button', 'g-start-recent-open');
      open.type = 'button';
      const icon = el('span', 'g-start-recent-icon');
      icon.innerHTML = FILE;
      const text = el('span', 'g-start-recent-text');
      text.append(el('strong', undefined, file.name), el('small', undefined, folderOf(file.path, this.home)));
      open.append(icon, text, el('span', 'g-start-recent-when', formatWhen(file.at)));
      open.addEventListener('click', () => void this.actions.openRecent(file));
      const forget = el('button', 'g-start-recent-forget', '✕');
      forget.type = 'button';
      forget.title = t('home.forget');
      forget.setAttribute('aria-label', t('home.forget'));
      forget.dataset.testid = 'start-forget';
      forget.addEventListener('click', async () => {
        await forgetRecent(file.id);
        this.render();
      });
      row.append(open, forget);
      box.append(row);
    }
    return box;
  }

  private settingsPane(): HTMLElement {
    const pane = el('div', 'g-start-settings');
    pane.append(el('h1', undefined, t('home.settings.title')));

    const appearance = el('fieldset');
    appearance.append(el('legend', undefined, t('home.settings.appearance')));
    const current = themeChoice();
    for (const [value, label] of [['system', t('home.settings.auto')], ['light', t('home.settings.light')], ['dark', t('home.settings.dark')]] as [ThemeChoice, string][]) {
      const l = el('label', 'g-check');
      const input = el('input');
      input.type = 'radio';
      input.name = 'g-theme';
      input.value = value;
      input.checked = current === value;
      input.dataset.testid = `setting-theme-${value}`;
      input.addEventListener('change', () => setThemeChoice(this.api, value));
      l.append(input, document.createTextNode(label));
      appearance.append(l);
    }
    const overflow = el('label', 'g-check');
    const overflowInput = el('input');
    overflowInput.type = 'checkbox';
    overflowInput.checked = textOverflowEnabled();
    overflowInput.dataset.testid = 'setting-overflow';
    overflowInput.addEventListener('change', () => setTextOverflow(this.api, overflowInput.checked));
    overflow.append(overflowInput, document.createTextNode(t('home.settings.overflow')));
    pane.append(appearance, overflow);

    if (isTauri()) {
      const updates = el('fieldset');
      updates.append(el('legend', undefined, t('home.settings.updates')));
      const auto = el('label', 'g-check');
      const autoInput = el('input');
      autoInput.type = 'checkbox';
      autoInput.checked = autoUpdateEnabled();
      autoInput.dataset.testid = 'setting-auto-update';
      autoInput.addEventListener('change', () => setAutoUpdate(autoInput.checked));
      auto.append(autoInput, document.createTextNode(t('home.settings.autoUpdate')));
      const now = el('button', 'g-start-button', t('home.settings.checkNow'));
      now.type = 'button';
      now.dataset.testid = 'setting-check-updates';
      now.addEventListener('click', () => void this.actions.checkUpdates());
      updates.append(auto, now);
      pane.append(updates);
    }

    const about = el('div', 'g-start-about');
    about.append(el('h2', undefined, t('home.settings.about')), el('p', undefined, t('home.version', { version: appVersion() })), el('p', 'g-start-sub', t('home.settings.aboutText')));
    pane.append(about);
    return pane;
  }
}
