/**
 * Barre supérieure : bouton « Fichier » (comme l'onglet Fichier d'Excel), nom du document, état.
 */
import { t } from '../i18n';
import { isTauri } from '../io/tauri-files';
import type { DocumentController, StatusKind } from './document';

interface MenuItem {
  id: string;
  label: string;
  shortcut?: string;
  run: () => unknown;
  separatorBefore?: boolean;
}

const isMac = /Mac|iPhone|iPad/.test(navigator.platform) || /Mac OS/.test(navigator.userAgent);
const mod = isMac ? '⌘' : 'Ctrl+';

export class Topbar {
  private readonly title: HTMLElement;
  private readonly status: HTMLElement;
  private readonly menu: HTMLElement;
  private readonly fileButton: HTMLButtonElement;
  private statusTimer: ReturnType<typeof setTimeout> | undefined;

  constructor(
    root: HTMLElement,
    private readonly doc: DocumentController,
  ) {
    root.innerHTML = '';
    root.classList.add('topbar');

    this.fileButton = document.createElement('button');
    this.fileButton.className = 'topbar-file';
    this.fileButton.type = 'button';
    this.fileButton.textContent = t('file.menu');
    this.fileButton.setAttribute('aria-haspopup', 'menu');
    this.fileButton.setAttribute('aria-expanded', 'false');
    this.fileButton.dataset.testid = 'file-menu';

    const brand = document.createElement('span');
    brand.className = 'topbar-brand';
    brand.textContent = t('app.name');

    this.title = document.createElement('span');
    this.title.className = 'topbar-title';
    this.title.dataset.testid = 'doc-title';

    this.status = document.createElement('span');
    this.status.className = 'topbar-status';
    this.status.setAttribute('role', 'status');
    this.status.dataset.testid = 'status';

    this.menu = document.createElement('div');
    this.menu.className = 'file-menu';
    this.menu.setAttribute('role', 'menu');
    this.menu.hidden = true;

    const items: MenuItem[] = [
      { id: 'new', label: t('file.new'), run: async () => ((await doc.newInNewWindow()) ? undefined : this.confirmThen(() => doc.newWorkbook())) },
      { id: 'open', label: t('file.open'), shortcut: `${mod}O`, run: () => doc.open() },
      { id: 'save', label: t('file.save'), shortcut: `${mod}S`, run: () => doc.save(), separatorBefore: true },
      { id: 'save-as', label: t('file.saveAs'), shortcut: isMac ? '⇧⌘S' : 'Ctrl+Maj+S', run: () => doc.saveAs() },
      { id: 'import-csv', label: t('file.importCsv'), run: () => doc.importCsv(), separatorBefore: true },
      { id: 'export-csv', label: t('file.exportCsv'), run: () => doc.exportCsv() },
      { id: 'protect', label: t('protection.menuProtect'), run: () => doc.protectSheet(), separatorBefore: true },
      { id: 'unprotect', label: t('protection.menuUnprotect'), run: () => doc.unprotectSheet() },
      { id: 'page-setup', label: t('file.pageSetup'), run: () => doc.pageSetup(), separatorBefore: true },
      { id: 'print', label: t('file.print'), shortcut: `${mod}P`, run: () => doc.print() },
      ...(isTauri() ? [{ id: 'update', label: t('update.check'), run: () => doc.checkForUpdates(false), separatorBefore: true }] : []),
    ];
    for (const item of items) {
      if (item.separatorBefore) {
        const sep = document.createElement('div');
        sep.className = 'file-menu-sep';
        this.menu.appendChild(sep);
      }
      const b = document.createElement('button');
      b.type = 'button';
      b.setAttribute('role', 'menuitem');
      b.dataset.testid = `file-${item.id}`;
      b.innerHTML = `<span></span><kbd></kbd>`;
      b.querySelector('span')!.textContent = item.label;
      b.querySelector('kbd')!.textContent = item.shortcut ?? '';
      b.addEventListener('click', () => {
        this.closeMenu();
        void item.run();
      });
      this.menu.appendChild(b);
    }

    this.fileButton.addEventListener('click', (e) => {
      e.stopPropagation();
      if (this.menu.hidden) this.openMenu();
      else this.closeMenu();
    });
    document.addEventListener('pointerdown', (e) => {
      if (!this.menu.hidden && !this.menu.contains(e.target as Node) && e.target !== this.fileButton) this.closeMenu();
    });
    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape' && !this.menu.hidden) this.closeMenu();
    });

    root.append(this.fileButton, brand, this.title, this.status, this.menu);
    this.refresh();
  }

  private confirmThen(fn: () => void): void {
    void this.doc.confirmDiscard().then((ok) => ok && fn());
  }

  private openMenu(): void {
    this.menu.hidden = false;
    this.fileButton.setAttribute('aria-expanded', 'true');
    (this.menu.querySelector('button') as HTMLButtonElement | null)?.focus();
  }

  private closeMenu(): void {
    this.menu.hidden = true;
    this.fileButton.setAttribute('aria-expanded', 'false');
  }

  refresh(): void {
    this.title.textContent = `${this.doc.baseName}${this.doc.dirty ? ' •' : ''}`;
    this.title.title = this.doc.dirty ? t('status.unsaved') : this.doc.fileName;
    document.title = `${this.doc.baseName}${this.doc.dirty ? ' •' : ''} — ${t('app.name')}`;
  }

  /** Message avec action (ex. « Nouvelle version disponible — recharger »). */
  setAction(message: string, onClick: () => void): void {
    clearTimeout(this.statusTimer);
    this.status.textContent = '';
    const b = document.createElement('button');
    b.type = 'button';
    b.className = 'topbar-action';
    b.textContent = message;
    b.addEventListener('click', onClick);
    this.status.appendChild(b);
  }

  setStatus(message: string, kind: StatusKind = 'info'): void {
    clearTimeout(this.statusTimer);
    this.status.textContent = message;
    this.status.dataset.kind = kind;
    if (message && kind !== 'error') this.statusTimer = setTimeout(() => (this.status.textContent = ''), 4000);
  }
}
