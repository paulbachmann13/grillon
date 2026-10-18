/**
 * Le document courant : chargement dans Univer, ouverture/enregistrement, sauvegarde automatique.
 */
import type { FUniver } from '@univerjs/presets';
import { t } from '../i18n';
import { openFile, saveAs, writeToHandle } from '../io/files';
import { isTauri, openInNewWindow, pathOfHandle } from '../io/tauri-files';
import type { IWorkbookData } from '../model/univer-types';
import { installFrenchDisplay } from '../univer/french-display';
import { installFrenchFormulas } from '../univer/french-formulas';
import { installCharts } from '../charts/chart-view';
import { installShapes } from '../charts/shape-view';
import { installSparklines } from '../charts/sparkline-view';
import { installQuickAnalysis } from '../univer/quick-analysis';
import { installSymbols } from '../univer/symbols';
import { installDataRibbon } from '../univer/data-ribbon';
import { installFormulasRibbon } from '../univer/formulas-ribbon';
import { installGridShortcuts } from '../univer/grid-shortcuts';
import { installViewRibbon } from '../univer/view-ribbon';
import { installBasics } from '../univer/basics';
import { installPageBreaks } from '../univer/page-breaks';
import { installSlicers } from '../univer/slicers';
import { installCellFormat } from '../univer/cell-format';
import { installOutlineGutter } from '../univer/outline-gutter';
import { installGridGrowth } from '../univer/grid-growth';
import { installTextOverflow } from '../univer/text-overflow';
import { printDialog } from './print';
import type { ThemeMode } from './theme-dialog';
import { csvImportDialog } from './csv-import';
import { installPivot } from '../univer/pivot-ui';
import { installHomeRibbon } from '../univer/home-ribbon';
import { LOCALE_FR } from '../model/univer-types';
import { clearAutosave, loadAutosave, saveAutosave } from '../storage/autosave';
import { confirmAsync } from './confirm';
import { applyStoredProtection, installProtectionGuard } from '../univer/protection';

/** CommandType.COMMAND = 0 : actions de l'utilisateur (les mutations internes, comme le calcul des formules, sont ignorées). */
const COMMAND = 0;
/** Commandes qui ne modifient pas le contenu du classeur. */
const NON_CONTENT_COMMANDS = /^doc\.|^grillon\.|header-(width|height)|copy|scroll|selection|zoom|find|search|activate|editor|preview|show|hide-panel|open|close/i;

const AUTOSAVE_DELAY_MS = 1500;

export type StatusKind = 'info' | 'success' | 'error';

export interface DocumentEvents {
  onStatus: (message: string, kind?: StatusKind) => void;
  onChange: () => void;
}


export class DocumentController {
  fileName = `${t('app.untitled')}.xlsx`;
  handle: FileSystemFileHandle | undefined;
  dirty = false;
  private loading = false;
  private ribbonReady = false;
  private autosaveTimer: ReturnType<typeof setTimeout> | undefined;

  constructor(
    private readonly api: FUniver,
    private readonly events: DocumentEvents,
  ) {
    api.addEvent(api.Event.CommandExecuted, (e) => {
      if (this.loading || e.type !== COMMAND || NON_CONTENT_COMMANDS.test(e.id)) return;
      this.markDirty();
    });
  }

  get baseName(): string {
    return this.fileName.replace(/\.[^.]+$/, '');
  }

  /** IWorkbookData du classeur affiché. */
  snapshot(): IWorkbookData {
    const wb = this.api.getActiveWorkbook();
    if (!wb) throw new Error('Aucun classeur ouvert');
    return wb.save() as unknown as IWorkbookData;
  }

  /** Remplace le classeur affiché. */
  load(data: Partial<IWorkbookData>, fileName: string, handle?: FileSystemFileHandle, dirty = false): void {
    this.loading = true;
    try {
      const current = this.api.getActiveWorkbook();
      if (current) this.api.disposeUnit(current.getId());
      // Police par défaut de chaque feuille (Univer la lit au niveau de la feuille) : celle du classeur, sinon Calibri 11 comme Excel.
      const defaultStyle = (data as { defaultStyle?: object }).defaultStyle ?? { ff: 'Calibri', fs: 11 };
      const sheets = Object.fromEntries(Object.entries(data.sheets ?? {}).map(([id, sh]) => [id, { defaultStyle, ...sh }]));
      this.api.createWorkbook({ locale: LOCALE_FR as IWorkbookData['locale'], ...data, sheets, name: fileName.replace(/\.[^.]+$/, '') } as never);
      installFrenchDisplay(this.api);
      installFrenchFormulas(this.api);
      installTextOverflow(this.api);
      // Feuilles protégées : la protection est réappliquée une fois le classeur prêt.
      setTimeout(() => void applyStoredProtection(this.api), 0);
      if (!this.ribbonReady) {
        installHomeRibbon(this.api, () => this.markDirtyPublic());
        installDataRibbon(this.api, this);
        installFormulasRibbon(this.api);
        installCharts(this.api);
        installSparklines(this.api);
        installSymbols(this.api);
        installShapes(this.api);
        installQuickAnalysis(this.api);
        installPivot(this.api);
        installSlicers(this.api);
        installGridShortcuts(this.api);
        installViewRibbon(this.api, () => this.print(), () => this.pageSetup(), (mode) => this.changeTheme(mode));
        installGridGrowth(this.api);
        installBasics(this.api);
        installCellFormat(this.api);
        installOutlineGutter(this.api);
        installPageBreaks(this.api, () => this.markDirtyPublic());
        installProtectionGuard(this.api);
        this.ribbonReady = true;
      }
      this.fileName = fileName;
      this.handle = handle;
      this.dirty = dirty;
    } finally {
      // Univer émet des mutations pendant quelques ticks après la création (calcul des formules…).
      setTimeout(() => (this.loading = false), 0);
    }
    this.events.onChange();
  }

  newWorkbook(): void {
    const id = `sheet-${Math.random().toString(36).slice(2, 10)}`;
    this.load(
      { defaultStyle: { ff: 'Calibri', fs: 11 }, sheetOrder: [id], sheets: { [id]: { id, name: t('app.defaultSheet'), cellData: {}, rowCount: 1000, columnCount: 26 } } },
      `${t('app.untitled')}.xlsx`,
    );
    void clearAutosave();
  }

  async confirmDiscard(): Promise<boolean> {
    return !this.dirty || confirmAsync(t('confirm.discard'));
  }

  /** Fenêtre vide et intacte : un fichier ouvert la remplace au lieu d'en ouvrir une autre. */
  isPristine(): boolean {
    return !this.dirty && !this.handle;
  }

  async open(): Promise<void> {
    const file = await openFile('any');
    if (!file) return;
    if (await this.openInOtherWindow(file.handle)) return;
    if (!(await this.confirmDiscard())) return;
    await this.openBytes(file.name, file.bytes, file.handle);
  }

  /**
   * Application de bureau : comme Excel, un fichier s'ouvre dans sa propre fenêtre, sans fermer
   * le classeur en cours. Renvoie false s'il faut l'ouvrir ici (fenêtre vide, navigateur).
   */
  private async openInOtherWindow(handle: FileSystemFileHandle | undefined): Promise<boolean> {
    const path = pathOfHandle(handle);
    if (!path || this.isPristine()) return false;
    await openInNewWindow(path);
    return true;
  }

  /** Ouvre le fichier d'un handle (Finder, fenêtre créée pour lui) ; fenêtre dédiée si celle-ci est utilisée. */
  async openHandle(handle: FileSystemFileHandle, inThisWindow = false): Promise<void> {
    if (!inThisWindow && (await this.openInOtherWindow(handle))) return;
    if (!(await this.confirmDiscard())) return;
    const file = await handle.getFile();
    await this.openBytes(file.name, new Uint8Array(await file.arrayBuffer()), handle);
  }

  /** Fichier > Nouveau : dans l'application de bureau, un classeur vide dans une nouvelle fenêtre. */
  async newInNewWindow(): Promise<boolean> {
    if (!isTauri() || this.isPristine()) return false;
    await openInNewWindow();
    return true;
  }

  async openBytes(name: string, bytes: Uint8Array, handle?: FileSystemFileHandle): Promise<void> {
    this.events.onStatus(t('status.loading', { name }));
    try {
      if (/\.xlsx$/i.test(name)) {
        const { importXlsxAsync } = await import('../io/xlsx/xlsx-client');
        const data = await importXlsxAsync(bytes, { name });
        this.load(data, name, handle);
      } else if (/\.(csv|txt|tsv)$/i.test(name)) {
        const { csvToWorkbook } = await import('../io/csv');
        const choice = await csvImportDialog(name, bytes);
        if (!choice) {
          this.events.onStatus('');
          return;
        }
        // Un CSV ouvert devient un classeur ; « Enregistrer » produira un .xlsx (comme Excel le propose).
        this.load(csvToWorkbook(choice.text, name, choice.options), name.replace(/\.[^.]+$/, '.xlsx'));
      } else {
        this.events.onStatus(t('errors.unsupported', { name }), 'error');
        return;
      }
      this.events.onStatus(t('status.saved'), 'success');
      void this.autosaveNow();
    } catch (e) {
      console.error(e);
      this.events.onStatus(t('errors.open', { message: (e as Error).message }), 'error');
    }
  }

  async exportXlsxBytes(): Promise<Uint8Array> {
    // Conversion dans un Web Worker (ExcelJS pèse ~1 Mo et un gros classeur figerait la page).
    const { exportXlsxAsync } = await import('../io/xlsx/xlsx-client');
    return exportXlsxAsync(this.snapshot());
  }

  async save(): Promise<void> {
    if (!this.handle) return this.saveAs();
    this.events.onStatus(t('status.saving'));
    try {
      const bytes = await this.exportXlsxBytes();
      if (await writeToHandle(this.handle, bytes)) {
        this.dirty = false;
        this.events.onStatus(t('status.savedTo', { name: this.fileName }), 'success');
        this.events.onChange();
        void this.autosaveNow();
      } else {
        await this.saveAs();
      }
    } catch (e) {
      console.error(e);
      this.events.onStatus(t('errors.save', { message: (e as Error).message }), 'error');
    }
  }

  async saveAs(): Promise<void> {
    this.events.onStatus(t('status.saving'));
    try {
      const bytes = await this.exportXlsxBytes();
      const name = /\.xlsx$/i.test(this.fileName) ? this.fileName : `${this.baseName}.xlsx`;
      const handle = await saveAs(bytes, name, 'xlsx');
      if (handle === null) {
        this.events.onStatus('');
        return;
      }
      if (handle) {
        this.handle = handle;
        this.fileName = handle.name;
        this.events.onStatus(t('status.savedTo', { name: handle.name }), 'success');
      } else {
        this.events.onStatus(t('status.downloaded', { name }), 'success');
      }
      this.dirty = false;
      this.events.onChange();
      void this.autosaveNow();
    } catch (e) {
      console.error(e);
      this.events.onStatus(t('errors.save', { message: (e as Error).message }), 'error');
    }
  }

  /** Ajoute le contenu d'un CSV dans une nouvelle feuille du classeur courant. */
  async importCsv(): Promise<void> {
    const file = await openFile('csv');
    if (!file) return;
    const choice = await csvImportDialog(file.name, file.bytes);
    if (!choice) return;
    const { csvToSheet } = await import('../io/csv');
    const data = this.snapshot();
    const used = new Set(Object.values(data.sheets).map((s) => s.name));
    let name = file.name.replace(/\.[^.]+$/, '').slice(0, 28) || 'CSV';
    for (let i = 2; used.has(name); i++) name = `${name.replace(/ \(\d+\)$/, '')} (${i})`;
    const sheet = csvToSheet(choice.text, name, choice.options);
    data.sheets[sheet.id] = sheet;
    data.sheetOrder.push(sheet.id);
    this.load(data, this.fileName, this.handle, true);
    this.api.getActiveWorkbook()?.setActiveSheet(this.api.getActiveWorkbook()!.getSheetBySheetId(sheet.id)!);
    this.markDirty();
  }

  /** Exporte la feuille active en CSV (séparateur « ; », virgule décimale, comme Excel en français). */
  async exportCsv(): Promise<void> {
    const { sheetToCsv } = await import('../io/csv');
    const wb = this.api.getActiveWorkbook();
    const sheet = wb?.getActiveSheet();
    if (!wb || !sheet) return;
    const data = this.snapshot();
    const sheetData = data.sheets[sheet.getSheetId()];
    if (!sheetData) return;
    const csv = sheetToCsv(sheetData, {}, (row, col) => {
      // Texte affiché (dates, monnaies…) quand Univer sait le calculer.
      try {
        const shown = sheet.getRange(row, col).getDisplayValue();
        return shown === '' ? undefined : shown;
      } catch {
        return undefined;
      }
    });
    const name = `${this.baseName} - ${sheet.getSheetName()}.csv`;
    const handle = await saveAs(csv, name, 'csv');
    if (handle !== null) this.events.onStatus(t('status.downloaded', { name }), 'success');
  }

  markDirtyPublic(): void {
    this.markDirty();
  }

  private markDirty(): void {
    if (!this.dirty) {
      this.dirty = true;
      this.events.onChange();
    }
    clearTimeout(this.autosaveTimer);
    this.autosaveTimer = setTimeout(() => void this.autosaveNow(), AUTOSAVE_DELAY_MS);
  }

  async autosaveNow(): Promise<void> {
    clearTimeout(this.autosaveTimer);
    try {
      await saveAutosave({ data: this.snapshot(), fileName: this.fileName, savedAt: Date.now(), dirty: this.dirty, handle: this.handle });
    } catch (e) {
      console.warn('Sauvegarde automatique impossible', e);
    }
  }

  /** Protéger / ôter la protection de la feuille active (aussi dans Accueil → Format). */
  async protectSheet(): Promise<void> {
    const { protectSheetDialog } = await import('../univer/protection');
    await protectSheetDialog(this.api, () => this.markDirtyPublic());
  }

  async unprotectSheet(): Promise<void> {
    const { unprotectSheetDialog } = await import('../univer/protection');
    await unprotectSheetDialog(this.api, () => this.markDirtyPublic());
  }

  /** Boîte « Mise en page » de la feuille active. */
  async pageSetup(): Promise<void> {
    const { pageSetupDialog } = await import('./page-setup');
    if (await pageSetupDialog(this.api)) this.markDirtyPublic();
  }

  /** Mise en page → Thèmes, Couleurs, Polices : recolore le classeur puis le recharge (l'historique d'annulation est perdu). */
  async changeTheme(mode: ThemeMode): Promise<void> {
    const { themeDialog } = await import('./theme-dialog');
    const { applyThemeToWorkbook, themeOf } = await import('../data/theme');
    const data = this.snapshot();
    const next = await themeDialog(mode, themeOf(data as never));
    if (!next) return;
    const wb = this.api.getActiveWorkbook();
    const activeSheetId = wb?.getActiveSheet().getSheetId();
    this.load(applyThemeToWorkbook(data as never, next) as unknown as IWorkbookData, this.fileName, this.handle, true);
    const loaded = this.api.getActiveWorkbook();
    const sheet = activeSheetId ? loaded?.getSheetBySheetId(activeSheetId) : null;
    if (loaded && sheet) loaded.setActiveSheet(sheet);
    this.markDirty();
    this.events.onStatus(t('theme.applied', { name: next.name }), 'success');
  }

  /** Impression / export PDF de la feuille active. */
  print(): Promise<void> {
    return printDialog(this.api, this.fileName);
  }

  /** Restaure la session précédente. Renvoie true si un classeur a été restauré. */
  async restore(): Promise<boolean> {
    const entry = await loadAutosave();
    if (!entry?.data?.sheetOrder?.length) return false;
    try {
      this.load(entry.data, entry.fileName, entry.handle, entry.dirty);
      return true;
    } catch (e) {
      console.warn('Restauration impossible', e);
      return false;
    }
  }

}
