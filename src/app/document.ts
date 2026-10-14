/**
 * Le document courant : chargement dans Univer, ouverture/enregistrement, sauvegarde automatique.
 */
import type { FUniver } from '@univerjs/presets';
import { t } from '../i18n';
import { openFile, saveAs, writeToHandle } from '../io/files';
import type { IWorkbookData } from '../model/univer-types';
import { installFrenchDisplay } from '../univer/french-display';
import { installHomeRibbon } from '../univer/home-ribbon';
import { LOCALE_FR } from '../model/univer-types';
import { clearAutosave, loadAutosave, saveAutosave } from '../storage/autosave';

/** CommandType.COMMAND = 0 : actions de l'utilisateur (les mutations internes, comme le calcul des formules, sont ignorées). */
const COMMAND = 0;
/** Commandes qui ne modifient pas le contenu du classeur. */
const NON_CONTENT_COMMANDS = /^doc\.|copy|scroll|selection|zoom|find|search|activate|editor|preview|show|hide-panel|open|close/i;

const AUTOSAVE_DELAY_MS = 1500;

export type StatusKind = 'info' | 'success' | 'error';

export interface DocumentEvents {
  onStatus: (message: string, kind?: StatusKind) => void;
  onChange: () => void;
}

/** Chargement paresseux du convertisseur (ExcelJS pèse ~1 Mo). */
const loadXlsx = () => Promise.all([import('../io/xlsx/import'), import('../io/xlsx/export')]);

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
      this.api.createWorkbook({ locale: LOCALE_FR as IWorkbookData['locale'], ...data, name: fileName.replace(/\.[^.]+$/, '') });
      installFrenchDisplay(this.api);
      if (!this.ribbonReady) {
        installHomeRibbon(this.api);
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

  private confirmDiscard(): boolean {
    return !this.dirty || window.confirm(t('confirm.discard'));
  }

  async open(): Promise<void> {
    if (!this.confirmDiscard()) return;
    const file = await openFile('any');
    if (file) await this.openBytes(file.name, file.bytes, file.handle);
  }

  async openBytes(name: string, bytes: Uint8Array, handle?: FileSystemFileHandle): Promise<void> {
    this.events.onStatus(t('status.loading', { name }));
    try {
      if (/\.xlsx$/i.test(name)) {
        const [{ importXlsx }] = await loadXlsx();
        const data = await importXlsx(bytes, { name });
        this.load(data, name, handle);
      } else if (/\.(csv|txt|tsv)$/i.test(name)) {
        const { csvToWorkbook } = await import('../io/csv');
        const text = new TextDecoder('utf-8').decode(bytes);
        // Un CSV ouvert devient un classeur ; « Enregistrer » produira un .xlsx (comme Excel le propose).
        this.load(csvToWorkbook(text, name), name.replace(/\.[^.]+$/, '.xlsx'));
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
    const [, { exportXlsx }] = await loadXlsx();
    return exportXlsx(this.snapshot());
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
    const { csvToSheet } = await import('../io/csv');
    const text = new TextDecoder('utf-8').decode(file.bytes);
    const data = this.snapshot();
    const used = new Set(Object.values(data.sheets).map((s) => s.name));
    let name = file.name.replace(/\.[^.]+$/, '').slice(0, 28) || 'CSV';
    for (let i = 2; used.has(name); i++) name = `${name.replace(/ \(\d+\)$/, '')} (${i})`;
    const sheet = csvToSheet(text, name);
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
