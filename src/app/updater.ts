/**
 * Mises à jour automatiques de l'application Mac : à l'ouverture (discrètement), puis à la demande
 * (Fichier → Rechercher les mises à jour). Les versions sont des Releases GitHub signées.
 * La logique est séparée de Tauri (`Deps`) pour pouvoir être testée.
 */
import { t } from '../i18n';

export interface AvailableUpdate {
  version: string;
  /** Ce que la nouvelle version apporte (le CHANGELOG de cette version, écrit dans latest.json). */
  notes?: string;
  /** Télécharge et installe ; `onProgress` reçoit les octets reçus et le total (s'il est connu). */
  install(onProgress: (received: number, total: number | undefined) => void): Promise<void>;
}

export interface Deps {
  check(): Promise<AvailableUpdate | null>;
  currentVersion(): Promise<string>;
  ask(message: string, notes?: string): Promise<boolean>;
  tell(message: string): Promise<void>;
  /** Où en est la mise à jour ; `percent` si le total est connu. */
  status(message: string, kind?: 'info' | 'success' | 'error', percent?: number): void;
  relaunch(): Promise<void>;
  /** Dernière sauvegarde automatique du classeur avant de redémarrer. */
  beforeRestart(): Promise<void>;
  /** Vrai si macOS exécute Grillon depuis un emplacement temporaire en lecture seule (lancé depuis le .dmg ou Téléchargements). */
  isTranslocated(): Promise<boolean>;
  /** Verrou commun à toutes les fenêtres : faux si une mise à jour est déjà en cours. */
  tryLock(): Promise<boolean>;
  unlock(): Promise<void>;
}

/** Chemin d'une application lancée depuis le .dmg ou Téléchargements : macOS la copie (« App Translocation ») en lecture seule. */
export const isTranslocatedPath = (path: string): boolean => /\/AppTranslocation\//.test(path);

export type UpdateResult = 'none' | 'declined' | 'installed' | 'error' | 'busy';

/**
 * Une seule vérification / installation à la fois dans toute l'application (verrou commun à toutes les fenêtres). Sans cela,
 * un clic sur « Rechercher les mises à jour » pendant le téléchargement, ou une autre fenêtre, ouvrait une deuxième fenêtre
 * « Installer » (et un deuxième téléchargement). Après une installation réussie, le verrou reste pris : Grillon redémarre.
 */
export async function runUpdateCheck(deps: Deps, silent: boolean): Promise<UpdateResult> {
  if (!(await deps.tryLock())) {
    if (!silent) await deps.tell(t('update.busy'));
    return 'busy';
  }
  let result: UpdateResult = 'error';
  try {
    result = await checkAndInstall(deps, silent);
    return result;
  } finally {
    if (result !== 'installed') await deps.unlock().catch(() => undefined);
  }
}

/**
 * `silent` : vérification au démarrage (pas de message si tout est à jour ou si la vérification échoue,
 * par exemple hors connexion). Sinon (demande de l'utilisateur) on répond toujours.
 */
async function checkAndInstall(deps: Deps, silent: boolean): Promise<UpdateResult> {
  let update: AvailableUpdate | null;
  try {
    update = await deps.check();
  } catch (e) {
    if (!silent) await deps.tell(t('update.checkFailed', { message: (e as Error)?.message ?? String(e) }));
    return 'error';
  }
  if (!update) {
    if (!silent) await deps.tell(t('update.upToDate', { version: await deps.currentVersion() }));
    return 'none';
  }
  // Lancée depuis le .dmg ou Téléchargements, l'application ne peut pas se remplacer elle-même.
  if (await deps.isTranslocated().catch(() => false)) {
    await deps.tell(t('update.moveApp', { version: update.version }));
    return 'error';
  }
  if (!(await deps.ask(t('update.available', { version: update.version, current: await deps.currentVersion() }), update.notes))) return 'declined';
  try {
    deps.status(t('update.downloading'));
    await update.install((received, total) => {
      const percent = total ? Math.min(100, Math.round((received / total) * 100)) : undefined;
      deps.status(percent !== undefined ? t('update.progress', { percent }) : t('update.downloading'), 'info', percent);
    });
    deps.status(t('update.installed'), 'success');
    await deps.beforeRestart().catch(() => undefined);
    await deps.relaunch();
    return 'installed';
  } catch (e) {
    deps.status(t('update.failed', { message: (e as Error)?.message ?? String(e) }), 'error');
    return 'error';
  }
}

/**
 * Panneau de progression : il couvre toute la fenêtre pendant le téléchargement, pour qu'on voie que la mise à jour avance
 * (un petit texte dans la barre passait inaperçu, et on recliquait). Un échec le ferme et s'affiche dans un message.
 */
function progressPanel(): { show(message: string, percent?: number): void; hide(): void } {
  let backdrop: HTMLDivElement | null = null;
  let label: HTMLParagraphElement;
  let bar: HTMLProgressElement;
  return {
    show(message, percent) {
      if (!backdrop) {
        backdrop = document.createElement('div');
        backdrop.className = 'g-dialog-backdrop';
        backdrop.dataset.testid = 'update-progress';
        const dialog = document.createElement('div');
        dialog.className = 'g-dialog';
        const title = document.createElement('h2');
        title.textContent = t('update.progressTitle');
        const body = document.createElement('div');
        body.className = 'g-dialog-body';
        label = document.createElement('p');
        bar = document.createElement('progress');
        bar.max = 100;
        bar.style.width = '100%';
        const hint = document.createElement('p');
        hint.textContent = t('update.progressHint');
        body.append(label, bar, hint);
        dialog.append(title, body);
        backdrop.append(dialog);
        document.body.append(backdrop);
      }
      label.textContent = message;
      if (percent === undefined) bar.removeAttribute('value');
      else bar.value = percent;
    },
    hide() {
      backdrop?.remove();
      backdrop = null;
    },
  };
}

/** Dépendances réelles (application de bureau). */
export async function tauriDeps(status: Deps['status'], beforeRestart: Deps['beforeRestart']): Promise<Deps> {
  const { check } = await import('@tauri-apps/plugin-updater');
  const { getVersion } = await import('@tauri-apps/api/app');
  const { ask, message } = await import('@tauri-apps/plugin-dialog');
  const { invoke } = await import('@tauri-apps/api/core');
  const panel = progressPanel();
  return {
    async check() {
      const update = await check();
      if (!update) return null;
      return {
        version: update.version,
        notes: update.body,
        async install(onProgress) {
          let received = 0;
          let total: number | undefined;
          await update.downloadAndInstall((event) => {
            if (event.event === 'Started') total = event.data.contentLength;
            else if (event.event === 'Progress') onProgress((received += event.data.chunkLength), total);
          });
        },
      };
    },
    currentVersion: getVersion,
    async ask(text, notes) {
      // Avec le détail de la version : notre fenêtre (la boîte du système ne sait pas afficher une liste).
      if (notes?.trim()) {
        const [{ openDialog }, { notesElement }] = await Promise.all([import('../ui/dialog'), import('./whats-new')]);
        const body = document.createElement('div');
        body.className = 'g-dialog-body';
        for (const paragraph of text.split('\n\n')) body.append(Object.assign(document.createElement('p'), { textContent: paragraph }));
        body.append(notesElement(notes));
        return openDialog({ title: t('app.name'), body, okLabel: t('update.install'), cancelLabel: t('update.later'), onOk: () => undefined, wide: true, testId: 'dialog-update' });
      }
      return ask(text, { title: t('app.name'), kind: 'info', okLabel: t('update.install'), cancelLabel: t('update.later') });
    },
    tell: async (text) => void (await message(text, { title: t('app.name') })),
    status(text, kind, percent) {
      if (kind === 'error') {
        panel.hide();
        status(text, kind);
        void message(text, { title: t('app.name'), kind: 'error' });
      } else panel.show(text, percent);
    },
    // Redémarrage maison (Rust) : celui du plugin crée une 2e instance, donc une 2e icône dans le Dock.
    relaunch: () => invoke<void>('relaunch_app'),
    beforeRestart,
    async isTranslocated() {
      const { resourceDir } = await import('@tauri-apps/api/path');
      return isTranslocatedPath(await resourceDir());
    },
    tryLock: () => invoke<boolean>('update_lock_acquire'),
    unlock: () => invoke<void>('update_lock_release'),
  };
}
