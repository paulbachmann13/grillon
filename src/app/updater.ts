/**
 * Mises à jour automatiques de l'application Mac : à l'ouverture (discrètement), puis à la demande
 * (Fichier → Rechercher les mises à jour). Les versions sont des Releases GitHub signées.
 * La logique est séparée de Tauri (`Deps`) pour pouvoir être testée.
 */
import { t } from '../i18n';

export interface AvailableUpdate {
  version: string;
  /** Télécharge et installe ; `onProgress` reçoit les octets reçus et le total (s'il est connu). */
  install(onProgress: (received: number, total: number | undefined) => void): Promise<void>;
}

export interface Deps {
  check(): Promise<AvailableUpdate | null>;
  currentVersion(): Promise<string>;
  ask(message: string): Promise<boolean>;
  tell(message: string): Promise<void>;
  status(message: string, kind?: 'info' | 'success' | 'error'): void;
  relaunch(): Promise<void>;
  /** Dernière sauvegarde automatique du classeur avant de redémarrer. */
  beforeRestart(): Promise<void>;
}

export type UpdateResult = 'none' | 'declined' | 'installed' | 'error';

/**
 * `silent` : vérification au démarrage (pas de message si tout est à jour ou si la vérification échoue,
 * par exemple hors connexion). Sinon (demande de l'utilisateur) on répond toujours.
 */
export async function runUpdateCheck(deps: Deps, silent: boolean): Promise<UpdateResult> {
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
  if (!(await deps.ask(t('update.available', { version: update.version, current: await deps.currentVersion() })))) return 'declined';
  try {
    deps.status(t('update.downloading'));
    await update.install((received, total) => deps.status(total ? t('update.progress', { percent: Math.min(100, Math.round((received / total) * 100)) }) : t('update.downloading')));
    deps.status(t('update.installed'), 'success');
    await deps.beforeRestart().catch(() => undefined);
    await deps.relaunch();
    return 'installed';
  } catch (e) {
    deps.status(t('update.failed', { message: (e as Error)?.message ?? String(e) }), 'error');
    return 'error';
  }
}

/** Dépendances réelles (application de bureau). */
export async function tauriDeps(status: Deps['status'], beforeRestart: Deps['beforeRestart']): Promise<Deps> {
  const { check } = await import('@tauri-apps/plugin-updater');
  const { relaunch } = await import('@tauri-apps/plugin-process');
  const { getVersion } = await import('@tauri-apps/api/app');
  const { ask, message } = await import('@tauri-apps/plugin-dialog');
  return {
    async check() {
      const update = await check();
      if (!update) return null;
      return {
        version: update.version,
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
    ask: (text) => ask(text, { title: t('app.name'), kind: 'info', okLabel: t('update.install'), cancelLabel: t('update.later') }),
    tell: async (text) => void (await message(text, { title: t('app.name') })),
    status,
    relaunch,
    beforeRestart,
  };
}
