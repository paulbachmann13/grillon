/**
 * Service worker (hors-ligne) et ouverture de fichiers depuis le système (PWA installée).
 */
import { registerSW } from 'virtual:pwa-register';
import { t } from '../i18n';
import { isTauri } from '../io/tauri-files';
import type { DocumentController, StatusKind } from './document';

interface LaunchParams {
  files: FileSystemFileHandle[];
}

export function setupPwa(doc: DocumentController, notify: (msg: string, kind?: StatusKind) => void, askReload: (msg: string, onAccept: () => void) => void): void {
  // Pas de service worker en développement ni dans l'application de bureau (les fichiers sont déjà locaux).
  if (import.meta.env.DEV || isTauri()) return;
  const updateSW = registerSW({
    onNeedRefresh() {
      askReload(t('status.updateAvailable'), () => void updateSW(true));
    },
    onOfflineReady() {
      notify(t('status.offlineReady'), 'success');
    },
  });

  // Ouverture d'un .xlsx double-cliqué quand Grillon est installée (Chrome/Edge).
  const launchQueue = (window as unknown as { launchQueue?: { setConsumer: (cb: (p: LaunchParams) => void) => void } }).launchQueue;
  launchQueue?.setConsumer(async (params) => {
    const handle = params.files?.[0];
    if (!handle) return;
    const file = await handle.getFile();
    await doc.openBytes(file.name, new Uint8Array(await file.arrayBuffer()), handle);
  });
}
