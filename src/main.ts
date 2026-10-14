import { DocumentController } from './app/document';
import { setupPwa } from './app/pwa';
import { installShortcuts } from './app/shortcuts';
import { Topbar } from './app/topbar';
import { t } from './i18n';
import { setupUniver } from './univer/setup';
import './style.css';

const api = setupUniver(document.getElementById('univer')!);

let topbar: Topbar | undefined;
const doc = new DocumentController(api, {
  onStatus: (message, kind) => topbar?.setStatus(message, kind),
  onChange: () => topbar?.refresh(),
});
topbar = new Topbar(document.getElementById('topbar')!, doc);
installShortcuts(doc);
setupPwa(
  doc,
  (msg, kind) => topbar?.setStatus(msg, kind),
  (msg, onAccept) => topbar?.setAction(msg, onAccept),
);

// Accès pour les tests de bout en bout et le débogage.
(window as unknown as { grillon: unknown }).grillon = { api, doc };

async function start(): Promise<void> {
  if (await doc.restore()) topbar?.setStatus(t('status.restored'), 'success');
  else doc.newWorkbook();

  // Ouverture de fichiers par glisser-déposer.
  window.addEventListener('dragover', (e) => e.preventDefault());
  window.addEventListener('drop', async (e) => {
    const file = e.dataTransfer?.files?.[0];
    if (!file) return;
    e.preventDefault();
    if (doc.dirty && !window.confirm(t('confirm.discard'))) return;
    await doc.openBytes(file.name, new Uint8Array(await file.arrayBuffer()));
  });

  // Dernière sauvegarde automatique avant de quitter.
  window.addEventListener('pagehide', () => void doc.autosaveNow());
  document.addEventListener('visibilitychange', () => {
    if (document.visibilityState === 'hidden') void doc.autosaveNow();
  });
}

void start();
