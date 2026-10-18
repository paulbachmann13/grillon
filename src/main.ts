import { DocumentController } from './app/document';
import { CHART_COMPONENT, createChartComponent } from './charts/chart-view';
import { createShapeComponent, SHAPE_COMPONENT } from './charts/shape-view';
import { createSlicerComponent, SLICER_COMPONENT } from './univer/slicers';
import { createSparklineComponent, SPARKLINE_COMPONENT } from './charts/sparkline-view';
import { setupPwa } from './app/pwa';
import { installShortcuts } from './app/shortcuts';
import { Topbar } from './app/topbar';
import { t } from './i18n';
import { isTauri, secondaryWindowLabel, takeStartupOpens, takeWindowOpen, tauriHandleFromPath } from './io/tauri-files';
import { disableAutosave } from './storage/autosave';
import { setupUniver } from './univer/setup';
import './style.css';
import { guardDesktopClose } from './app/confirm';

const api = setupUniver(document.getElementById('univer')!);
// Le composant graphique doit exister avant le chargement d'un classeur qui en contient.
api.registerComponent(CHART_COMPONENT, createChartComponent(api) as never);
api.registerComponent(SPARKLINE_COMPONENT, createSparklineComponent(api) as never);
api.registerComponent(SHAPE_COMPONENT, createShapeComponent(api) as never);
api.registerComponent(SLICER_COMPONENT, createSlicerComponent(api) as never);

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
  const secondary = isTauri() ? secondaryWindowLabel() : null;
  if (secondary) {
    // Fenêtre supplémentaire (Fichier > Nouveau / Ouvrir, double-clic dans le Finder) : un classeur
    // à elle ; la sauvegarde automatique et la restauration restent à la fenêtre principale.
    disableAutosave();
    const path = await takeWindowOpen(secondary);
    if (path) await doc.openHandle(tauriHandleFromPath(path), true);
    else doc.newWorkbook();
  } else {
    if (await doc.restore()) topbar?.setStatus(t('status.restored'), 'success');
    else doc.newWorkbook();
    // Application de bureau : fichiers ouverts depuis le Finder (double-clic) avant le démarrage ;
    // les suivants arrivent dans leur propre fenêtre (côté natif).
    if (isTauri()) {
      for (const path of await takeStartupOpens()) await doc.openHandle(tauriHandleFromPath(path));
      // Mise à jour : vérification discrète quelques secondes après le démarrage (jamais pendant les tests).
      setTimeout(() => void doc.checkForUpdates(true), 5000);
    }
  }

  // Ouverture de fichiers par glisser-déposer.
  window.addEventListener('dragover', (e) => e.preventDefault());
  window.addEventListener('drop', async (e) => {
    const file = e.dataTransfer?.files?.[0];
    if (!file) return;
    e.preventDefault();
    if (!(await doc.confirmDiscard())) return;
    await doc.openBytes(file.name, new Uint8Array(await file.arrayBuffer()));
  });

  // Application de bureau : confirmation avant de fermer un classeur non enregistré.
  void guardDesktopClose(() => doc.dirty, () => doc.autosaveNow());

  // Dernière sauvegarde automatique avant de quitter.
  window.addEventListener('pagehide', () => void doc.autosaveNow());
  document.addEventListener('visibilitychange', () => {
    if (document.visibilityState === 'hidden') void doc.autosaveNow();
  });
}

void start();
