/**
 * Questions oui / non : boîte native dans l'application de bureau (la fenêtre web du Mac
 * n'affiche pas toujours window.confirm), boîte du navigateur sinon.
 */
import { t } from '../i18n';
import { isTauri } from '../io/tauri-files';

export async function confirmAsync(message: string, title = t('app.name')): Promise<boolean> {
  if (isTauri()) {
    try {
      const { ask } = await import('@tauri-apps/plugin-dialog');
      return await ask(message, { title, kind: 'warning', okLabel: t('confirm.yes'), cancelLabel: t('confirm.no') });
    } catch {
      /* boîte native indisponible : on retombe sur celle de la fenêtre web */
    }
  }
  return window.confirm(message);
}

/** Application de bureau : demande confirmation avant de fermer une fenêtre dont le classeur n'est pas enregistré. */
export async function guardDesktopClose(isDirty: () => boolean, beforeClose: () => Promise<void>): Promise<void> {
  if (!isTauri()) return;
  const { getCurrentWindow } = await import('@tauri-apps/api/window');
  await getCurrentWindow().onCloseRequested(async (event) => {
    if (isDirty() && !(await confirmAsync(t('confirm.close')))) {
      event.preventDefault();
      return;
    }
    // Dernière sauvegarde automatique (le classeur sera restauré à la prochaine ouverture).
    await beforeClose().catch(() => undefined);
  });
}
