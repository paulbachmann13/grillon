/** Réglages de l'application (par appareil), enregistrés dans localStorage. */
const AUTO_UPDATE = 'grillon-auto-update';

/** Vérifier les mises à jour au démarrage (application de bureau) : activé par défaut. */
export function autoUpdateEnabled(): boolean {
  try {
    return localStorage.getItem(AUTO_UPDATE) !== 'off';
  } catch {
    return true;
  }
}

export function setAutoUpdate(enabled: boolean): void {
  try {
    if (enabled) localStorage.removeItem(AUTO_UPDATE);
    else localStorage.setItem(AUTO_UPDATE, 'off');
  } catch {
    /* stockage indisponible : le choix vaut pour la session */
  }
}
