/**
 * Textes de l'interface Grillon (français).
 * Les textes d'Univer (ruban, menus contextuels…) viennent de ses propres fichiers fr-FR.
 * Pour ajouter une langue : copier ce fichier (ex. en.ts) en gardant les mêmes clés.
 */
export const fr = {
  app: {
    name: 'Grillon',
    untitled: 'Classeur1',
    defaultSheet: 'Feuil1',
  },
  file: {
    menu: 'Fichier',
    new: 'Nouveau classeur',
    open: 'Ouvrir…',
    save: 'Enregistrer',
    saveAs: 'Enregistrer sous…',
    importCsv: 'Importer un CSV…',
    exportCsv: 'Exporter la feuille en CSV',
    shortcutOpen: 'Ctrl+O',
    shortcutSave: 'Ctrl+S',
  },
  status: {
    loading: 'Ouverture de « {name} »…',
    saving: 'Enregistrement…',
    saved: 'Enregistré',
    savedTo: 'Enregistré dans « {name} »',
    downloaded: 'Téléchargé : « {name} »',
    autosaved: 'Sauvegarde automatique',
    restored: 'Classeur restauré depuis la dernière session',
    unsaved: 'Modifications non enregistrées',
    offlineReady: 'Grillon fonctionne maintenant hors ligne',
    updateAvailable: 'Nouvelle version disponible — recharger',
  },
  errors: {
    open: 'Impossible d’ouvrir ce fichier : {message}',
    save: 'Échec de l’enregistrement : {message}',
    unsupported: 'Format non pris en charge : {name}. Formats acceptés : .xlsx, .csv',
  },
  confirm: {
    discard: 'Le classeur actuel a des modifications non enregistrées. Continuer quand même ?',
  },
} as const;

type Widen<T> = { [K in keyof T]: T[K] extends string ? string : Widen<T[K]> };
export type Messages = Widen<typeof fr>;
