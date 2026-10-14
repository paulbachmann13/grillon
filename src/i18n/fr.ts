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
  ribbon: {
    cut: 'Couper',
    copy: 'Copier',
    pasteFormulas: 'Coller les formules',
    insert: 'Insérer',
    insertCellsDown: 'Insérer des cellules (décaler vers le bas)',
    insertCellsRight: 'Insérer des cellules (décaler vers la droite)',
    insertRows: 'Insérer des lignes dans la feuille',
    insertCols: 'Insérer des colonnes dans la feuille',
    insertSheet: 'Insérer une feuille',
    delete: 'Supprimer',
    deleteCellsUp: 'Supprimer des cellules (décaler vers le haut)',
    deleteCellsLeft: 'Supprimer des cellules (décaler vers la gauche)',
    deleteRows: 'Supprimer des lignes dans la feuille',
    deleteCols: 'Supprimer des colonnes dans la feuille',
    deleteSheet: 'Supprimer la feuille',
    format: 'Format',
    autoRowHeight: 'Ajuster la hauteur de ligne',
    autoColWidth: 'Ajuster la largeur de colonne',
    hideRows: 'Masquer les lignes',
    hideCols: 'Masquer les colonnes',
    renameSheet: 'Renommer la feuille',
    clear: 'Effacer',
    clearAll: 'Effacer tout',
    clearFormats: 'Effacer les formats',
    clearContents: 'Effacer le contenu',
    indentMore: 'Augmenter le retrait',
    indentLess: 'Diminuer le retrait',
    conditionalFormatting: 'Mise en forme conditionnelle',
    sortFilter: 'Trier et filtrer',
    sortAsc: 'Trier de A à Z',
    sortDesc: 'Trier de Z à A',
    sortCustom: 'Tri personnalisé…',
    filter: 'Filtrer',
    find: 'Rechercher',
  },
  confirm: {
    discard: 'Le classeur actuel a des modifications non enregistrées. Continuer quand même ?',
  },
} as const;

type Widen<T> = { [K in keyof T]: T[K] extends string ? string : Widen<T[K]> };
export type Messages = Widen<typeof fr>;
