# Changelog

Format inspiré de [Keep a Changelog](https://keepachangelog.com/fr/1.1.0/).

## [1.1.0] — 2026-10-15

### Ajouté
- **Grille extensible** : au-delà de la ligne 1000 et de la colonne Z, des lignes et des colonnes s'ajoutent quand on défile ou qu'on déplace la sélection, jusqu'à 1 048 576 lignes et 16 384 colonnes (XFD) comme Excel.
- **Assistant d'import CSV** (Fichier → Ouvrir / Importer un CSV, Données → À partir d'un CSV) : origine du fichier (UTF-8 ou Windows/ANSI), séparateur, séparateur décimal, première ligne en en-têtes (gras et figée), aperçu.
- Vue → **Texte qui déborde** : par défaut, un texte trop long est coupé au bord de sa cellule ; ce bouton rétablit le débordement sur les cellules vides (comme Excel).

- **Application Mac** (Tauri) : .dmg universel construit par GitHub Actions (workflow « Application Mac »), Cmd+S réécrit le fichier d'origine, ouverture par double-clic dans le Finder. Non signée.

### Corrigé
- Les CSV enregistrés par Excel sous Windows (encodage ANSI) s'ouvrent avec leurs accents.

## [1.0.0] — Phase 5 : finitions — 2026-10-15

### Ajouté
- **Mode sombre** (Vue → Mode sombre ; suit le réglage du système par défaut, choix mémorisé).
- **Impression / export PDF** (Fichier → Imprimer, Ctrl/⌘+P, ou Vue) : feuille active ou sélection, portrait/paysage, quadrillage, ajustement à la largeur ; mise en forme, fusions, graphiques et images inclus. Le PDF s'obtient avec « Enregistrer au format PDF ».
- **Images** et **notes** (commentaires) enregistrées dans le .xlsx et relues depuis Excel.
- Raccourcis **Ctrl+;** (date du jour) et **Ctrl+:** (heure) ; les autres raccourcis Excel (Ctrl+C/V/X/Z/Y/B/I/U, F2, F4, Ctrl+flèches, Ctrl+Maj+flèches, Ctrl+E, Ctrl+S/O/P…) sont vérifiés par des tests.
- **Performances** : ouverture et enregistrement des .xlsx dans un Web Worker (100 000 lignes sans figer l'interface) ; grille virtualisée et calcul des formules en arrière-plan (Univer).

### Corrigé
- Les commandes d'interface (mode sombre, audit…) ne marquent plus le classeur comme modifié.

## [0.4.0] — Phase 4 : formules et analyse — 2026-10-15

### Ajouté
- **Graphiques** (ECharts, Apache-2.0) : histogramme, barres, courbes, aires (groupés ou empilés), secteurs, nuage de points, **graphique combiné** et **axe secondaire** ; insertion depuis la sélection (Insérer → Graphique), déplacement et redimensionnement à la souris, mise à jour en direct, boîte « Modifier le graphique » (type, titre, légende, titres d'axes, couleur et type de chaque série, axe secondaire, suppression).
- Graphiques **enregistrés dans le .xlsx** (DrawingML : s'ouvrent dans Excel et LibreOffice) et **relus** depuis les fichiers Excel.
- **Tableau croisé dynamique** simple : lignes, colonnes, valeurs (somme, nombre, moyenne, min, max), filtres, totaux et total général, dans une nouvelle feuille ; Actualiser et Modifier (Données → Tableau croisé dynamique).
- **Valeur cible** (Données).
- Formules : **Gestionnaire de noms**, **Repérer les antécédents / les dépendants** (liste cliquable).
- Fonctions **ISO.CEILING** et **TRIMRANGE** ; documentation de la couverture des fonctions (docs/FONCTIONS.md).

### Corrigé
- Serveur de développement sous Windows : le worker de calcul n'était pas trouvé dans certains dossiers redirigés.

## [0.3.0] — Phase 3 : onglet Données — 2026-10-14

### Ajouté
- Données : **Supprimer les doublons** (choix des colonnes, en-têtes, insensible à la casse), **Convertir** (texte en colonnes avec séparateurs, guillemets, nombres et dates reconnus), **Remplissage instantané** (bouton et Ctrl+E), **Filtre avancé** (zone de critères ET/OU, jokers, sur place ou copie, sans doublon) et « Afficher tout ».
- Données : **Plan** (grouper/dissocier lignes et colonnes, masquer/afficher le détail) et **Sous-total** (à chaque changement de…, fonctions Somme/Nombre/Moyenne/Max/Min/Produit, total général, niveaux de plan).
- **Tableaux structurés** (plugin Univer open source) : « Mettre sous forme de tableau », **ligne de totaux** (SUBTOTAL), références structurées dans les formules (`Ventes[Prix]`).
- Données : « À partir d'un CSV ».
- Import/export .xlsx : **validations de données** (liste, nombre entier/décimal, date, longueur de texte, formule, messages d'aide et d'erreur, listes pointant vers une plage, format x14 d'Excel), **filtres automatiques avec leurs critères** (lignes filtrées masquées comme dans Excel), **tableaux** (nom, style, ligne de totaux), **niveaux de plan**, option « Interrompre si vrai » des MFC.
- Une couche de lecture/écriture XML (JSZip + fast-xml-parser, MIT) complète ExcelJS ; les validations sur des colonnes entières ne sont plus dépliées cellule par cellule.

### Corrigé
- Format Standard : au plus 11 caractères, comme Excel (10/3 → 3,333333333).

## [0.2.0] — Phase 2 : onglet Démarrer complet — 2026-10-14

### Ajouté
- Ruban en grille (deux rangées, groupes), plus proche du ruban d'Excel.
- Démarrer : boutons Couper / Copier, « Coller les formules » dans le menu Coller, Augmenter / Diminuer le retrait.
- Démarrer : menus Insérer, Supprimer (cellules, lignes, colonnes, feuille), Format (ajustement automatique de la hauteur et de la largeur, masquer, renommer la feuille), Effacer (tout, formats, contenu).
- Démarrer : raccourcis Mise en forme conditionnelle, Trier et filtrer, Rechercher.
- Import/export .xlsx des **mises en forme conditionnelles** : règles de valeurs, texte, dates, doublons/valeurs uniques, top/bottom, moyenne, formule, barres de données, nuances de couleurs, jeux d'icônes.
- Import/export du **texte enrichi** (plusieurs styles dans une cellule, sauts de ligne).
- Annuler/rétablir sans limite (Univer en gardait 50).

### Corrigé
- Téléchargements plus fiables : le lien temporaire n'est plus retiré immédiatement.

## [0.1.0] — Phase 1 : fondations et import/export — 2026-10-14

### Ajouté
- Application Vite + TypeScript avec Univer (presets open source uniquement), interface en français.
- Menu Fichier : nouveau, ouvrir, enregistrer, enregistrer sous, importer/exporter CSV ; raccourcis Ctrl/⌘+O, Ctrl/⌘+S, Ctrl/⌘+Maj+S ; glisser-déposer.
- Convertisseur .xlsx ⇄ modèle Univer (ExcelJS) : valeurs, formules, types, dates, styles, couleurs de thème, bordures, alignements, formats de nombre, fusions, tailles, volets figés, feuilles masquées, noms définis, liens.
- Import/export CSV au format Excel français (« ; », virgule décimale, BOM UTF-8), détection automatique à l'import.
- Affichage à la française des nombres au format Standard et des booléens.
- Sauvegarde automatique dans IndexedDB et restauration au démarrage.
- PWA : manifeste, icônes, service worker, fonctionnement hors ligne, ouverture de fichiers depuis le système.
- Tests : Vitest (convertisseurs, round-trip, CSV, ouverture LibreOffice), Playwright (UI, round-trip via l'interface, hors-ligne, WebKit iPad).
- CI GitHub Actions (tests + déploiement GitHub Pages).
