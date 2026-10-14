# Changelog

Format inspiré de [Keep a Changelog](https://keepachangelog.com/fr/1.1.0/).

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
