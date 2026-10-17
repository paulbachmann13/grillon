# Grillon — consignes pour Claude Code

Grillon est un **tableur local et gratuit**, modelé sur **Excel** (c'est la référence : « le plus proche
possible d'Excel »). PWA + application Mac (Tauri). Interface **en français**.
Dépôt public : paulbachmann13/grillon — site : https://paulbachmann13.github.io/grillon/

## Façon de travailler (préférences de l'utilisateur)

- **Parler français** avec l'utilisateur ; textes de l'interface dans `src/i18n/fr.ts` (jamais en dur ailleurs).
- **Autonomie** : avancer sans poser de questions ; ne demander que si une décision bloque vraiment.
- **Excel est la référence** : s'inspirer du comportement, des noms et des raccourcis d'Excel (pas d'un autre
  logiciel). `docs/EXCEL-COMPARAISON.md` liste les commandes du ruban d'Excel et leur état : c'est la feuille de route.
- **Composants gratuits uniquement**, licences vérifiées (Apache-2.0, MIT…). Jamais de paquet `@univerjs-pro/*`
  (payant) ; HyperFormula exclu (GPLv3).
- Après chaque étape : mettre à jour `CHANGELOG.md` et `docs/` (FONCTIONNALITES.md, EXCEL-COMPARAISON.md,
  limites connues), puis un **court résumé avec le lien** du site à l'utilisateur.
- **Commits** : Conventional Commits en français (`feat:`, `fix:`, `docs:`, `ci:`…). Pousser sur `main` :
  la CI déploie sur GitHub Pages si les tests passent.
- Le numéro de version est dans `package.json`, `src-tauri/tauri.conf.json` et `src-tauri/Cargo.toml` (à garder identiques).

## Commandes

```bash
npm install
npm run dev            # http://localhost:5173
npm run typecheck
npm test               # Vitest (tests unitaires, Node)
npm run test:e2e       # Playwright : construit puis sert dist/ ; Chromium + WebKit (iPad)
npm run build          # build de production
npm run tauri dev      # application Mac en développement (Rust requis)
npm run build:mac      # .dmg universel (Mac avec Rust ; sinon via GitHub Actions)
```

Avant de pousser : `npm run typecheck && npm test && npm run test:e2e`.
Les tests e2e tournent sur le **build de production** : un import dynamique de `/src/...` n'y marche pas ;
utiliser `window.grillon.api` et `window.grillon.doc` (exposés par `src/main.ts` pour les tests).
Les tests LibreOffice (`tests/unit/libreoffice.test.ts`) sont ignorés si `soffice` est absent
(`brew install --cask libreoffice` pour les lancer en local) ; ils tournent toujours en CI.

## Architecture (voir docs/ARCHITECTURE.md)

- `src/main.ts` : point d'entrée. `src/app/document.ts` : classeur courant (ouvrir, enregistrer, sauvegarde auto).
- `src/univer/` : configuration d'Univer, ajouts au ruban (`*-ribbon.ts`), protection, plan, format de cellule…
- `src/data/` : **logique pure et testée** (sans DOM) — à privilégier pour toute nouvelle règle métier.
- `src/io/xlsx/` : .xlsx ⇄ modèle Univer (ExcelJS + couche OOXML maison avec JSZip). Sans DOM ni Univer.
- `src/charts/` : graphiques ECharts (+ cascade, entonnoir, histogramme, boîte à moustaches), formes, mini-graphiques.
- `src-tauri/` : application Mac (Tauri 2). `src/io/tauri-files.ts` : fichiers et fenêtres côté JS.
- `tests/unit`, `tests/e2e`, `tests/fixtures`. Chaque fonctionnalité a ses tests, y compris l'aller-retour .xlsx.

## Pièges connus (Univer 1.0.x)

- Les presets open source seulement. Le moteur de formules est celui d'Univer ; `getFormula().executeFormulas`
  calcule des sous-expressions (utilisé par « Évaluer la formule »).
- Marge du plan : les en-têtes de lignes/colonnes sont agrandis (`setRowHeaderWidth`) ; les boutons sont une
  couche HTML posée sur le canevas (`src/univer/outline-gutter.ts`).
- Protection : `protect()` + `setReadOnly()` bloquent tout, mais l'état n'est pas enregistré dans le snapshot ;
  il est donc gardé dans `sheet.custom.grillon` et réappliqué à l'ouverture. Avec des cellules déverrouillées,
  c'est une garde maison (`installProtectionGuard`, événement `BeforeCommandExecute`) qui protège.
- L'événement `SelectionChanged` ne se déclenche pas pour une sélection programmatique : écouter la commande
  `sheet.operation.set-selections`.
- `sheet.command.set-range-values` fusionne les styles et garde les valeurs ; ne pas écraser `custom.grillon`
  d'une feuille (le fusionner).
- Le calcul manuel met en pause `FormulaCalculationTriggerService` (champs privés) : à revérifier à chaque
  mise à jour d'Univer.
- Écriture .xlsx : ExcelJS puis post-traitement XML (ordre des éléments : `WORKSHEET_ORDER`). Tableaux croisés,
  mini-graphiques, formes, graphiques chartEx, protection sont écrits à la main.

## Application Mac (Tauri)

- Une fenêtre = un classeur (comme Excel). `src-tauri/src/lib.rs` : commandes `take_pending_opens`,
  `take_window_open`, `new_window`. Les fenêtres secondaires (`doc-N`, URL `?w=doc-N`) ne font ni sauvegarde
  automatique ni restauration (réservées à la fenêtre `main`). Autorisations dans `src-tauri/capabilities/default.json`.
- Le .dmg est **non signé** (pas de compte développeur Apple) : « Ouvrir quand même » dans Réglages Système,
  ou `xattr -cr /Applications/Grillon.app`. Il est construit par le workflow GitHub « Application Mac »
  (`gh workflow run mac.yml`), artefact `Grillon-mac`.
- Le code Rust n'a longtemps été écrit sans pouvoir le compiler localement : sur le Mac, lancer
  `npm run tauri dev` pour le vérifier.

## Pas encore vérifié (à tester sur un vrai Mac / avec le vrai Excel)

- Ouverture dans Microsoft Excel des tableaux croisés dynamiques et des graphiques chartEx écrits par Grillon.
- Impression (⌘P) dans l'application Mac (WKWebView).
- ⌘C / ⌘V / ⌘Z, ouverture par double-clic dans le Finder, plusieurs fenêtres.

## Idées pour la suite (voir docs/EXCEL-COMPARAISON.md)

Triangles verts sur la grille, flèches de repérage, segments et chronologies, regroupement de dates et
« afficher les valeurs en » dans les tableaux croisés, protection du classeur et plages autorisées,
sauts de page manuels, orthographe, thèmes, plus de formes, autres types de graphiques (boursier, surface…).
