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
- Formules françaises (`src/univer/french-formulas.ts`) : le classeur garde l'anglais. `next()` des intercepteurs de
  cellule d'Univer ne chaîne pas et l'intercepteur du plugin de formules coupe la chaîne `BEFORE_CELL_EDIT` : on
  enveloppe donc `writeCellInterceptor.fetchThroughInterceptors` (monkeypatch) pour traduire le résultat final ;
  `AFTER_CELL_EDIT` est un intercepteur normal. Autocomplétion et aide : `IDescriptionService` est patché de la même
  façon. À revérifier à chaque mise à jour d'Univer (tests `tests/e2e/formules-fr.spec.ts`).
- Saisie : l'analyse d'Univer (fr) lit « 1.5 » comme 15, « 3/4 » comme le 4 mars et laisse en texte « 1 234,5 », « 31/12 »,
  VRAI. `src/univer/french-entry.ts` (intercepteur `AFTER_CELL_EDIT` prioritaire, logique dans `src/data/french-entry.ts`)
  convertit ces saisies et arrête la chaîne (sinon l'intercepteur des formats relit la valeur). Test : `tests/e2e/saisie.spec.ts`.
- Formules partagées d'Univer : une formule étendue (autoFill, Remplir, collage) n'est portée que par la première cellule du
  groupe (`f` + `si`) ; les autres n'ont que `si` et une valeur. L'export doit les reconstruire (`shiftFormula`,
  `src/data/formula-shift.ts`) sinon elles sont écrites comme des nombres. Tout code qui lit `cell.f` directement doit y penser.
- Tri : Univer réordonne les lignes (`sheet.command.reorder-range`) sans adapter les formules ; `src/univer/sort-formulas.ts`
  ajoute, via `interceptCommand`, les mutations qui décalent chaque formule de sa ligne (comme Excel ; une seule annulation).
- À l'ouverture, Univer reprend les valeurs enregistrées sans recalculer : `DocumentController.load` lance un calcul complet
  (sinon les formules qui débordent n'affichent que leur 1re valeur). ExcelJS n'écrit pas une ligne sans cellule ni hauteur :
  on donne une hauteur aux lignes masquées. `tests/e2e/persistence.spec.ts` vérifie formules et valeurs après enregistrement.
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

## Écran d'accueil

- `src/app/start-screen.ts` (interface), `src/data/recent.ts` (logique testée), `src/storage/recent.ts` (liste dans localStorage,
  handles du navigateur dans IndexedDB), `src/app/settings.ts`. Le numéro de version vient de `package.json` (`__APP_VERSION__`, vite.config.ts).
- Il s'affiche **toujours** au démarrage. Les travaux non enregistrés des dernières sessions (cinq au plus) sont mis de côté (`stashDraft`,
  clé `drafts` d'IndexedDB ; l'ancienne clé `draft` est migrée) et proposés en « Reprendre » ; jamais rouverts d'office. Les tests e2e posent
  `localStorage['grillon-no-home']` (dans `gotoApp`) : accueil masqué et ancien comportement (reprise automatique) ;
  `tests/e2e/home.spec.ts` teste le vrai démarrage sans ce réglage.

## Mises à jour automatiques (application Mac)

- Plugin `tauri-plugin-updater` : `src/app/updater.ts` (logique testée) ; vérification discrète 5 s après le démarrage
  (fenêtre principale seulement) et Fichier → Rechercher les mises à jour. Point d'accès :
  `https://github.com/paulbachmann13/grillon/releases/latest/download/latest.json` (dans `tauri.conf.json`).
- Signature des mises à jour : clé générée avec `npx tauri signer generate`. La clé **privée** est dans
  `~/.tauri/grillon-updater.key` (sauvegarde à garder : sans elle, plus aucune mise à jour possible pour les
  copies déjà installées) et dans le secret GitHub `TAURI_SIGNING_PRIVATE_KEY` ; la clé publique est dans `tauri.conf.json`.
- Le workflow « Application Mac » (étiquette `vX.Y.Z`, qui doit égaler la version de `package.json`) publie le .dmg,
  `Grillon_X.Y.Z_universal.app.tar.gz` et `latest.json` (`scripts/make-latest-json.mjs`).
- Ne jamais changer la clé publique sans prévenir : les versions installées refuseraient les mises à jour suivantes.

## Pas encore vérifié (à tester sur un vrai Mac / avec le vrai Excel)

- Vérifié le 18/10 dans Excel (AppleScript) : un tableau croisé écrit par Grillon s'ouvre et se calcule, 8 graphiques
  classiques s'ouvrent, les sauts de page sont relus. Pas encore vérifié : graphiques chartEx (cascade, entonnoir…), thèmes.
- Vérifié dans WebKit (moteur de l'app Mac) : ⌘C / ⌘V / ⌘Z, ⌘B… ; l'app Mac compile et se lance, la mise à jour automatique
  fonctionne (1.12 → 1.13). Pas encore vérifié : impression (⌘P) dans l'app Mac, double-clic dans le Finder, plusieurs fenêtres,
  bouton Coller du ruban (autorisation du presse-papiers de WebKit).

## Idées pour la suite (voir docs/EXCEL-COMPARAISON.md)

Triangles verts sur la grille, flèches de repérage, segments et chronologies dans le .xlsx (x14), regroupement de dates et
« afficher les valeurs en » dans les tableaux croisés, protection du classeur et plages autorisées,
orthographe, effets de thème, plus de formes, autres types de graphiques (boursier, surface…).
