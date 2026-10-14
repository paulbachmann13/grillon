# Architecture de Grillon

## Vue d'ensemble

```
┌──────────────────────────── Navigateur (PWA) ─────────────────────────────┐
│  Barre Fichier (src/app/topbar.ts)                                        │
│  ┌─────────────────────────── Univer (Apache-2.0) ──────────────────────┐ │
│  │ Ruban · barre de formule · grille · onglets de feuilles               │ │
│  │ Modèle : IWorkbookData        Moteur de formules ──► Web Worker       │ │
│  └───────────────────────────────────────────────────────────────────────┘ │
│        ▲ load(IWorkbookData)            │ snapshot() = workbook.save()     │
│        │                                ▼                                  │
│  DocumentController (src/app/document.ts) ── IndexedDB (sauvegarde auto)   │
│        ▲                                │                                  │
│  Convertisseurs purs (src/io) : xlsx ⇄ IWorkbookData, csv ⇄ IWorkbookData │
│        ▲                                ▼                                  │
│  Fichiers : File System Access API (Chrome/Edge) ou sélecteur/téléchargement│
└────────────────────────────────────────────────────────────────────────────┘
```

Tout tourne dans le navigateur : aucune donnée n'est envoyée sur un serveur.

## Dossiers

| Dossier | Rôle |
|---|---|
| `src/main.ts` | Point d'entrée : crée Univer, le document, la barre Fichier, la PWA. |
| `src/univer/` | Configuration d'Univer (presets open source uniquement), worker de calcul, affichage à la française. |
| `src/app/` | Logique d'application : document courant, barre Fichier, raccourcis, PWA. |
| `src/io/xlsx/` | Convertisseur .xlsx ⇄ modèle Univer (ExcelJS). **Sans DOM ni Univer à l'exécution** : testable dans Node. |
| `src/io/csv.ts` | Import/export CSV (PapaParse). |
| `src/io/files.ts` | Ouvrir/enregistrer (File System Access API ou repli). |
| `src/storage/` | Sauvegarde automatique (IndexedDB). |
| `src/i18n/` | Textes de l'interface Grillon (`fr.ts`). Les textes d'Univer viennent de ses fichiers `fr-FR`. |
| `src/model/` | Types du modèle Univer et constantes recopiées (pour ne pas charger Univer dans Node). |
| `tests/fixtures/` | Générateur des classeurs de test (ExcelJS). |
| `tests/unit/` | Tests Vitest (convertisseurs, round-trip, LibreOffice). |
| `tests/e2e/` | Tests Playwright sur le build de production (UI, round-trip via l'interface, hors-ligne). |

## Modèle de données

Le modèle de référence est celui d'Univer, `IWorkbookData` :

- `styles` : dictionnaire de styles partagés (`IStyleData`), référencés par identifiant depuis les cellules.
  L'import dédoublonne les styles (un fichier de 50 000 lignes n'a que quelques styles).
- `sheets[id]` : `cellData[ligne][colonne] = { v, t, f, s }`
  - `v` valeur, `t` type (1 texte, 2 nombre, 3 booléen, 4 texte forcé), `f` formule (`=SUM(A1:A3)`), `s` style.
  - Les dates sont des numéros de série Excel (nombre + format de date), exactement comme dans Excel.
- `rowData` / `columnData` : hauteurs/largeurs en **pixels**, masquage.
- `mergeData` : plages fusionnées ; `freeze` : volets figés.
- `resources` : données des plugins (noms définis `SHEET_DEFINED_NAME_PLUGIN`, et plus tard
  mises en forme conditionnelles, validations, filtres…).

## Couche de conversion xlsx

`importXlsx(bytes) → IWorkbookData` et `exportXlsx(IWorkbookData) → bytes`, construits sur **ExcelJS (MIT)**.

| Élément | Excel (ExcelJS) | Univer |
|---|---|---|
| Largeur de colonne | caractères | px = car × 7 + 5 |
| Hauteur de ligne | points | px = pt × 96/72 |
| Couleur | ARGB, thème + teinte, indexée | `#RRGGBB` (thème résolu depuis `theme1.xml`, teinte appliquée en HSL) |
| Bordures | `thin`, `double`… | `BorderStyleTypes` |
| Alignements, retour à la ligne, rotation, retrait | `alignment` | `ht`, `vt`, `tb`, `tr`, `pd` |
| Formats de nombre | `numFmt` | `n.pattern` (même syntaxe) |
| Formules partagées | maître + plages | formule traduite pour chaque cellule |
| Dates | `Date` | numéro de série |
| Texte enrichi | `richText` | texte brut (limite connue) |

Les éléments qui n'ont pas encore de représentation (graphiques, mises en forme conditionnelles…)
seront ajoutés dans les phases suivantes, chacun avec son test de round-trip.

## Compléments au-dessus d'ExcelJS

-  : accès direct au paquet (JSZip + fast-xml-parser) pour ce qu'ExcelJS ignore :
  critères de filtre, tableaux structurés, « Interrompre si vrai » des MFC, graphiques et images
  (DrawingML), validations sur de grandes plages (lues sans être dépliées cellule par cellule).
  Les éléments sont insérés dans la feuille dans l'ordre exigé par le schéma OOXML.
-  : la conversion tourne dans un Web Worker (repli dans la page si besoin).

## Fonctionnalités maison

| Module | Rôle |
|---|---|
|  | Graphiques ECharts dans des « DOM flottants » Univer (sauvegardés avec le classeur) |
|  | Logique pure et testée : doublons, convertir, remplissage instantané, sous-totaux, filtre avancé, plan, TCD, valeur cible, références de formules |
|  | Ajouts au ruban (Démarrer, Données, Formules, Vue), commandes annulables |
|  | Affichage à la française du format Standard et des booléens |
|  | Fonctions ajoutées au moteur (worker et page) |
|  | Impression / PDF |

## Moteur de formules — choix

Univer embarque son propre moteur de formules (`@univerjs/engine-formula`, Apache-2.0), branché
sur la barre de formule (autocomplétion, aide aux arguments, coloration des références) et exécuté
dans un Web Worker. **Décision (validée) : on utilise ce moteur plutôt que HyperFormula.**

- HyperFormula est sous **GPLv3** (ou licence commerciale) : l'utiliser aurait imposé la GPLv3 à
  toute l'application et demandé une couche de synchronisation Univer ⇄ HyperFormula.
- Le moteur d'Univer couvre plus de 400 fonctions Excel, noms en anglais (`SUM`, `XLOOKUP`…).
- Les fonctions manquantes seront ajoutées comme fonctions personnalisées Univer (phase 4).

## Licences

Uniquement des composants gratuits et open source :

| Composant | Licence | Usage |
|---|---|---|
| Univer (`@univerjs/*`, presets open source) | Apache-2.0 | Grille, ruban, formules |
| ExcelJS | MIT | Lecture/écriture .xlsx |
| PapaParse | MIT | CSV |
| React (dépendance d'Univer) | MIT | Interface d'Univer |
| vite-plugin-pwa / Workbox | MIT | Hors-ligne |

Les plugins **Univer Pro** (import/export xlsx serveur, tableaux croisés, graphiques Pro, impression…)
ne sont **pas** utilisés. Aucun paquet `@univerjs-pro/*` n'est installé (vérifié en CI).

## Stratégie de tests

1. **Unitaires (Vitest, Node)** — convertisseurs :
   - lecture fidèle de classeurs générés (styles, formats, formules, noms, fusions, volets, 50 000 lignes) ;
   - round-trip import → export → import identique ;
   - CSV (séparateurs, décimales, dates, guillemets).
2. **LibreOffice (headless)** — chaque export est converti par LibreOffice puis relu : il doit
   s'ouvrir sans erreur et donner les mêmes valeurs (en CI ; ignoré si LibreOffice est absent).
3. **Bout en bout (Playwright, build de production)** :
   - démarrage en français, saisie et recalcul ;
   - round-trip **via l'interface** : Fichier → Ouvrir, puis Fichier → Enregistrer, comparaison avec l'original ;
   - CSV, sauvegarde automatique, onglets de feuilles ;
   - PWA : manifeste et fonctionnement hors ligne ; smoke test WebKit (profil iPad).

## Prêt pour Tauri

- Aucune API serveur : l'app est un ensemble de fichiers statiques.
- `base` Vite configurable (`GRILLON_BASE`) : `/grillon/` pour GitHub Pages, `/` ailleurs.
- L'accès aux fichiers est isolé dans `src/io/files.ts` : il suffira d'y ajouter une variante
  utilisant les API fichiers de Tauri.
