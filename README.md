# Grillon

Tableur local et gratuit, pensé pour remplacer Excel au quotidien : formules, formats, analyse de
données. Interface en français, noms de fonctions en anglais comme dans Excel (`SUM`, `IF`, `VLOOKUP`…).

**Version en ligne : https://paulbachmann13.github.io/grillon/**

- Ouvre et enregistre les fichiers **.xlsx** (et .csv) ;
- fonctionne **hors ligne** et s'installe comme une application (PWA) sur Mac, iPad, iPhone, Windows ;
- vos fichiers restent sur votre appareil : rien n'est envoyé sur un serveur ;
- sauvegarde automatique locale : le classeur en cours est restauré à la réouverture.

## Installer l'application

- **Chrome / Edge (Mac, Windows)** : ouvrir le lien, puis l'icône « Installer » dans la barre d'adresse.
- **Safari (Mac)** : Fichier → Ajouter au Dock.
- **iPad / iPhone** : Safari → Partager → « Sur l'écran d'accueil ».

Une fois installée, l'application fonctionne sans connexion.

## Utilisation

| Action | Comment |
|---|---|
| Ouvrir un .xlsx / .csv | Fichier → Ouvrir… (Ctrl/⌘+O), ou glisser-déposer le fichier sur la fenêtre |
| Enregistrer | Fichier → Enregistrer (Ctrl/⌘+S). Sur Chrome/Edge, le fichier d'origine est réécrit ; sur Safari/iOS, le fichier est téléchargé. |
| Enregistrer sous | Fichier → Enregistrer sous… (Ctrl/⌘+Maj+S) |
| Ajouter un CSV comme nouvelle feuille | Fichier → Importer un CSV… |
| Exporter la feuille en CSV | Fichier → Exporter la feuille en CSV (séparateur « ; », virgule décimale, UTF-8 avec BOM, comme Excel en français) |

Voir [docs/FONCTIONNALITES.md](docs/FONCTIONNALITES.md) pour la liste des fonctionnalités et des limites connues, et [docs/FONCTIONS.md](docs/FONCTIONS.md) pour les fonctions de formule.

## Développement

Prérequis : Node.js 20+.

```bash
npm install
npm run dev          # http://localhost:5173
npm test             # tests unitaires (Vitest)
npm run fixtures     # génère les classeurs de test dans tests/fixtures/out
npm run test:e2e     # tests Playwright sur le build de production
npm run build        # build de production dans dist/
```

Le déploiement sur GitHub Pages est automatique à chaque push sur `main` (GitHub Actions),
après passage des tests.

## Choix techniques et licences

- **Grille, ruban, barre de formule et moteur de formules : [Univer](https://github.com/dream-num/univer)**,
  uniquement les paquets open source (Apache-2.0). Les plugins « Pro » payants ne sont pas utilisés.
- **Formules : moteur intégré d'Univer (Apache-2.0)** plutôt que HyperFormula.
  HyperFormula est distribué sous GPLv3 (ou licence commerciale) ; le moteur d'Univer est plus permissif,
  déjà intégré à la barre de formule (autocomplétion, aide aux arguments) et calcule dans un Web Worker.
  Les fonctions Excel manquantes seront ajoutées en fonctions personnalisées.
- **Import/export .xlsx : convertisseur maison** sur [ExcelJS](https://github.com/exceljs/exceljs) (MIT),
  couvert par des tests de round-trip et une vérification d'ouverture dans LibreOffice.
- **CSV** : PapaParse (MIT). **PWA** : vite-plugin-pwa / Workbox (MIT).

Détails : [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md).

## Licence

Apache-2.0.
