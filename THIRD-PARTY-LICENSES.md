# Composants utilisés par Grillon

Grillon est fait avec des composants libres et gratuits. Leurs licences s'appliquent à leur code ; le texte complet
de chaque licence est dans le paquet correspondant (`node_modules/<paquet>/LICENSE`) et sur sa page npm.
Aucun paquet payant (`@univerjs-pro/*`) n'est utilisé. JSZip est utilisé sous licence MIT (au choix MIT ou GPL-3.0).

Grillon est un tableur compatible avec les fichiers Excel ; il n'est ni affilié à Microsoft ni approuvé par Microsoft.
Excel est une marque de Microsoft Corporation.

## Application web

| Paquet | Version | Licence |
| --- | --- | --- |
| @fontsource/carlito | 5.3.0 | OFL-1.1 |
| @tauri-apps/api | 2.12.4 | Apache-2.0 OR MIT |
| @tauri-apps/plugin-dialog | 2.8.3 | MIT OR Apache-2.0 |
| @tauri-apps/plugin-fs | 2.7.1 | MIT OR Apache-2.0 |
| @tauri-apps/plugin-process | 2.4.0 | MIT OR Apache-2.0 |
| @tauri-apps/plugin-updater | 2.14.0 | MIT OR Apache-2.0 |
| @univerjs/icons | 1.43.1 | MIT |
| @univerjs/preset-sheets-conditional-formatting | 1.0.4 | Apache-2.0 |
| @univerjs/preset-sheets-core | 1.0.4 | Apache-2.0 |
| @univerjs/preset-sheets-data-validation | 1.0.4 | Apache-2.0 |
| @univerjs/preset-sheets-drawing | 1.0.4 | Apache-2.0 |
| @univerjs/preset-sheets-filter | 1.0.4 | Apache-2.0 |
| @univerjs/preset-sheets-find-replace | 1.0.4 | Apache-2.0 |
| @univerjs/preset-sheets-hyper-link | 1.0.4 | Apache-2.0 |
| @univerjs/preset-sheets-note | 1.0.4 | Apache-2.0 |
| @univerjs/preset-sheets-sort | 1.0.4 | Apache-2.0 |
| @univerjs/preset-sheets-table | 1.0.4 | Apache-2.0 |
| @univerjs/presets | 1.0.4 | Apache-2.0 |
| echarts | 6.1.0 | Apache-2.0 |
| exceljs | 4.4.0 | MIT |
| fast-xml-parser | 5.11.2 | MIT |
| jszip | 3.10.2 | (MIT OR GPL-3.0-or-later) |
| papaparse | 5.7.0 | MIT |
| react | 18.3.1 | MIT |
| react-dom | 18.3.1 | MIT |
| rxjs | 7.8.2 | Apache-2.0 |

## Application Mac

Tauri 2 et ses modules (dialog, fs, process, updater) : MIT ou Apache-2.0, au choix. Voir `src-tauri/Cargo.toml`.

Fichier généré par `node scripts/third-party.mjs`.
