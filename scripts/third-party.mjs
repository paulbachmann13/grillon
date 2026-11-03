// Liste des composants utilisés par Grillon et de leurs licences (THIRD-PARTY-LICENSES.md).
// À relancer après un changement de dépendances : node scripts/third-party.mjs
import { readFileSync, writeFileSync } from 'node:fs';

const pkg = JSON.parse(readFileSync('package.json', 'utf8'));
const rows = Object.keys(pkg.dependencies)
  .sort()
  .map((name) => {
    const dep = JSON.parse(readFileSync(`node_modules/${name}/package.json`, 'utf8'));
    const license = typeof dep.license === 'string' ? dep.license : JSON.stringify(dep.license);
    return `| ${name} | ${dep.version} | ${license} |`;
  });

writeFileSync(
  'THIRD-PARTY-LICENSES.md',
  `# Composants utilisés par Grillon

Grillon est fait avec des composants libres et gratuits. Leurs licences s'appliquent à leur code ; le texte complet
de chaque licence est dans le paquet correspondant (\`node_modules/<paquet>/LICENSE\`) et sur sa page npm.
Aucun paquet payant (\`@univerjs-pro/*\`) n'est utilisé. JSZip est utilisé sous licence MIT (au choix MIT ou GPL-3.0).

Grillon est un tableur compatible avec les fichiers Excel ; il n'est ni affilié à Microsoft ni approuvé par Microsoft.
Excel est une marque de Microsoft Corporation.

## Application web

| Paquet | Version | Licence |
| --- | --- | --- |
${rows.join('\n')}

## Application Mac

Tauri 2 et ses modules (dialog, fs, process, updater) : MIT ou Apache-2.0, au choix. Voir \`src-tauri/Cargo.toml\`.

Fichier généré par \`node scripts/third-party.mjs\`.
`,
);
