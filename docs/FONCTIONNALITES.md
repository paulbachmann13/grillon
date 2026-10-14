# Fonctionnalités et limites connues

État à la fin de la **phase 1** (fondations + import/export).

## Fonctionnalités

### Interface
- Ruban en français : Démarrer, Insérer, Formules, Données, Vue ; menu **Fichier**.
- Barre de formule, grille, onglets de feuilles (ajouter, renommer, supprimer, réordonner, masquer, couleur).
- Nombres au format Standard affichés à la française (« 1,5 ») et booléens en « VRAI »/« FAUX ».
- Annuler/rétablir, gras/italique/souligné/barré, polices, couleurs, bordures, alignements,
  fusion, formats de nombre (fournis par Univer).
- Mise en forme conditionnelle, validation de données, filtres, tri, rechercher/remplacer,
  notes, liens hypertexte, images : présents dans l'interface (plugins Univer open source).
  Leur **enregistrement dans le .xlsx** arrive aux phases 2 et 3.

### Fichiers
- **Import .xlsx** : valeurs, formules (y compris partagées), résultats en cache, types
  (nombre, texte, date, booléen, erreurs), largeurs/hauteurs, lignes/colonnes masquées,
  cellules fusionnées, styles (police, taille, gras/italique/souligné/barré, couleurs de texte
  et de fond, couleurs de thème, bordures, alignements, retour à la ligne, rotation, retrait),
  formats de nombre/date/pourcentage/monétaire/fraction/scientifique, plusieurs feuilles,
  feuilles masquées, couleur d'onglet, quadrillage masqué, volets figés, noms définis, liens.
- **Export .xlsx** avec les mêmes éléments.
- **CSV** : ouverture (séparateur et décimales détectés), import dans une nouvelle feuille,
  export de la feuille active au format Excel français.
- **Ouvrir/enregistrer** : File System Access API (Chrome/Edge : Ctrl+S réécrit le fichier),
  sinon sélecteur de fichier + téléchargement (Safari, iOS, Firefox). Glisser-déposer.
- **Sauvegarde automatique** dans IndexedDB, restauration à la réouverture.

### PWA
- Installable, fonctionne hors ligne (tout est mis en cache, y compris le convertisseur xlsx).
- Une fois installée sur Chrome/Edge, Grillon peut ouvrir les .xlsx depuis le système.

### Performance
- 50 000 lignes × 6 colonnes : ouverture ≈ 2 s, enregistrement ≈ 2 s (Chrome, PC portable).
- Recalcul des formules dans un Web Worker (l'interface ne se fige pas).

## Limites connues

| Limite | Contournement / plan |
|---|---|
| Le texte enrichi (plusieurs styles dans une même cellule) est importé en texte simple. | Phase 2. |
| Mises en forme conditionnelles, validations, filtres, graphiques, commentaires, images ne sont pas encore lus/écrits dans le .xlsx. | Phases 2 à 4. |
| Les noms définis par une formule (et non une plage) ne sont pas importés. | Limite d'ExcelJS ; à traiter en phase 4. |
| Motifs de remplissage (hachures) et dégradés : remplacés par une couleur unie. | Univer n'affiche que des fonds unis. |
| Police par défaut Calibri : si elle n'est pas installée (iPad), le navigateur la remplace. | Affichage légèrement différent, contenu identique. |
| Les fichiers protégés par mot de passe, .xls (ancien format) et .xlsb ne sont pas pris en charge. | Les réenregistrer en .xlsx depuis Excel/LibreOffice. |
| Sur Safari/iOS, « Enregistrer » télécharge une copie (pas de réécriture du fichier d'origine). | Limite du navigateur ; l'app Tauri (phase 5) le permettra sur Mac. |
| Le format Standard affiche jusqu'à 15 chiffres significatifs, sans tenir compte de la largeur de colonne (Excel arrondit selon la largeur). | — |
