# Fonctionnalités et limites connues

État à la fin de la **phase 2** (onglet Démarrer complet).

## Fonctionnalités

### Interface
- Ruban en français, en grille (comme Excel) : Démarrer, Insérer, Formules, Données, Vue ; menu **Fichier**.
- Barre de formule, grille, onglets de feuilles (ajouter, renommer, supprimer, réordonner, masquer, couleur).
- Nombres au format Standard affichés à la française (« 1,5 ») et booléens en « VRAI »/« FAUX ».

### Onglet Démarrer
| Groupe Excel | Dans Grillon |
|---|---|
| Presse-papiers | Couper, Copier, Coller ; collage spécial : valeurs, formats, **formules**, largeurs de colonnes, sans bordures ; Reproduire la mise en forme |
| Police | Police, taille, agrandir/réduire, gras, italique, souligné, barré, bordures, couleur de remplissage, couleur de police |
| Alignement | Horizontal, vertical, renvoi à la ligne, ajuster au contenu, rotation, fusionner et centrer, **retrait** |
| Nombre | Standard, Nombre, Monétaire, Comptabilité, Pourcentage, Scientifique, Date, Heure, Durée, Texte, formats personnalisés (dont fractions) ; ajouter/retirer des décimales |
| Styles | Mise en forme conditionnelle : règles de cellules, texte, dates, doublons, top/bottom, moyenne, formule, barres de données, nuances de couleurs, jeux d'icônes ; gestionnaire de règles |
| Cellules | Insérer / Supprimer cellules, lignes, colonnes, feuille ; Format : ajustement automatique hauteur/largeur, masquer, renommer |
| Édition | Effacer (tout, formats, contenu), Trier (A→Z, Z→A, personnalisé), Filtrer, Rechercher/Remplacer |
| — | Annuler / Rétablir sans limite |

### Fichiers
- **Import/export .xlsx** : valeurs, formules (y compris partagées), résultats en cache, types
  (nombre, texte, date, booléen, erreurs), texte enrichi, largeurs/hauteurs, lignes/colonnes masquées,
  cellules fusionnées, styles (police, taille, gras/italique/souligné/barré, couleurs de texte et de
  fond, couleurs de thème, bordures, alignements, retour à la ligne, rotation, retrait), formats de
  nombre, plusieurs feuilles, feuilles masquées, couleur d'onglet, quadrillage masqué, volets figés,
  noms définis, liens, **mises en forme conditionnelles**.
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
| Validations de données, filtres, graphiques, commentaires et images ne sont pas encore lus/écrits dans le .xlsx (ils fonctionnent dans l'application). | Phases 3 à 5. |
| Mise en forme conditionnelle : l'option « Interrompre si vrai » n'est ni lue ni écrite (ExcelJS ne la gère pas). Les règles « doublons », « commence par », « se termine par », « ne contient pas » sont écrites sous forme de règle formule équivalente (relues correctement par Grillon). | Effet identique dans Excel. |
| Mise en forme conditionnelle : un seuil de barre de données ou de jeu d'icônes défini par une formule est remplacé par sa valeur numérique. | — |
| Les noms définis par une formule (et non une plage) ne sont pas importés. | Limite d'ExcelJS ; à traiter en phase 4. |
| Motifs de remplissage (hachures) et dégradés : remplacés par une couleur unie. | Univer n'affiche que des fonds unis. |
| Les formats de date proposés dans la liste rapide sont au format ISO (2026-10-14) ; le format français jj/mm/aaaa est dans « Formats ». Les dates des fichiers Excel gardent leur format. | Phase 5 (liste de formats francisée). |
| Police par défaut Calibri : si elle n'est pas installée (iPad), le navigateur la remplace. | Affichage légèrement différent, contenu identique. |
| Les fichiers protégés par mot de passe, .xls (ancien format) et .xlsb ne sont pas pris en charge. | Les réenregistrer en .xlsx depuis Excel/LibreOffice. |
| Sur Safari/iOS, « Enregistrer » télécharge une copie (pas de réécriture du fichier d'origine). | Limite du navigateur ; l'app Tauri (phase 5) le permettra sur Mac. |
| Le collage depuis le presse-papiers demande l'autorisation du navigateur la première fois. | — |
