# Fonctionnalités et limites connues

État à la fin de la **phase 5** (version 1.0).

## Fonctionnalités

### Interface
- Ruban en français, en grille (comme Excel) : Démarrer, Insérer, Formules, Données, Vue ; menu **Fichier**.
- Barre de formule, grille, onglets de feuilles (ajouter, renommer, supprimer, réordonner, masquer, couleur).
- Nombres au format Standard affichés à la française (« 1,5 ») et booléens en « VRAI »/« FAUX ».
- Grille qui s'agrandit d'elle-même quand on défile ou qu'on déplace la sélection vers le bord
  (par 1 000 lignes et 26 colonnes), jusqu'aux limites d'Excel : 1 048 576 lignes, 16 384 colonnes (XFD).
- Texte trop long **coupé au bord de la cellule** par défaut ; Vue → **Texte qui déborde** pour
  le laisser déborder sur les cellules vides comme dans Excel (réglage d'affichage mémorisé sur l'appareil).

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

### Onglet Données
| Groupe Excel | Dans Grillon |
|---|---|
| Récupérer des données | Importer un CSV dans une nouvelle feuille |
| Trier et filtrer | Tri simple et multi-niveaux (tri personnalisé), filtre automatique (valeurs, conditions, couleurs), **filtre avancé** (zone de critères, sur place ou copie, sans doublon), Afficher tout |
| Outils de données | **Convertir** (texte en colonnes), **Remplissage instantané** (Ctrl+E), **Supprimer les doublons**, Validation des données (liste, nombre, date, longueur de texte, formule, messages) |
| Plan | **Grouper / Dissocier** lignes et colonnes, **Masquer / Afficher le détail**, **Sous-total** |
| Tableaux | **Mettre sous forme de tableau**, ligne d'en-tête, **ligne de totaux**, références structurées (`Ventes[Prix]`), filtres de tableau |

### Formules et analyse
- Barre de formule avec **autocomplétion**, **aide sur les arguments**, **coloration des références**, **F4** (références absolues), insertion de fonction par catégorie (Univer).
- **Gestionnaire de noms**, **Repérer les antécédents / les dépendants**.
- ≈ 520 fonctions Excel (noms anglais), matrices dynamiques, LAMBDA/LET ; liste des fonctions absentes : [FONCTIONS.md](FONCTIONS.md).
- **Graphiques** : histogramme, barres, courbes, aires (groupés/empilés), secteurs, nuage de points, combiné, axe secondaire ; titre, légende, titres d'axes, couleurs ; mise à jour en direct.
- **Tableau croisé dynamique** (lignes, colonnes, valeurs somme/nombre/moyenne/min/max, filtres, totaux, actualisation).
- **Valeur cible**.

### Fichiers
- **Import/export .xlsx** : valeurs, formules (y compris partagées), résultats en cache, types
  (nombre, texte, date, booléen, erreurs), texte enrichi, largeurs/hauteurs, lignes/colonnes masquées,
  cellules fusionnées, styles (police, taille, gras/italique/souligné/barré, couleurs de texte et de
  fond, couleurs de thème, bordures, alignements, retour à la ligne, rotation, retrait), formats de
  nombre, plusieurs feuilles, feuilles masquées, couleur d'onglet, quadrillage masqué, volets figés,
  noms définis, liens, **mises en forme conditionnelles**, **validations de données**, **filtres
  automatiques avec critères**, **tableaux structurés** (style, ligne de totaux), **plan** (groupes),
  **graphiques**, **images**, **notes**.
- **CSV** : ouverture ou import dans une nouvelle feuille avec un **assistant** (comme « À partir
  d'un fichier texte/CSV » d'Excel) : origine du fichier (UTF-8 ou Windows/ANSI, détectée),
  séparateur (; , tabulation, |, espace, autre), séparateur décimal, ligne d'en-têtes (gras et figée),
  aperçu ; export de la feuille active au format Excel français.
- **Ouvrir/enregistrer** : File System Access API (Chrome/Edge : Ctrl+S réécrit le fichier),
  sinon sélecteur de fichier + téléchargement (Safari, iOS, Firefox). Glisser-déposer.
- **Sauvegarde automatique** dans IndexedDB, restauration à la réouverture.

### Affichage, impression, raccourcis
- Vue : figer les volets, zoom, quadrillage, plein écran (Univer), **mode sombre**, **texte qui déborde**.
- **Impression / export PDF** (Ctrl/⌘+P) : feuille ou sélection, portrait/paysage, quadrillage, ajustement à la largeur.
- Insérer : images, liens, notes, listes déroulantes, graphiques, tableau croisé dynamique.
- Raccourcis Excel : Ctrl+C/V/X, Ctrl+Z/Y, Ctrl+B/I/U, F2, F4, Ctrl+flèches, Ctrl+Maj+flèches, Ctrl+D (recopier vers le bas), Alt+= (somme auto), Ctrl+E, Ctrl+; (date), Ctrl+: (heure), Ctrl+F (rechercher), Ctrl+S/O/P.

### PWA
- Installable, fonctionne hors ligne (tout est mis en cache, y compris le convertisseur xlsx).
- Une fois installée sur Chrome/Edge, Grillon peut ouvrir les .xlsx depuis le système.

### Performance
- 50 000 lignes × 6 colonnes : ouverture ≈ 2 s, enregistrement ≈ 2 s ; 100 000 lignes : ≈ 7 s chacun (Chrome, PC portable).
- Lecture/écriture des .xlsx et recalcul des formules dans des Web Workers : l'interface ne se fige pas ; grille virtualisée (défilement fluide).

## Limites connues

| Limite | Contournement / plan |
|---|---|
| Notes : texte seul (pas de mise en forme ni de taille de bulle) ; les commentaires en fil de discussion des versions récentes d'Excel ne sont pas importés (seules les notes le sont). Formes (rectangles, flèches…) non prises en charge. | — |
| Graphiques : pas de graphiques 3D, radar, boursiers, en cascade, entonnoir, cartes ; un graphique Excel de ce type n'est pas importé. Étiquettes de données, courbes de tendance et mises en forme fines (dégradés, ombres) ne sont pas conservées. Les secteurs reprennent la palette par défaut. | Alternative : ECharts couvre l'essentiel ; types supplémentaires possibles plus tard. |
| Tableau croisé dynamique : calculé par Grillon et écrit en valeurs. Dans le .xlsx il devient un tableau de valeurs (pas un TCD Excel) ; sa définition est conservée dans le classeur Grillon (sauvegarde automatique) mais pas dans le .xlsx. Pas de glisser-déposer des champs, pas de groupement de dates. Les TCD des fichiers Excel sont importés en valeurs. | Implémentation maison, Univer Pro étant payant. |
| Sparklines et « Analyse rapide » non disponibles. | Optionnels ; à envisager plus tard. |
| Audit des formules : liste des antécédents/dépendants, sans flèches dessinées sur la grille. | — |
| Mise en forme conditionnelle : les règles « doublons », « commence par », « se termine par », « ne contient pas » sont écrites sous forme de règle formule équivalente (relues correctement par Grillon). | Effet identique dans Excel. |
| Mise en forme conditionnelle : un seuil de barre de données ou de jeu d'icônes défini par une formule est remplacé par sa valeur numérique. | — |
| Les noms définis par une formule (et non une plage) ne sont pas importés. | Limite d'ExcelJS. |
| Motifs de remplissage (hachures) et dégradés : remplacés par une couleur unie. | Univer n'affiche que des fonds unis. |
| Les formats de date proposés dans la liste rapide sont au format ISO (2026-10-14) ; le format français jj/mm/aaaa est dans « Formats ». Les dates des fichiers Excel gardent leur format. | — |
| Police par défaut Calibri : si elle n'est pas installée (iPad), le navigateur la remplace. | Affichage légèrement différent, contenu identique. |
| Plan : pas de boutons +/− ni de niveaux 1-2-3 dans la marge de la grille (Univer ne les propose pas) ; on utilise Données → Plan → Masquer/Afficher le détail. | — |
| Filtres : les filtres par couleur, « 10 premiers » et dynamiques d'un fichier Excel ne sont pas importés (le filtre lui-même l'est). Les filtres propres à un tableau structuré ne sont pas enregistrés dans le .xlsx. | — |
| Tableaux : la ligne de totaux est une ligne ordinaire sous le tableau dans Grillon (Univer ne gère pas encore les totaux) ; elle est réécrite comme vraie ligne de totaux dans le .xlsx. Les styles de tableau d'Excel sont conservés à l'export mais affichés avec les styles Univer. | — |
| Validation : le type « case à cocher » d'Univer n'existe pas dans Excel et n'est pas exporté. | — |
| Les fichiers protégés par mot de passe, .xls (ancien format) et .xlsb ne sont pas pris en charge. | Les réenregistrer en .xlsx depuis Excel/LibreOffice. |
| Sur Safari/iOS, « Enregistrer » télécharge une copie (pas de réécriture du fichier d'origine). | Limite du navigateur ; une app Mac (Tauri) le permettrait (voir README). |
| Impression : mise en page simple (pas d'en-têtes/pieds de page, de zones d'impression ni de sauts de page manuels) ; la mise en page d'impression des fichiers Excel n'est ni lue ni conservée. | — |
| iPhone : ruban complet, défilable horizontalement ; pas d'interface dédiée au téléphone. | Version simplifiée suffisante pour consulter et saisir. |
| Le collage depuis le presse-papiers demande l'autorisation du navigateur la première fois. | — |
