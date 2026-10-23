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
| Styles | Mise en forme conditionnelle : règles de cellules, texte, dates, doublons, top/bottom, moyenne, formule, barres de données, nuances de couleurs, jeux d'icônes ; gestionnaire de règles ; **Mettre sous forme de tableau** (galerie des 21 styles d'Excel : Clair, Moyen, Foncé) ; **Styles de cellule** |
| Cellules | Insérer / Supprimer cellules, lignes, colonnes, feuille ; Format : ajustement automatique hauteur/largeur, masquer, renommer |
| Édition | **Somme automatique** (Somme, Moyenne, Nb, Max, Min), **Remplir** (bas, droite, haut, gauche), Effacer (tout, formats, contenu), Trier (A→Z, Z→A, personnalisé), Filtrer, Rechercher/Remplacer |
| — | Annuler / Rétablir sans limite |

### Onglet Données
| Groupe Excel | Dans Grillon |
|---|---|
| Récupérer des données | Importer un CSV dans une nouvelle feuille |
| Trier et filtrer | Tri simple et multi-niveaux (tri personnalisé), filtre automatique (valeurs, conditions, couleurs), **filtre avancé** (zone de critères, sur place ou copie, sans doublon), Afficher tout |
| Outils de données | **Convertir** (texte en colonnes), **Remplissage instantané** (Ctrl+E), **Supprimer les doublons**, Validation des données (liste, nombre, date, longueur de texte, formule, messages) |
| Plan | **Grouper / Dissocier** lignes et colonnes, **boutons 1 2 3 et − / + dans la marge** (comme Excel), **Masquer / Afficher le détail**, **Sous-total** |
| Tableaux | **Mettre sous forme de tableau** (galerie de styles, « Mon tableau comporte des en-têtes »), **Style du tableau**, ligne d'en-tête, **ligne de totaux**, références structurées (`Ventes[Prix]`), filtres de tableau |

### Formules et analyse
- **Saisie à la française** : « 1 234,5 », « -12,50 € », « 1.5 », « 3/4 » (3 avril), « 31/12 », VRAI / FAUX sont reconnus comme dans Excel en français.
- **Formules en français** (`=SOMME(A1:A3;10)`, `=SI(…;…;…)`) : saisie, barre de formule, autocomplétion et aide en français ; le classeur et le .xlsx gardent les noms anglais.
- Barre de formule avec **autocomplétion**, **aide sur les arguments**, **coloration des références**, **F4** (références absolues), insertion de fonction par catégorie (Univer).
- **Somme automatique** (Σ), **Afficher les formules** (aussi dans Vue, Ctrl+`).
- **Gestionnaire de noms**, **Repérer les antécédents / les dépendants**.
- **Vérification des erreurs** (#DIV/0!, #N/A, #NOM?…, nombres stockés sous forme de texte : convertir, ignorer, précédent / suivant) et **Évaluer la formule** (calcul pas à pas, partie soulignée, comme Excel). Les erreurs s'affichent avec les codes français d'Excel (#NOM?, #VALEUR!, #NOMBRE!…).
- **Écran d'accueil** (toujours au démarrage, et Fichier → Accueil ; les cinq derniers travaux non enregistrés sont proposés en « Reprendre ») : Nouveau, Ouvrir, fichiers récents, Paramètres (apparence, mises à jour) et numéro de version.
- **Mises à jour automatiques** (application Mac) : Grillon propose d'installer les nouvelles versions publiées sur GitHub ; Fichier → Rechercher les mises à jour….
- **Fenêtre Espion** : surveille la valeur et la formule de cellules (même d'une autre feuille) ; liste gardée pour la session, non enregistrée dans le .xlsx.
- **Segments et chronologies** (Insertion) : filtres à boutons pour un tableau ou un tableau croisé dynamique ; plusieurs se combinent.
- **Sauts de page** (Vue) : insérer / supprimer / rétablir les sauts manuels, aperçu des sauts de page, impression en plusieurs bandes de colonnes ; enregistrés dans le .xlsx.
- **Thèmes** (Vue → Thèmes, Couleurs, Polices) : 8 thèmes, couleurs et polices personnalisables ; le thème est enregistré dans le .xlsx. Limites : voir ci-dessous.
- **Options de calcul** : Automatique ou **Manuel**, **Calculer maintenant** (F9).
- ≈ 520 fonctions Excel (noms anglais), matrices dynamiques, LAMBDA/LET ; liste des fonctions absentes : [FONCTIONS.md](FONCTIONS.md).
- **Graphiques** (comme Insérer → Graphiques d'Excel) : boîte « Tous les graphiques » avec aperçu en direct ; histogramme, barres, courbes, aires (groupés, empilés, 100 % empilés), secteurs, anneau, nuage de points (marqueurs, lignes droites ou lissées), bulles, radar, combinés (courbe sur axe secondaire), **cascade** (totaux), **entonnoir**, **histogramme de répartition** (largeur ou nombre de classes, dépassement, soupassement) et **boîte à moustaches** (moyenne, valeurs aberrantes).
  Réglages de la boîte « Format du graphique » : **Données** (plages des séries, ajouter / supprimer / réordonner, intervertir lignes et colonnes, cellules masquées, cellules vides) ; **Éléments** (titre, titres d'axes, légende, étiquettes de données avec valeur / catégorie / série / pourcentage, position, format de nombre et séparateur, quadrillage principal et secondaire, tableau de données) ; **Axes** (minimum, maximum, unité, échelle logarithmique, ordre inverse, format de nombre, graduations, angle des étiquettes, axe masqué) ; **Séries** (couleur, type, axe secondaire, épaisseur et style de ligne, marqueurs, lissage, opacité, étiquettes propres, **courbes de tendance** linéaire / exponentielle / logarithmique / polynomiale / puissance / moyenne mobile avec équation, R², prévision et rétrospective, **barres d'erreur** fixes / pourcentage / écart type) ; **Mise en forme** (palette, arrière-plan, bordure, taille du texte, espacement et chevauchement, trou de l'anneau, angle et éclatement des secteurs, taille des bulles).
  Tout est enregistré dans le .xlsx (DrawingML ; cascade, entonnoir, histogramme et boîte à moustaches au format « chartEx » d'Excel 2016) et relu.
- **Styles de cellule** (Accueil) : la galerie d'Excel (Satisfaisant, Neutre, Calcul, Entrée, Titres, Total, accents, formats de nombre).
- **Format de cellule** (Ctrl+1 / ⌘1, ou Accueil → Format) : Nombre (catégories d'Excel, aperçu), Alignement, Police, Bordure (contour, intérieur, côté par côté), Remplissage.
- **Protéger la feuille** (Accueil → Format, ou Fichier) : mot de passe facultatif, haché comme Excel ; **cellules déverrouillées** (Format de cellule → Protection) qui restent modifiables ; écrit et relu dans le .xlsx.
- **Mini-graphiques** (Insérer → Mini-graphique) : courbes, histogramme, positif/négatif ; points haut, bas, premier, dernier, négatifs ; enregistrés dans le .xlsx (extension Excel 2010) et relus.
- **Formes et zones de texte** (Insérer → Formes) : rectangle, coins arrondis, ellipse, ligne, flèche, zone de texte ; texte, remplissage, contour, police ; enregistrées dans le .xlsx.
- **Analyse rapide** : bouton ⚡ sous la sélection (mise en forme, graphiques, totaux, tableaux, mini-graphiques).
- **Symbole** (Insérer) : grille de symboles par catégorie.
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
- **Mise en page** (Fichier, Vue) : papier (A4, A3, A5, Lettre, Légal), orientation, échelle ou ajustement en largeur, marges (normales, étroites, larges, personnalisées), centrage, **zone d'impression**, **lignes à répéter**, quadrillage, en-têtes de lignes et de colonnes, **en-têtes et pieds de page** (&P, &N, &D, &F, &A). Réglages par feuille, lus et écrits dans le .xlsx.
- **Impression / export PDF** (Ctrl/⌘+P) : réglages et **aperçu des pages** côte à côte (comme Excel) ; feuille ou sélection, portrait/paysage, quadrillage, ajustement à la largeur.
- Insérer : images, liens, notes, listes déroulantes, graphiques, tableau croisé dynamique.
- Raccourcis Excel : Ctrl+C/V/X, Ctrl+Z/Y, Ctrl+B/I/U, F2, F4, Ctrl+flèches, Ctrl+Maj+flèches, Ctrl+Origine/Fin, Ctrl+Espace / Maj+Espace (colonne / ligne), Ctrl+D / Ctrl+R (recopier), **Alt+=** (somme automatique), Ctrl+E, Ctrl+; (date), Ctrl+: (heure), Ctrl+Maj+1 à 6 (formats nombre, heure, date, monétaire, %, scientifique), Ctrl+Maj++ / Ctrl+- (insérer / supprimer des cellules), Ctrl+` (afficher les formules), Ctrl+Maj+L (filtre), Ctrl+F (rechercher), Ctrl+S/O/P, Ctrl+1 / ⌘1 (Format de cellule), **F9** (calculer).

### Application Mac
- .dmg universel (Intel et Apple Silicon), non signé ; Cmd+S réécrit le fichier d'origine, double-clic dans le Finder, boîtes natives. **Plusieurs fenêtres** : Nouveau et Ouvrir ouvrent un classeur par fenêtre, comme Excel. Voir le README.

### PWA
- Installable, fonctionne hors ligne (tout est mis en cache, y compris le convertisseur xlsx).
- Une fois installée sur Chrome/Edge, Grillon peut ouvrir les .xlsx depuis le système.

### Performance
- 50 000 lignes × 6 colonnes : ouverture ≈ 2 s, enregistrement ≈ 2 s ; 100 000 lignes : ≈ 7 s chacun (Chrome, PC portable).
- Lecture/écriture des .xlsx et recalcul des formules dans des Web Workers : l'interface ne se fige pas ; grille virtualisée (défilement fluide).

## Limites connues

| Limite | Contournement / plan |
|---|---|
| Formules en français : les formules des règles de mise en forme conditionnelle, de la validation des données et des noms définis, la boîte « Évaluer la formule » et l'option « Afficher les formules » restent en anglais (noms anglais, virgule). Une fonction récente sans nom français connu s'écrit en anglais. Comme dans Excel en français, `TRIM` désigne le taux de rendement interne modifié (MIRR) ; la fonction anglaise TRIM s'écrit `SUPPRESPACE`. | Utiliser les noms anglais dans ces boîtes. |
| Segments et chronologies : ils ne sont pas écrits dans le .xlsx (un fichier ouvert dans Excel garde les lignes masquées, sans segment) et les segments d'un fichier Excel ne sont pas importés. Ils filtrent en masquant des lignes : les lignes masquées à la main dans la plage filtrée sont réaffichées quand le segment les retient. Dans un tableau croisé, les éléments des champs de dates sont les valeurs brutes de la source. Pas de plusieurs colonnes de boutons ni de styles de segment. | Utiliser les filtres automatiques pour un fichier destiné à Excel. |
| Sauts de page : l'aperçu des sauts de page n'a pas de glisser-déposer des lignes bleues (utiliser Insérer / Supprimer) ni d'ordre des pages « vers la droite, puis vers le bas » ; le découpage automatique est calculé par Grillon (hauteurs de ligne et largeurs de colonne) et peut différer d'une page de celui d'Excel. Insérer / supprimer un saut ne s'annule pas avec Ctrl+Z. | — |
| Thèmes : Univer ne garde pas le lien d'une cellule avec une couleur de thème. Changer de thème recolore les couleurs égales à celles du thème actuel (et leurs teintes standard), pas les couleurs choisies à la main ; les règles de mise en forme conditionnelle, les graphiques et les textes enrichis ne sont pas recolorés. Le changement recharge le classeur : Ctrl+Z ne l'annule pas (réappliquer l'ancien thème rétablit les couleurs). Le sélecteur de couleurs n'affiche pas la palette du thème. Pas d'effets de thème. | Réappliquer le thème précédent. |
| Notes : texte seul (pas de mise en forme ni de taille de bulle) ; les commentaires en fil de discussion des versions récentes d'Excel ne sont pas importés (seules les notes le sont). Formes : seules les formes simples (rectangle, coins arrondis, ellipse, ligne, flèche, zone de texte) sont prises en charge ; les autres, SmartArt, WordArt et images en forme ne sont pas importés. | — |
| Graphiques : pas de graphiques 3D, boursiers, cartes, surface, Pareto ni sunburst/treemap ; un graphique Excel de ce type n'est pas importé. Cascade, entonnoir, histogramme et boîte à moustaches : pas de couleur par point ni de lignes de connexion de la cascade ; LibreOffice ne sait pas les afficher (le reste du fichier s'ouvre). Les courbes de tendance et barres d'erreur ne sont pas proposées sur les barres horizontales ; les barres d'erreur sont verticales ; étiquettes d'axe « en haut / en bas » non disponibles ; dégradés, ombres, étiquettes personnalisées par point et titres liés à une cellule ne sont pas conservés. Le tableau de données n'est pas imprimé. Le rendu (ECharts) diffère légèrement de celui d'Excel. | Alternative : ECharts couvre l'essentiel ; types supplémentaires possibles plus tard. |
| Tableau croisé dynamique : un seul par feuille dans Grillon (les suivants d'un fichier Excel restent en valeurs) ; pas de glisser-déposer des champs, de regroupement de dates ni de « afficher les valeurs en » ; mise en page tabulaire (Excel la recalcule à l'ouverture). Il est enregistré comme un vrai tableau croisé Excel et relu. | Implémentation maison, Univer Pro étant payant. |
| Mini-graphiques : le mini-graphique couvre la cellule sauf une bande à gauche (cliquer dans cette bande sélectionne la cellule) ; pas d'axe horizontal ni de groupes de réglages communs (chaque cellule garde les siens) ; non imprimés. Formes et mini-graphiques ne sont pas imprimés. | — |
| Orthographe (Révision) et vérification d'accessibilité : non disponibles (aucun dictionnaire libre léger intégré). | Le correcteur du navigateur s'applique aux zones de texte de l'interface. |
| Audit des formules : liste des antécédents/dépendants, sans flèches dessinées sur la grille ; pas de triangles verts dans les cellules (la Vérification des erreurs les liste). Évaluer la formule calcule aussi la branche non retenue d'un SI. Calcul manuel : non enregistré dans le fichier (le classeur rouvre en automatique), et une formule saisie n'est calculée qu'avec F9. | — |
| Mise en forme conditionnelle : les règles « doublons », « commence par », « se termine par », « ne contient pas » sont écrites sous forme de règle formule équivalente (relues correctement par Grillon). | Effet identique dans Excel. |
| Mise en forme conditionnelle : un seuil de barre de données ou de jeu d'icônes défini par une formule est remplacé par sa valeur numérique. | — |
| Les noms définis par une formule (et non une plage) ne sont pas importés. | Limite d'ExcelJS. |
| Motifs de remplissage (hachures) et dégradés : remplacés par une couleur unie. | Univer n'affiche que des fonds unis. |
| Les formats de date proposés dans la liste rapide sont au format ISO (2026-10-14) ; le format français jj/mm/aaaa est dans « Formats ». Les dates des fichiers Excel gardent leur format. | — |
| Police par défaut Calibri : si elle n'est pas installée (iPad), le navigateur la remplace. | Affichage légèrement différent, contenu identique. |
| Filtres : les filtres par couleur, « 10 premiers » et dynamiques d'un fichier Excel ne sont pas importés (le filtre lui-même l'est). Les filtres propres à un tableau structuré ne sont pas enregistrés dans le .xlsx. | — |
| Tableaux : la ligne de totaux est une ligne ordinaire sous le tableau dans Grillon (Univer ne gère pas encore les totaux) ; elle est réécrite comme vraie ligne de totaux dans le .xlsx. Les styles de tableau d'Excel sont affichés avec le style le plus proche de la galerie (couleur et famille) et le nom d'origine est réécrit tel quel tant qu'on ne change pas de style. | — |
| Validation : le type « case à cocher » d'Univer n'existe pas dans Excel et n'est pas exporté. | — |
| Les fichiers protégés par mot de passe, .xls (ancien format) et .xlsb ne sont pas pris en charge. | Les réenregistrer en .xlsx depuis Excel/LibreOffice. |
| Plusieurs fenêtres : uniquement dans l'application Mac ; dans le navigateur, ouvrir un autre fichier remplace le classeur courant (ouvrir Grillon dans un autre onglet pour en avoir deux). Les fenêtres supplémentaires n'ont pas de sauvegarde automatique (confirmation avant fermeture). | — |
| Sur Safari/iOS, « Enregistrer » télécharge une copie (pas de réécriture du fichier d'origine). | Limite du navigateur ; une app Mac (Tauri) le permettrait (voir README). |
| Impression : l'aperçu montre le contenu des pages, sans les en-têtes et pieds de page ni les lignes à répéter ; le découpage final est celui du navigateur. Les en-têtes et pieds de page (et les numéros de page) s'impriment avec Chrome et Edge (marges de page CSS) ; Safari et Firefox les ignorent. Pas de sauts de page manuels ; « centrer verticalement » est enregistré dans le fichier mais pas appliqué à l'impression. | — |
| iPhone : ruban complet, défilable horizontalement ; pas d'interface dédiée au téléphone. | Version simplifiée suffisante pour consulter et saisir. |
| Le collage depuis le presse-papiers demande l'autorisation du navigateur la première fois. | — |
