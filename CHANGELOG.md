# Changelog

Format inspiré de [Keep a Changelog](https://keepachangelog.com/fr/1.1.0/).

## [1.13.3] — 2026-10-22

### Corrigé
- **Formules étendues perdues à l'enregistrement** : quand on étendait une formule (poignée de recopie, Remplir, copier-coller), seule la première cellule gardait sa formule dans le .xlsx ; les autres étaient écrites comme de simples nombres, et à la réouverture les formules avaient disparu. Chaque cellule est maintenant écrite avec sa propre formule (références décalées comme à la recopie). « Afficher les formules » (Ctrl+`) montre aussi la formule des cellules étendues. Les fichiers déjà enregistrés avec ce défaut ne peuvent pas être réparés : les cellules concernées contiennent des nombres, il faut y remettre la formule.
- **Tri avec des formules** : les lignes étaient réordonnées sans adapter les formules, et les résultats devenaient faux (après un tri, « 8 » gardait « =A3*10 » et affichait 30). Comme dans Excel, chaque formule suit maintenant sa ligne ; un seul Ctrl+Z annule le tri et les formules.
- **Formules qui débordent** (SEQUENCE, TRIER, FILTRE…) : à la réouverture, seule la première valeur s'affichait ; le classeur est recalculé à l'ouverture (sauf en calcul manuel).
- **Volets figés** : « Figer la ligne supérieure » (ou la première colonne seule) n'était pas enregistré dans le .xlsx.
- **Lignes masquées vides** : une ligne masquée sans contenu redevenait visible à la réouverture.
- Bordures posées par certains scripts (style en texte) : elles sont maintenant enregistrées.

## [1.13.2]

## [1.13.2] — 2026-10-18

### Modifié
- **Jusqu'à cinq travaux non enregistrés** sont gardés (au lieu d'un seul) et proposés en haut des récents, le plus récent d'abord, chacun avec « Reprendre : nom », sa date et son ✕. Le brouillon unique des versions 1.13.0 et 1.13.1 est repris dans la liste.
- Reprendre un travail pendant qu'un autre, non enregistré, est ouvert ne perd plus ce dernier : il passe dans la liste.

## [1.13.1] — 2026-10-18

### Modifié
- Grillon s'ouvre **toujours sur l'écran d'accueil**, au lieu de rouvrir le dernier classeur. Le travail non enregistré de la dernière session apparaît en haut des récents (« Reprendre : … », ✕ pour le supprimer après confirmation) ; un classeur enregistré se rouvre depuis les récents.

## [1.13.0] — 2026-10-18

### Ajouté
- **Écran d'accueil**, comme celui d'Excel : au démarrage (quand il n'y a rien à reprendre) et par Fichier → Accueil. Menu vert à gauche (Accueil, Nouveau, Ouvrir, Paramètres), tuiles « Classeur vide », « Ouvrir un classeur » et « Importer un fichier CSV », liste des **fichiers récents** (nom, dossier, date ; ✕ pour retirer un fichier) et **numéro de version**.
- **Paramètres** (Fichier → Paramètres) : apparence (automatique, clair, sombre), débordement du texte, recherche des mises à jour au démarrage (désactivable) et « Rechercher maintenant », À propos avec la version.
- Une session déjà enregistrée n'est plus rouverte automatiquement derrière l'accueil ; un travail non enregistré, lui, est toujours repris.

## [1.12.2] — 2026-10-18

### Corrigé
- Mise à jour automatique : quand Grillon est lancé depuis le .dmg ou le dossier Téléchargements, macOS l'exécute en lecture seule (« App Translocation ») et la mise à jour échouait avec « Read-only file system ». Grillon explique maintenant qu'il faut le glisser dans Applications, éjecter le disque et le relancer depuis Applications.

## [1.12.1] — 2026-10-18

### Corrigé
- Barre de formule : les flèches gauche / droite déplacent le curseur dans la formule (comme Excel). Avant, elles déplaçaient la plage de cellules de la formule (`A1:A3` devenait `Y1000`). Maj + flèche sélectionne du texte.

## [1.12.0] — 2026-10-18

### Ajouté
- **Mises à jour automatiques** de l'application Mac : quelques secondes après le démarrage, Grillon regarde s'il existe une Release GitHub plus récente et propose de l'installer (un clic, redémarrage automatique). Aussi dans Fichier → Rechercher les mises à jour…. Les mises à jour sont signées (clé de Grillon, sans rapport avec Apple).
- Chaque Release contient maintenant `latest.json` et l'archive `.app.tar.gz` utilisés par la mise à jour, en plus du .dmg.

**Cette version doit être installée à la main** (.dmg) ; les suivantes se mettront à jour toutes seules.

## [1.11.0] — 2026-10-18

### Ajouté
- **Formules en français**, comme Excel en français : on tape `=SOMME(A1:A10;B1)`, `=SI(A1>10;"grand";"petit")`, `=RECHERCHEV(A2;B:C;2;FAUX)`, avec « ; » entre les arguments et la virgule décimale (`0,196`). Environ 400 fonctions ont leur nom français ; la barre de formule, l'autocomplétion, l'aide de fonction, « Insérer une fonction » et la fenêtre Espion affichent les noms français.
- Une formule saisie en anglais (`=SUM(A1,B1)`) fonctionne toujours. Le classeur et le fichier .xlsx gardent la forme anglaise, donc rien ne change pour Excel.

## [1.10.1] — 2026-10-18

### Corrigé
- Application Mac : un clic droit sur la grille colorait tout le fond en bleu (sélection de texte de la page par WebKit) ; la grille n'est plus sélectionnable comme du texte, les zones de saisie le restent.

## [1.10.0] — 2026-10-18

### Ajouté
- **Segments** et **chronologies** (Insertion) : des boutons posés sur la feuille qui filtrent un tableau (lignes masquées) ou un tableau croisé dynamique (filtre du champ). Clic : un seul élément ; Ctrl / ⌘ / Maj : ajouter ou retirer ; ✕ : effacer le filtre. Les éléments sans donnée (selon les autres filtres) sont grisés.
- Chronologie : années, trimestres, mois ou jours ; un clic choisit une période, Maj étend la plage.
- Plusieurs segments et chronologies se combinent (ET) ; supprimer un segment rend les lignes masquées et retire son filtre du tableau croisé. Les segments sont enregistrés avec le classeur.

## [1.9.0] — 2026-10-18

### Ajouté
- **Sauts de page manuels** (Vue → Sauts de page) : Insérer un saut de page (au-dessus et à gauche de la cellule active, ou ligne / colonne entière), Supprimer le saut de page, Rétablir tous les sauts de page. Lus et écrits dans le .xlsx (`rowBreaks`, `colBreaks`).
- **Aperçu des sauts de page** (Vue) : zone imprimée encadrée, sauts manuels en bleu plein, sauts automatiques en pointillés, « Page n » en filigrane.
- L'impression et l'aperçu avant impression suivent le même découpage en pages : sauts manuels, sauts automatiques, lignes de titre répétées ; une feuille plus large que la page est imprimée en plusieurs bandes de colonnes (de haut en bas, puis de gauche à droite). Comme dans Excel, « Ajuster à la largeur » ignore les sauts manuels.
- L'aperçu avant impression montre le contenu réel de chaque page.

## [1.8.0] — 2026-10-18

### Ajouté
- **Thèmes, Couleurs et Polices** (Vue) : 8 thèmes (Office 2013 – 2022, Office, Office 2007 – 2010 et cinq thèmes Grillon), couleurs et polices personnalisables. Les cellules qui utilisent une couleur du thème (avec ses teintes plus claires / plus foncées) ou la police du thème prennent celles du nouveau thème.
- Le thème est lu dans le .xlsx (`theme1.xml`) et réécrit à l'enregistrement.

### Corrigé
- Les fichiers enregistrés par Grillon contenaient le thème Office 2007 (Cambria/Calibri) au lieu du thème Office 2013 – 2022 : le thème du classeur est maintenant toujours écrit.
- Tests e2e « annuler » et « copier/coller » sous macOS (touches Ctrl/⌘ de Chromium).

## [1.7.0] — 2026-10-17

### Ajouté
- **Fenêtre Espion** (Formules) : surveille la valeur et la formule de cellules choisies, même sur d'autres feuilles ; tri par colonne, double-clic pour atteindre la cellule.

## [1.6.0] — 2026-10-17

### Ajouté
- **Plusieurs fenêtres** (application Mac) : Fichier → Nouveau et Fichier → Ouvrir, ou un double-clic dans le Finder, ouvrent un classeur dans sa propre fenêtre sans fermer celui en cours.
- CLAUDE.md (consignes du projet pour Claude Code) et guide « Développer sur Mac » dans le README.

## [1.5.0] — 2026-10-17

### Ajouté
- **Plan dans la marge**, comme Excel : boutons de niveau 1, 2, 3… et − / + pour replier ou développer chaque groupe (lignes et colonnes).
- **Cellules déverrouillées** (Format de cellule → Protection) : sur une feuille protégée, seules ces cellules restent modifiables ; enregistrées dans le .xlsx.
- **Galerie des styles de tableau** (Mettre sous forme de tableau, aussi dans Accueil) : 21 styles d'Excel, « Mon tableau comporte des en-têtes », changement de style ; le style est écrit et relu dans le .xlsx.
- **Vérification des erreurs** et **Évaluer la formule** (Formules) ; erreurs affichées avec les codes français d'Excel (#NOM?, #VALEUR!…).
- Graphiques **cascade**, **entonnoir**, **histogramme** de répartition et **boîte à moustaches**, enregistrés au format chartEx d'Excel.
- **Aperçu avant impression** dans la boîte Imprimer.
- **Options de calcul** : Automatique / Manuel, **Calculer maintenant** (F9).

## [1.4.0] — 2026-10-17

### Ajouté
- **Vrais tableaux croisés dynamiques dans le .xlsx** : cache, enregistrements et définition (lignes, colonnes, valeurs, filtres) ; Excel les actualise à l'ouverture, LibreOffice les reconnaît. Les tableaux croisés des fichiers Excel sont relus et peuvent être actualisés dans Grillon.
- **Styles de cellule** (galerie d'Excel) et boîte **Format de cellule** (Ctrl+1 / ⌘1) : Nombre, Alignement, Police, Bordure, Remplissage.
- **Protéger la feuille** avec mot de passe facultatif (hachage SHA-512 d'Excel), enregistré dans le .xlsx.
- Application Mac : boîtes de confirmation natives, confirmation avant de fermer un classeur non enregistré, ⌘⇧T (somme automatique).
- docs/EXCEL-COMPARAISON.md : inventaire du ruban d'Excel et état de Grillon.

### Corrigé
- Police par défaut Calibri 11 (comme Excel) au lieu d'Arial.
- Un tableau croisé ne remplace plus la mise en page de sa feuille.

## [1.3.0] — 2026-10-15

### Ajouté
- **Graphiques façon Excel** : boîte « Tous les graphiques » (sous-types, aperçu en direct) et boîte « Format du graphique » à six onglets (Type, Données, Éléments, Axes, Séries, Mise en forme). Nouveaux types : anneau, bulles, radar, courbes empilées, secteurs, nuage lissé… ; étiquettes de données, quadrillage secondaire, tableau de données, axes (minimum, maximum, unité, logarithmique, ordre inverse, format de nombre), courbes de tendance (équation, R²), barres d'erreur, marqueurs, lissage, couleurs de points, palettes, arrière-plan, taille du texte. Tout est enregistré dans le .xlsx et relu.
- **Mise en page** (Fichier, Vue) : papier, orientation, échelle / ajustement, marges, centrage, zone d'impression, lignes à répéter, quadrillage, en-têtes et pieds de page ; appliquée à l'impression et lue / écrite dans le .xlsx.
- **Mini-graphiques** (courbes, histogramme, positif/négatif), **formes et zones de texte**, **Analyse rapide** et **Symbole**.

### Corrigé
- Les cellules de lignes masquées ne sont plus tracées dans les graphiques (comme Excel).

## [1.2.0] — 2026-10-15

### Ajouté
- **Somme automatique** (Démarrer et Formules, **Alt+=**) : Somme, Moyenne, Nb, Max, Min ; comme dans Excel, additionne les nombres au-dessus (sinon à gauche), ou ajoute un total sous chaque colonne d'une sélection.
- **Remplir** vers le bas, la droite, le haut et la gauche.
- **Afficher les formules** (Formules, Vue, Ctrl+`).
- Raccourcis : Ctrl+Espace / Maj+Espace (colonne / ligne entière), Ctrl+Origine / Ctrl+Fin, Ctrl+Maj+1 à 6 (formats), Ctrl+Maj++ et Ctrl+- (boîte Insérer / Supprimer des cellules).

### Corrigé
- La documentation annonçait Alt+= alors qu'il ne marchait pas : c'est maintenant le cas.

## [1.1.0] — 2026-10-15

### Ajouté
- **Grille extensible** : au-delà de la ligne 1000 et de la colonne Z, des lignes et des colonnes s'ajoutent quand on défile ou qu'on déplace la sélection, jusqu'à 1 048 576 lignes et 16 384 colonnes (XFD) comme Excel.
- **Assistant d'import CSV** (Fichier → Ouvrir / Importer un CSV, Données → À partir d'un CSV) : origine du fichier (UTF-8 ou Windows/ANSI), séparateur, séparateur décimal, première ligne en en-têtes (gras et figée), aperçu.
- Vue → **Texte qui déborde** : par défaut, un texte trop long est coupé au bord de sa cellule ; ce bouton rétablit le débordement sur les cellules vides (comme Excel).

- **Application Mac** (Tauri) : .dmg universel construit par GitHub Actions (workflow « Application Mac »), Cmd+S réécrit le fichier d'origine, ouverture par double-clic dans le Finder. Non signée.

### Corrigé
- Les CSV enregistrés par Excel sous Windows (encodage ANSI) s'ouvrent avec leurs accents.

## [1.0.0] — Phase 5 : finitions — 2026-10-15

### Ajouté
- **Mode sombre** (Vue → Mode sombre ; suit le réglage du système par défaut, choix mémorisé).
- **Impression / export PDF** (Fichier → Imprimer, Ctrl/⌘+P, ou Vue) : feuille active ou sélection, portrait/paysage, quadrillage, ajustement à la largeur ; mise en forme, fusions, graphiques et images inclus. Le PDF s'obtient avec « Enregistrer au format PDF ».
- **Images** et **notes** (commentaires) enregistrées dans le .xlsx et relues depuis Excel.
- Raccourcis **Ctrl+;** (date du jour) et **Ctrl+:** (heure) ; les autres raccourcis Excel (Ctrl+C/V/X/Z/Y/B/I/U, F2, F4, Ctrl+flèches, Ctrl+Maj+flèches, Ctrl+E, Ctrl+S/O/P…) sont vérifiés par des tests.
- **Performances** : ouverture et enregistrement des .xlsx dans un Web Worker (100 000 lignes sans figer l'interface) ; grille virtualisée et calcul des formules en arrière-plan (Univer).

### Corrigé
- Les commandes d'interface (mode sombre, audit…) ne marquent plus le classeur comme modifié.

## [0.4.0] — Phase 4 : formules et analyse — 2026-10-15

### Ajouté
- **Graphiques** (ECharts, Apache-2.0) : histogramme, barres, courbes, aires (groupés ou empilés), secteurs, nuage de points, **graphique combiné** et **axe secondaire** ; insertion depuis la sélection (Insérer → Graphique), déplacement et redimensionnement à la souris, mise à jour en direct, boîte « Modifier le graphique » (type, titre, légende, titres d'axes, couleur et type de chaque série, axe secondaire, suppression).
- Graphiques **enregistrés dans le .xlsx** (DrawingML : s'ouvrent dans Excel et LibreOffice) et **relus** depuis les fichiers Excel.
- **Tableau croisé dynamique** simple : lignes, colonnes, valeurs (somme, nombre, moyenne, min, max), filtres, totaux et total général, dans une nouvelle feuille ; Actualiser et Modifier (Données → Tableau croisé dynamique).
- **Valeur cible** (Données).
- Formules : **Gestionnaire de noms**, **Repérer les antécédents / les dépendants** (liste cliquable).
- Fonctions **ISO.CEILING** et **TRIMRANGE** ; documentation de la couverture des fonctions (docs/FONCTIONS.md).

### Corrigé
- Serveur de développement sous Windows : le worker de calcul n'était pas trouvé dans certains dossiers redirigés.

## [0.3.0] — Phase 3 : onglet Données — 2026-10-14

### Ajouté
- Données : **Supprimer les doublons** (choix des colonnes, en-têtes, insensible à la casse), **Convertir** (texte en colonnes avec séparateurs, guillemets, nombres et dates reconnus), **Remplissage instantané** (bouton et Ctrl+E), **Filtre avancé** (zone de critères ET/OU, jokers, sur place ou copie, sans doublon) et « Afficher tout ».
- Données : **Plan** (grouper/dissocier lignes et colonnes, masquer/afficher le détail) et **Sous-total** (à chaque changement de…, fonctions Somme/Nombre/Moyenne/Max/Min/Produit, total général, niveaux de plan).
- **Tableaux structurés** (plugin Univer open source) : « Mettre sous forme de tableau », **ligne de totaux** (SUBTOTAL), références structurées dans les formules (`Ventes[Prix]`).
- Données : « À partir d'un CSV ».
- Import/export .xlsx : **validations de données** (liste, nombre entier/décimal, date, longueur de texte, formule, messages d'aide et d'erreur, listes pointant vers une plage, format x14 d'Excel), **filtres automatiques avec leurs critères** (lignes filtrées masquées comme dans Excel), **tableaux** (nom, style, ligne de totaux), **niveaux de plan**, option « Interrompre si vrai » des MFC.
- Une couche de lecture/écriture XML (JSZip + fast-xml-parser, MIT) complète ExcelJS ; les validations sur des colonnes entières ne sont plus dépliées cellule par cellule.

### Corrigé
- Format Standard : au plus 11 caractères, comme Excel (10/3 → 3,333333333).

## [0.2.0] — Phase 2 : onglet Démarrer complet — 2026-10-14

### Ajouté
- Ruban en grille (deux rangées, groupes), plus proche du ruban d'Excel.
- Démarrer : boutons Couper / Copier, « Coller les formules » dans le menu Coller, Augmenter / Diminuer le retrait.
- Démarrer : menus Insérer, Supprimer (cellules, lignes, colonnes, feuille), Format (ajustement automatique de la hauteur et de la largeur, masquer, renommer la feuille), Effacer (tout, formats, contenu).
- Démarrer : raccourcis Mise en forme conditionnelle, Trier et filtrer, Rechercher.
- Import/export .xlsx des **mises en forme conditionnelles** : règles de valeurs, texte, dates, doublons/valeurs uniques, top/bottom, moyenne, formule, barres de données, nuances de couleurs, jeux d'icônes.
- Import/export du **texte enrichi** (plusieurs styles dans une cellule, sauts de ligne).
- Annuler/rétablir sans limite (Univer en gardait 50).

### Corrigé
- Téléchargements plus fiables : le lien temporaire n'est plus retiré immédiatement.

## [0.1.0] — Phase 1 : fondations et import/export — 2026-10-14

### Ajouté
- Application Vite + TypeScript avec Univer (presets open source uniquement), interface en français.
- Menu Fichier : nouveau, ouvrir, enregistrer, enregistrer sous, importer/exporter CSV ; raccourcis Ctrl/⌘+O, Ctrl/⌘+S, Ctrl/⌘+Maj+S ; glisser-déposer.
- Convertisseur .xlsx ⇄ modèle Univer (ExcelJS) : valeurs, formules, types, dates, styles, couleurs de thème, bordures, alignements, formats de nombre, fusions, tailles, volets figés, feuilles masquées, noms définis, liens.
- Import/export CSV au format Excel français (« ; », virgule décimale, BOM UTF-8), détection automatique à l'import.
- Affichage à la française des nombres au format Standard et des booléens.
- Sauvegarde automatique dans IndexedDB et restauration au démarrage.
- PWA : manifeste, icônes, service worker, fonctionnement hors ligne, ouverture de fichiers depuis le système.
- Tests : Vitest (convertisseurs, round-trip, CSV, ouverture LibreOffice), Playwright (UI, round-trip via l'interface, hors-ligne, WebKit iPad).
- CI GitHub Actions (tests + déploiement GitHub Pages).
