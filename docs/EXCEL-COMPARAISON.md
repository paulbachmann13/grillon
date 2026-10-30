# Grillon face à Excel

Inventaire des commandes du ruban d'Excel (Microsoft 365, version Windows/Mac en français),
groupe par groupe, avec leur état dans Grillon. C'est la liste de référence pour les prochaines
étapes : on travaille dans l'ordre de ce que les utilisateurs d'Excel emploient le plus.

Légende : **Oui** = disponible · **Partiel** = disponible avec des limites (voir la note) · **Non** = absent.

## Fichier

| Commande Excel | Grillon | Note |
|---|---|---|
| Nouveau, Ouvrir, Enregistrer, Enregistrer sous | Oui | .xlsx et .csv ; ⌘S réécrit le fichier (Chrome, Edge, app Mac) |
| Imprimer, aperçu, Exporter en PDF | Oui | Réglages et aperçu des pages ; PDF par « Enregistrer au format PDF » |
| Mise en page (boîte) | Oui | Papier, marges, échelle, zone, titres, en-têtes/pieds |
| Ouvrir .xls, .ods, .xlsb | Non | Réenregistrer en .xlsx |
| Mot de passe d'ouverture | Non | |
| Partager / coauteur en temps réel | Non | Hors périmètre (100 % local) |
| Historique des versions | Partiel | Annuler sans limite ; sauvegarde automatique locale |

## Accueil

| Commande Excel | Grillon | Note |
|---|---|---|
| Presse-papiers : couper, copier, coller, collage spécial, reproduire la mise en forme | Oui | |
| Police : police, taille, gras/italique/souligné/barré, bordures, remplissage, couleur | Oui | |
| Alignement : horizontal, vertical, orientation, renvoi à la ligne, fusionner, retrait | Oui | |
| Nombre : formats, %, €, séparateur de milliers, décimales | Oui | |
| Mise en forme conditionnelle (règles, barres, nuances, icônes, gestion) | Oui | |
| Mettre sous forme de tableau (galerie de styles) | Oui | 21 styles (Clair, Moyen, Foncé × 7 couleurs) ; les autres styles d'Excel sont rapprochés |
| Styles de cellule (Bon, Neutre, Titre 1…, Entrée, Calcul…) | Oui | Galerie dans Accueil |
| Format de cellule (boîte Ctrl+1 / ⌘1) | Oui | Nombre, Alignement, Police, Bordure, Remplissage, Protection (Verrouillée) |
| Cellules : insérer, supprimer, format (hauteur, largeur, masquer, renommer) | Oui | |
| Édition : somme automatique, remplir, effacer, trier et filtrer, rechercher/remplacer | Oui | « Série » (remplir) partiel |
| Analyse rapide | Oui | |

## Insertion

| Commande Excel | Grillon | Note |
|---|---|---|
| Tableau croisé dynamique | Oui | Vrai tableau croisé dans le .xlsx ; voir « Données » |
| Tableau | Oui | |
| Images, formes, zones de texte | Oui | Formes simples (6) ; pas d'icônes, SmartArt, modèles 3D |
| Graphiques (histogramme, courbes, secteurs, barres, aires, nuage, bulles, radar, combinés) | Oui | |
| Graphiques cascade, entonnoir, histogramme statistique, boîte à moustaches | Oui | Format chartEx d'Excel 2016, vérifié dans Excel pour Mac |
| Graphiques boursier, surface, compartimentage, rayons de soleil, Pareto, cartes | Non | |
| Graphique croisé dynamique | Non | |
| Mini-graphiques | Oui | |
| Segments, chronologies | Partiel | Sur les tableaux (plage avec en-têtes) et les tableaux croisés de Grillon ; les segments sont de vrais segments d'Excel dans le .xlsx (et ceux d'Excel sont relus) ; les chronologies ne sont gardées que pour Grillon (pas de chronologie Excel) ; pas de style de segment ni de colonnes multiples |
| Lien | Oui | |
| Commentaire (fil de discussion), note | Partiel | Notes uniquement |
| En-tête et pied de page | Oui | Via Mise en page |
| WordArt, signature, objet, équation | Non | |
| Symbole | Oui | |

## Mise en page

| Commande Excel | Grillon | Note |
|---|---|---|
| Thèmes (couleurs, polices, effets) | Partiel | Thèmes, Couleurs et Polices dans Vue, lus et écrits dans le .xlsx ; pas d'effets ; la palette d'Univer (sélecteur de couleur) n'affiche pas les couleurs du thème |
| Marges, orientation, taille, zone d'impression, imprimer les titres | Oui | Boîte Mise en page |
| Sauts de page manuels, aperçu des sauts de page | Oui | Insérer, supprimer, rétablir ; aperçu en lecture seule (pas de glisser des lignes bleues) ; ordre « de haut en bas puis à droite » uniquement |
| Arrière-plan de feuille | Oui | (Univer) |
| Mise à l'échelle (largeur, hauteur, %) | Oui | |
| Quadrillage et en-têtes : afficher / imprimer | Oui | |
| Organiser (premier plan, aligner, grouper) | Non | |

## Formules

| Commande Excel | Grillon | Note |
|---|---|---|
| Insérer une fonction, bibliothèque par catégorie | Oui | |
| Somme automatique | Oui | |
| Gestionnaire de noms, définir un nom | Oui | |
| Repérer les antécédents / dépendants | Partiel | Liste cliquable, pas de flèches sur la grille |
| Afficher les formules | Oui | |
| Vérification des erreurs, évaluer la formule | Oui | Erreurs et nombres stockés en texte ; pas de triangles verts sur la grille |
| Fenêtre Espion | Oui | Boîte (non flottante) ; liste gardée en mémoire, non enregistrée dans le .xlsx |
| Options de calcul (automatique / manuel), calculer maintenant | Oui | F9 ; mode non enregistré dans le fichier |
| ≈ 520 fonctions, tableaux dynamiques, LAMBDA | Oui | 23 fonctions absentes (docs/FONCTIONS.md) |

## Données

| Commande Excel | Grillon | Note |
|---|---|---|
| Récupérer des données : texte/CSV | Oui | Assistant d'import |
| Récupérer des données : web, base de données, Power Query | Non | Hors périmètre (local) |
| Actualiser tout | Partiel | Tableaux croisés : un par un |
| Trier, filtrer, filtre avancé, effacer, réappliquer | Oui | Comme Excel : une cellule ou une colonne étend le tri au tableau, l'en-tête est reconnu, ordre alphabétique français, les formules suivent leur ligne |
| Convertir, remplissage instantané, supprimer les doublons, validation | Oui | |
| Consolider | Non | |
| Relations, modèle de données | Non | |
| Analyse de scénarios : valeur cible | Oui | |
| Analyse de scénarios : gestionnaire de scénarios, table de données | Non | |
| Feuille de prévision | Non | |
| Plan : grouper, dissocier, sous-total | Oui | Boutons 1 2 3 et − / + dans la marge |
| Tableau croisé : champs, filtres, totaux, actualiser | Oui | Calculé par Grillon |
| Tableau croisé : vrai TCD dans le .xlsx (modifiable dans Excel) | Oui | Cache, enregistrements et définition ; actualisé à l'ouverture ; reconnu par LibreOffice |
| Tableau croisé : regroupement de dates, « afficher les valeurs en » | Non | |

## Révision

| Commande Excel | Grillon | Note |
|---|---|---|
| Orthographe, dictionnaire des synonymes | Non | |
| Vérifier l'accessibilité | Non | |
| Traduire | Non | |
| Commentaires (fil), notes | Partiel | Notes |
| Protéger la feuille / le classeur, autoriser la modification des plages | Partiel | Feuille entière, mot de passe facultatif (hachage Excel), enregistrée dans le .xlsx ; cellules déverrouillées ; pas de protection du classeur ni de plages autorisées |
| Suivi des modifications | Non | |

## Affichage

| Commande Excel | Grillon | Note |
|---|---|---|
| Normal, mise en page, aperçu des sauts de page | Partiel | Normal uniquement |
| Affichages personnalisés | Non | |
| Règle, quadrillage, barre de formule, en-têtes | Partiel | Quadrillage |
| Zoom, 100 %, zoom sur la sélection | Oui | |
| Nouvelle fenêtre (un classeur par fenêtre), réorganiser, fractionner | Partiel | App Mac : une fenêtre par classeur ; pas de réorganisation ni de fractionnement |
| Figer les volets | Oui | |
| Masquer / afficher une feuille | Oui | |
| Macros (VBA), Office Scripts | Non | Hors périmètre |
| Mode sombre | Oui | |

## Hors ruban

| Fonction Excel | Grillon | Note |
|---|---|---|
| Barre d'état (somme, moyenne, nombre) | Oui | |
| Poignée de recopie, séries | Oui | |
| Menu contextuel (clic droit) | Oui | |
| Raccourcis clavier courants | Oui | Voir docs/FONCTIONNALITES.md |
| Touches d'accès (Alt…) | Non | |
| Grille jusqu'à 1 048 576 × 16 384 | Oui | |
| Classeurs de 100 000 lignes | Oui | Quelques secondes à l'ouverture |
| Application Mac et Windows | Partiel | Mac : .dmg construit, impression et mise à jour vérifiées ; Windows : web/PWA |
| iPad / iPhone | Oui | Appli installée depuis Safari |

## Comparaison automatique avec le vrai Excel

Relevé du 30/10/2026 avec Excel pour Mac (version sans abonnement, qui ouvre et calcule mais n'enregistre pas). Le même fichier, écrit par
Grillon, est calculé par Excel et par Grillon ; les résultats d'Excel sont dans `tests/fixtures/excel-reference.json` et vérifiés à
chaque test (`tests/e2e/excel-reference.spec.ts`).

| Domaine | Résultat |
|---|---|
| 635 formules (maths, stats, logique, texte, dates, recherche, finance, ingénierie, base de données, usages courants) | 566 identiques ; 43 non vérifiables (fonction absente de cette version d'Excel, voir ci-dessous) ; 26 écarts acceptés (dont 16 nombres complexes) |
| 274 valeurs × formats de nombre | Identiques, aux écarts acceptés près |
| Classeur relu dans Excel : polices, couleurs, bordures, alignements, rotation, retrait, largeurs, hauteurs, masquage, plan, fusions, volets, zoom, quadrillage, impression, noms, filtre, lien, validation, mise en forme conditionnelle, feuille masquée, protection | Conformes |
| 353 fichiers de test d'Apache POI : ouverture, enregistrement, réouverture | 326 sans perte de valeur ni de formule ; 26 fichiers abîmés refusés ; 1 fichier d'attaque (bombe de chaînes) |

**Non vérifiables ici** : `SI.CONDITIONS`, `SI.MULTIPLE`, `CONCAT`, `JOINDRE.TEXTE`, `MAX.SI.ENS`, `MIN.SI.ENS`, `RECHERCHEX`, `EQUIVX`, `LET`, `TRIER`, `TRIERPAR`, `FILTRE`, `UNIQUE`, `SEQUENCE`, `ASSEMB.H`, `ASSEMB.V`… Cette version d'Excel ne les a pas : le fichier s'ouvre sans réparation et elle affiche `#NOM?`. Les résultats de Grillon pour ces fonctions n'ont donc pas été comparés à Excel, et le débordement des formules marquées matricielles n'a pas été vu dans un Excel 365.

**Écarts acceptés** :

| Cas | Excel | Grillon |
|---|---|---|
| `NB.SI(A:A;">=12.5")` avec un point | Excel français : aucun résultat | Comprend le point (fichiers venus d'un Excel anglais) |
| Mois et jours abrégés (`mmm`, `jjj`) | Mac : `janv`, `lun` | `janv.`, `lun.` (forme d'Excel Windows, à vérifier) |
| Format `h:mm:ss` à minuit | `00:00:00` (format n° 21) | `0:00:00` |
| Nombres complexes (`IMDIV`…) | `0,6+0,2i` | `0.6+0.2i` |
| `EQUIV(x;liste croissante;-1)` | `#N/A` | Une position |
| `ADRESSE(…;FAUX)` | `L2C3` | `R2C3` |
| `ZONES((A1;C1))` (union) | `2` | `#VALEUR!` |
| `DBCS` | Inchangé en français | Pleine chasse |
