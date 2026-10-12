# Rush 2 – Analyse des ventes d'une pharmacie

Analyse de près de 6 ans de ventes (janvier 2014 → octobre 2019) de 8 groupes de médicaments pour une pharmacie indépendante : qualité des données, statistiques, prévision de la demande, outil de sélection Excel et recommandations.

## Équipe et répartition des tâches

| Membre | Responsabilités |
|---|---|
| Maïlys | Contrôle de cohérence des 4 exports, statistiques et test de prévision en Python, notebooks `analyse.ipynb` et `classeur.ipynb`, interface de mise à jour, analyse par catégorie, mémo partie 1 (prévision), deck manager, gestion du dépôt GitHub |
| Sofia | Classeur Excel (synthèse, statistiques, état des données), outil de sélection pour la pharmacienne, README, deck pharmacienne des achats |
| Côme | Données publiques (grippe, Réseau Sentinelles) et `donnees_externes.py`, mémo partie 2 (partage des données / RGPD), deck propriétaire |
| Toutes et tous | Relecture croisée, répétition des oraux |

## Contenu du dépôt

| Fichier / dossier | Rôle |
|---|---|
| `data/` | Les 4 exports du client (horaire, journalier, hebdo, mensuel) |
| `comparaison.py` | Vérifie que les 4 exports concordent entre eux |
| `statistiques.py` | Calcule les statistiques descriptives → `statistiques.xlsx` |
| `prevision.py` | Teste les méthodes de prévision et prévoit le mois suivant → `prevision.xlsx` |
| `donnees_externes.py` | Croise les ventes avec la grippe (Réseau Sentinelles) |
| `analyse.ipynb` | Notebook Jupyter : toute l'analyse expliquée étape par étape |
| `classeur.ipynb` | Met à jour le classeur Excel à partir du dernier export (Run All) |
| `mettre_a_jour.bat` | Double-clic : ouvre l'interface de mise à jour |
| `interface.py` | L'interface : importer les CSV, mettre à jour le classeur, voir les résultats (tableau, graphique) et les catégories |
| `categories.py` | Range les médicaments par catégorie grâce à leur code ATC, et repère les nouveaux médicaments |
| `excel/` | Classeur final avec l'outil de sélection |
| `donnees_externes/` | Données publiques utilisées dans l'analyse |
| `memo/` | Mémo PDF (prévision + partage des données) |
| `decks/` | Les 3 présentations (manager, pharmacienne, propriétaire) |

## Classeur Excel et outil de sélection

Le classeur se trouve dans le dossier `excel/`. Il s'ouvre dans Excel 365 et ne contient aucune macro.

### Ce qu'il contient

| Onglet | Ce qu'on y trouve |
|---|---|
| Synthèse | Page d'accueil : les conclusions principales et les chiffres clés |
| Ventes (outil de sélection) | Les ventes d'un ou plusieurs groupes sur un jour ou une période |
| Calendrier | Les mois forts et les mois creux de chaque groupe |
| Conseils | Pour chaque groupe : évolution, saison et conseil de commande |
| Prévision | Peut-on prévoir les ventes du mois suivant : le test et son résultat |
| Questions | Les réponses aux questions les plus fréquentes |
| Médicaments | Ce que contient chaque groupe, selon la classification de l'OMS |
| Méthode | Comment chaque chiffre est obtenu |
| Statistiques | Les chiffres détaillés : par groupe, mois, jour, heure, année |
| Catégories | Les médicaments regroupés par catégorie (douleur, anxiété et sommeil, respiratoire, allergies) et les nouveaux médicaments détectés |
| État des données | L'export retenu, ce qui a été vérifié, corrigé ou écarté |
| Données | Les ventes jour par jour, lues par l'onglet Ventes |

### Utiliser l'outil de sélection

L'outil s'adresse à la pharmacienne responsable des achats. Il ne demande aucune connaissance de Python.

1. Ouvrir le classeur dans Excel 365.
2. Aller sur l'onglet Ventes.
3. Dans les cases jaunes, saisir une date de début et une date de fin (la même date deux fois pour un seul jour).
4. Mettre OUI ou NON devant chaque groupe de médicaments.
5. Les ventes de la période, la moyenne par jour et la comparaison avec l'année précédente s'affichent aussitôt.

### Nom courant des groupes

Les exports utilisent les codes de la classification ATC. Le classeur et le deck pharmacienne emploient les noms courants.

| Code | Nom courant |
|---|---|
| M01AB | Anti-inflammatoires type diclofénac |
| M01AE | Anti-inflammatoires type ibuprofène |
| N02BA | Aspirine et dérivés |
| N02BE | Paracétamol |
| N05B | Anxiolytiques |
| N05C | Somnifères et sédatifs |
| R03 | Asthme et BPCO |
| R06 | Antihistaminiques (allergies) |

### D'où viennent les chiffres

Le notebook `classeur.ipynb` met à jour le classeur à partir des exports journalier et horaire : onglet Données, formules de l'onglet Ventes, onglet Statistiques, onglet Prévision et chiffres de prévision de la Synthèse. Les textes écrits par l'équipe (Calendrier, Conseils, État des données) sont à relire à la main.

L'unité des quantités n'est pas précisée dans les exports. Le classeur présente donc des volumes et des écarts, pas un nombre de boîtes.

## Relancer l'analyse avec un nouvel export

Le plus simple : double-cliquer sur `mettre_a_jour.bat`, cliquer sur « Choisir les fichiers CSV… » pour importer le nouvel export, puis sur « Mettre à jour le classeur ».

Pour le faire à la main :

1. Installer Python 3 (python.org, en cochant « Add python.exe to PATH »).
2. Remplacer les 4 fichiers du dossier `data/` par le nouvel export (mêmes noms de fichiers).
3. Ouvrir un terminal dans ce dossier et installer les librairies :

       python -m pip install -r requirements.txt

4. Lancer les scripts :

       python comparaison.py
       python statistiques.py
       python prevision.py
       python donnees_externes.py

   Les fichiers `statistiques.xlsx` et `prevision.xlsx` sont recréés avec les données à jour.

5. Ouvrir `classeur.ipynb` et cliquer sur Run All : le classeur Excel est mis à jour.

## Choix sur les données

- **Fichier de référence :** l'export journalier. Il concorde à 100 % avec l'horaire et l'hebdo.
- **L'export mensuel est écarté :** 31 mois sur 70 ne correspondent pas à la somme des jours.
- **Les mois incomplets** (janvier 2014 et octobre 2019) sont retirés des analyses mensuelles.
- **Les dates de l'export journalier sont au format américain** (mois en premier) : `1/2/2014` correspond au 2 janvier 2014.
- **La colonne `Hour` de l'export journalier** n'a pas de sens à l'échelle d'une journée et n'est pas utilisée.
