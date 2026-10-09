# Rush 2 – Analyse des ventes d'une pharmacie

Analyse de près de 6 ans de ventes (janvier 2014 → octobre 2019) de 8 groupes de médicaments
pour une pharmacie indépendante : qualité des données, statistiques, prévision de la demande,
outil de sélection Excel et recommandations.

## Équipe et répartition des tâches

| Membre | Responsabilités |
|---|---|
| Maïlys | Contrôle de cohérence des 4 exports, statistiques descriptives en Python, test de prévision, mémo partie 1 (prévision), deck manager |
| Sophia | Classeur Excel (synthèse, statistiques, état des données), outil de sélection pour la pharmacienne, deck pharmacienne des achats |
| Côme | Données publiques externes, mémo partie 2 (partage des données / RGPD), deck propriétaire |
| Toutes et tous | Relecture croisée, README, répétition des oraux |

## Contenu du dépôt

| Fichier / dossier | Rôle |
|---|---|
| `data/` | Les 4 exports du client (horaire, journalier, hebdo, mensuel) |
| `comparaison.py` | Vérifie que les 4 exports concordent entre eux |
| `statistiques.py` | Calcule les statistiques descriptives → `statistiques.xlsx` |
| `prevision.py` | Teste les méthodes de prévision et prévoit le mois suivant → `prevision.xlsx` |
| `excel/` | Classeur final avec l'outil de sélection |
| `donnees_externes/` | Données publiques utilisées dans l'analyse |
| `memo/` | Mémo PDF (prévision + partage des données) |
| `decks/` | Les 3 présentations (manager, pharmacienne, propriétaire) |

## Relancer l'analyse avec un nouvel export

1. Installer Python 3 (python.org, en cochant « Add python.exe to PATH »).
2. Remplacer les 4 fichiers du dossier `data/` par le nouvel export (mêmes noms de fichiers).
3. Ouvrir un terminal dans ce dossier et installer les librairies :

       pip install -r requirements.txt

4. Lancer les scripts :

       python comparaison.py
       python statistiques.py
       python prevision.py

5. Les fichiers `statistiques.xlsx` et `prevision.xlsx` sont recréés avec les données à jour.

## Choix sur les données

- Fichier de référence : l'export **journalier**. Il concorde à 100 % avec l'horaire et l'hebdo.
- L'export **mensuel** est écarté : 31 mois sur 70 ne correspondent pas à la somme des jours.
- Les mois incomplets (dernier mois de l'export) sont retirés des analyses mensuelles.