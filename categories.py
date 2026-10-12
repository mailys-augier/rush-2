"""Regroupe les médicaments par catégorie à partir de leur code ATC (le « nom de référence » de l'export).

Le code ATC se lit de gauche à droite, du plus large au plus précis :
    N02BE  =  N (système nerveux) > N02 (analgésiques) > N02B (autres analgésiques et antipyrétiques) > N02BE (paracétamol)

Si un nouvel export contient un nouveau médicament (une nouvelle colonne, par exemple N02BB ou R05),
il est rangé automatiquement dans la bonne catégorie grâce au début de son code, et signalé comme nouveau.

Le manager peut aussi ajouter ou reclasser des médicaments à la main, depuis l'interface (onglet Catégories).
Ses choix sont enregistrés dans categories_perso.json et passent avant le classement automatique,
partout : dans l'interface et dans le classeur Excel.
"""

import json
import re
from pathlib import Path

import pandas as pd

FICHIER_PERSO = Path(__file__).resolve().parent / "categories_perso.json"   # les ajouts faits à la main

# Les 8 groupes de l'export d'origine (2014-2019)
GROUPES_CONNUS = ["M01AB", "M01AE", "N02BA", "N02BE", "N05B", "N05C", "R03", "R06"]

# Catégories « pharmacie » : début du code ATC -> catégorie. Le début le plus long qui correspond gagne.
CATEGORIES = {
    "M01": "Douleur et inflammation",
    "N02": "Douleur et inflammation",
    "N05": "Anxiété et sommeil",
    "N06": "Anxiété et sommeil",
    "R03": "Respiratoire",
    "R05": "Respiratoire",
    "R06": "Allergies",
}

# Nom officiel des sous-groupes ATC utiles ici (pour l'affichage)
SOUS_GROUPES = {
    "M01A": "Anti-inflammatoires non stéroïdiens",
    "M01AB": "Dérivés de l'acide acétique (diclofénac)",
    "M01AE": "Dérivés de l'acide propionique (ibuprofène)",
    "N02A": "Opioïdes",
    "N02B": "Autres analgésiques et antipyrétiques",
    "N02BA": "Acide salicylique et dérivés (aspirine)",
    "N02BB": "Pyrazolones",
    "N02BE": "Anilides (paracétamol)",
    "N02C": "Antimigraineux",
    "N05A": "Antipsychotiques",
    "N05B": "Anxiolytiques",
    "N05C": "Hypnotiques et sédatifs",
    "N06A": "Antidépresseurs",
    "R03": "Médicaments de l'asthme et de la BPCO",
    "R05": "Médicaments de la toux et du rhume",
    "R06": "Antihistaminiques",
}

# Si le début du code n'est dans aucune catégorie ci-dessus : la grande famille (1re lettre du code)
FAMILLES = {
    "A": "Digestion et métabolisme", "B": "Sang", "C": "Cœur et vaisseaux", "D": "Peau",
    "G": "Gynécologie et urologie", "H": "Hormones", "J": "Anti-infectieux", "L": "Cancer et immunité",
    "M": "Muscles et articulations", "N": "Système nerveux", "P": "Antiparasitaires",
    "R": "Respiratoire", "S": "Yeux et oreilles", "V": "Divers",
}

MOTIF_ATC = re.compile(r"^[A-Z]\d{2}([A-Z]{1,2}(\d{2})?)?$")   # N02, N02B, N02BE, N02BE01…


def codes_atc(colonnes):
    """Les colonnes qui sont des codes ATC (on ignore datum, Year, Month, Hour, Weekday Name…)."""
    return [c for c in colonnes if MOTIF_ATC.match(str(c))]


# ----------------------------------------------------------------- les ajouts faits à la main

def charger_perso():
    """Les médicaments ajoutés ou reclassés à la main : {code: {"nom": ..., "categorie": ...}}."""
    try:
        return json.loads(FICHIER_PERSO.read_text(encoding="utf-8"))
    except (FileNotFoundError, ValueError):
        return {}


def enregistrer_perso(perso):
    """Enregistre les ajouts faits à la main dans categories_perso.json."""
    FICHIER_PERSO.write_text(json.dumps(perso, ensure_ascii=False, indent=2, sort_keys=True), encoding="utf-8")


def code_valide(code):
    """Vrai si le texte a la forme d'un code ATC (N02BE, R05, N02BE01…)."""
    return bool(MOTIF_ATC.match(code))


# ----------------------------------------------------------------- le classement

def categorie_auto(code):
    """La catégorie d'un code ATC selon la classification, sans tenir compte des ajouts à la main."""
    for longueur in range(len(code), 2, -1):
        if code[:longueur] in CATEGORIES:
            return CATEGORIES[code[:longueur]]
    return f"Autres : {FAMILLES.get(code[0], 'famille inconnue')}"


def categorie(code):
    """La catégorie d'un code ATC, même inconnu. Un choix fait à la main passe en premier."""
    perso = charger_perso().get(code, {})
    return perso.get("categorie") or categorie_auto(code)


def sous_groupe(code):
    """Le nom le plus précis connu pour ce code (le nom saisi à la main passe en premier)."""
    perso = charger_perso().get(code, {})
    if perso.get("nom"):
        return perso["nom"]
    for longueur in range(len(code), 2, -1):
        if code[:longueur] in SOUS_GROUPES:
            return SOUS_GROUPES[code[:longueur]]
    return "Sous-groupe à identifier"


def toutes_les_categories():
    """Les catégories connues (classification + ajouts à la main), pour les listes de choix."""
    noms = set(CATEGORIES.values()) | {p["categorie"] for p in charger_perso().values() if p.get("categorie")}
    return sorted(noms)


def nouveaux_codes(colonnes):
    """Les médicaments présents dans l'export mais absents de l'export d'origine."""
    return [c for c in codes_atc(colonnes) if c not in GROUPES_CONNUS]


def tableau_codes(colonnes):
    """Un tableau : code, sous-groupe ATC, catégorie, nouveau ?"""
    return pd.DataFrame([{"Code": c, "Sous-groupe ATC": sous_groupe(c), "Catégorie": categorie(c),
                          "Nouveau ?": "Oui" if c not in GROUPES_CONNUS else "Non"} for c in codes_atc(colonnes)])


def ventes_par_categorie(jour):
    """Ventes par mois complet et par catégorie. `jour` = l'export journalier avec une colonne datum (dates)."""
    codes = codes_atc(jour.columns)
    par_jour = jour.set_index("datum")[codes]
    mois = par_jour.resample("MS").sum()
    nb_jours = par_jour.resample("MS").size()
    mois = mois[nb_jours == nb_jours.index.days_in_month]          # mois complets seulement
    return mois.T.groupby(categorie).sum().T                        # additionne les codes de chaque catégorie
