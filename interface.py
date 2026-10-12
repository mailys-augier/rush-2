"""Interface de suivi des ventes de la pharmacie, pour le cabinet (le manager).

Lancement : double-clic sur mettre_a_jour.bat (ou : python interface.py).

Quatre pages, dans le menu de gauche :
- « Mise à jour » : importer les nouveaux exports CSV, puis remettre à jour le classeur Excel (en un bouton) ;
- « Ventes » : une vue de synthèse, avec des filtres (période, mois ou dates, médicaments, et des filtres
  que l'on ajoute soi-même : jours de la semaine, plage horaire, catégorie) et un tableau en couleurs ;
- « Prévision » : la prévision du mois suivant, sa fourchette, sa fiabilité, et une alerte « mois inhabituel » ;
- « Catégories » : les médicaments par catégorie ; on peut en ajouter ou en reclasser à la main.

L'interface ne remplace pas le classeur Excel : elle donne une vue rapide. Le détail reste dans Excel.
"""

import io
import json
import os
import re
import shutil
import subprocess
import sys
import threading
from datetime import datetime
from pathlib import Path
from urllib.request import urlopen

import openpyxl
import pandas as pd

import categories as cat

DOSSIER = Path(__file__).resolve().parent          # le dossier rush-2
DATA = DOSSIER / "data"
CLASSEUR = DOSSIER / "excel" / "Pharma_Analyse.xlsx"
NOTEBOOK = DOSSIER / "classeur.ipynb"
FICHIER_GRIPPE = DOSSIER / "donnees_externes" / "inc-3-PAY-ds2.csv"     # Réseau Sentinelles (fichier de Côme)
URL_GRIPPE = "https://www.sentiweb.fr/datasets/all/inc-3-PAY-ds2.csv"
SEUIL_EPIDEMIE = 150                                   # cas pour 100 000 habitants par semaine (même seuil que l'analyse)
FICHIER_EVENEMENTS = DOSSIER / "evenements.json"       # les événements notés par l'équipe
FICHIER_SUIVI = DOSSIER / "suivi_previsions.csv"       # les prévisions enregistrées, mois après mois
DOSSIER_RAPPORTS = DOSSIER / "rapports"                # les synthèses PDF
TYPES_EVENEMENTS = ["Rupture de stock", "Fermeture ou travaux", "Épidémie", "Crise sanitaire (type Covid)",
                    "Promotion ou action commerciale", "Autre"]

GROUPES = cat.GROUPES_CONNUS
NOMS = {"M01AB": "Diclofénac", "M01AE": "Ibuprofène", "N02BA": "Aspirine", "N02BE": "Paracétamol",
        "N05B": "Anxiolytiques", "N05C": "Somnifères", "R03": "Asthme et BPCO", "R06": "Antihistaminiques"}
NOMS_FICHIERS = {
    "Journalier": "Pharma_Ventes_Daily.csv",
    "Horaire": "Pharma_Ventes_Hourly.csv",
    "Hebdomadaire": "Pharma_Ventes_Weekly.csv",
    "Mensuel": "Pharma_Ventes_Monthly.csv",
}
MOTS_CLES = [
    ("Horaire", ["hour", "heure", "horaire"]),
    ("Hebdomadaire", ["week", "hebdo", "semaine"]),
    ("Mensuel", ["month", "mensuel", "mois"]),
    ("Journalier", ["daily", "day", "jour"]),
]
MOIS_FR = ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août",
           "septembre", "octobre", "novembre", "décembre"]
MOIS_COURT = ["janv.", "févr.", "mars", "avr.", "mai", "juin", "juil.", "août", "sept.", "oct.", "nov.", "déc."]
JOURS = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]


def nom_court(code):
    """Le nom à afficher pour un code ATC : nom courant, sinon nom saisi à la main, sinon le code."""
    if code in NOMS:
        return NOMS[code]
    if cat.charger_perso().get(code, {}).get("nom"):
        return cat.charger_perso()[code]["nom"]
    return code                                             # nouveau médicament : son code (à nommer dans Catégories)


def nombre(x, decimales=0):
    """1307.4 -> '1 307' ; 3.75 -> '3,8' (avec decimales=1)."""
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return "–"
    texte = f"{x:,.{decimales}f}".replace(",", " ").replace(".", ",")
    return texte


# ======================================================================= import et mise à jour (sans fenêtre)

def lire_dates(colonne):
    """Lit la colonne des dates, quel que soit le format utilisé par le logiciel de caisse."""
    for fmt in ("%m/%d/%Y %H:%M", "%m/%d/%Y", "%Y-%m-%d", "%Y-%m-%d %H:%M:%S"):
        try:
            return pd.to_datetime(colonne, format=fmt)
        except (ValueError, TypeError):
            continue
    raise ValueError("format de date non reconnu")


def reconnaitre(chemin):
    """Renvoie (type d'export, nombre de lignes, première date, dernière date, nouveaux codes)."""
    table = pd.read_csv(chemin)
    manquants = [c for c in ["datum"] + GROUPES if c not in table.columns]
    if manquants:
        raise ValueError(f"ce n'est pas un export de ventes complet (colonnes absentes : {', '.join(manquants)})")
    dates = lire_dates(table["datum"])
    nom = Path(chemin).name.lower()
    type_export = next((t for t, mots in MOTS_CLES if any(m in nom for m in mots)), None)
    if type_export is None:                         # rien dans le nom : on regarde l'écart entre deux dates
        if (dates.dt.hour != 0).any():
            type_export = "Horaire"
        else:
            ecart = dates.sort_values().diff().dt.days.median()
            type_export = "Journalier" if ecart <= 1 else "Hebdomadaire" if ecart <= 7 else "Mensuel"
    return type_export, len(table), dates.min(), dates.max(), cat.nouveaux_codes(table.columns)


def importer(chemins):
    """Copie chaque CSV choisi dans data/ sous le nom attendu. Renvoie les lignes du compte rendu."""
    DATA.mkdir(exist_ok=True)
    compte_rendu, vus = [], {}
    for chemin in chemins:
        nom = Path(chemin).name
        try:
            type_export, n, debut, fin, nouveaux = reconnaitre(chemin)
        except Exception as erreur:
            compte_rendu.append(f"✗ {nom} : refusé, {erreur}")
            continue
        if type_export in vus:
            compte_rendu.append(f"✗ {nom} : refusé, un export {type_export.lower()} a déjà été choisi ({vus[type_export]})")
            continue
        vus[type_export] = nom
        cible = DATA / NOMS_FICHIERS[type_export]
        if Path(chemin).resolve() != cible.resolve():
            shutil.copyfile(chemin, cible)
        compte_rendu.append(f"✓ {type_export} : {n} lignes, du {debut:%d/%m/%Y} au {fin:%d/%m/%Y}  →  data/{cible.name}")
        if nouveaux:
            compte_rendu.append(f"   ★ Nouveaux médicaments dans cet export : {', '.join(nouveaux)} (voir la page Catégories)")
    manquants = [t for t in ("Journalier", "Horaire") if t not in vus]
    if manquants and vus:
        compte_rendu.append(f"! Pas de nouvel export {' ni '.join(m.lower() for m in manquants)} : l'ancien est gardé.")
    return compte_rendu


def verifier():
    """Contrôle rapide : l'horaire additionné par jour doit donner le journalier."""
    jour = pd.read_csv(DATA / NOMS_FICHIERS["Journalier"])
    heure = pd.read_csv(DATA / NOMS_FICHIERS["Horaire"])
    jour.index = lire_dates(jour["datum"]).dt.normalize()
    heure_par_jour = heure[GROUPES].groupby(lire_dates(heure["datum"]).dt.normalize()).sum()
    communs = jour.index.intersection(heure_par_jour.index)
    ecarts = ((jour.loc[communs, GROUPES] - heure_par_jour.loc[communs, GROUPES]).abs() > 0.01).any(axis=1).sum()
    fin_jour, fin_heure = jour.index.max(), heure_par_jour.index.max()
    if fin_jour != fin_heure:
        return (f"! Contrôle : le journalier s'arrête le {fin_jour:%d/%m/%Y} et l'horaire le {fin_heure:%d/%m/%Y}. "
                "Importez les deux exports du même envoi.")
    if ecarts == 0:
        return f"✓ Contrôle : l'horaire et le journalier concordent sur {len(communs)} jours."
    return f"! Contrôle : {ecarts} jours où l'horaire et le journalier ne concordent pas. À signaler au client."


def excel_ouvert():
    """Excel crée un fichier ~$… à côté du classeur tant qu'il est ouvert."""
    return (CLASSEUR.parent / f"~${CLASSEUR.name}").exists()


def mettre_a_jour():
    """Lance classeur.ipynb sans ouvrir Jupyter. Renvoie (réussi ?, message)."""
    if not NOTEBOOK.exists():
        return False, "le fichier classeur.ipynb est introuvable dans le dossier rush-2 (vérifiez son nom exact)."
    if not (DOSSIER / "categories.py").exists():
        return False, "le fichier categories.py manque à côté de classeur.ipynb (dossier rush-2)."
    python = Path(sys.executable)
    if python.name.lower() == "pythonw.exe":                      # pythonw n'affiche pas les erreurs : on prend python.exe
        python = python.with_name("python.exe")
    commande = [str(python), "-m", "jupyter", "nbconvert", "--to", "notebook",
                "--execute", "--inplace", str(NOTEBOOK)]
    resultat = subprocess.run(commande, cwd=DOSSIER, capture_output=True, text=True, encoding="utf-8",
                              errors="replace", creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    if resultat.returncode != 0:
        sortie = re.sub(r"\x1b\[[0-9;]*m", "", (resultat.stderr or "") + "\n" + (resultat.stdout or "")).strip().splitlines()
        erreurs = [l.strip() for l in sortie if re.search(r"[A-Za-z]+(Error|Exception)\b", l) and "Equivalent" not in l]
        if erreurs:
            return False, erreurs[-1]
        return False, sortie[-1] if sortie else "erreur inconnue (aucun message)."
    return True, "✓ Classeur mis à jour."


def ouvrir(chemin):
    if sys.platform.startswith("win"):
        os.startfile(chemin)
    else:
        subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", str(chemin)])


# ======================================================================= les ventes, avec filtres (sans fenêtre)

def charger_ventes():
    """Les ventes heure par heure (ou jour par jour si l'export horaire manque).

    L'export horaire additionné par jour donne exactement le journalier (contrôle de comparaison.py) :
    on peut donc tout calculer à partir de lui, et filtrer aussi par heure.
    """
    fichier_h, fichier_j = DATA / NOMS_FICHIERS["Horaire"], DATA / NOMS_FICHIERS["Journalier"]
    horaire = fichier_h.exists()
    table = pd.read_csv(fichier_h if horaire else fichier_j)
    table["datum"] = lire_dates(table["datum"])
    codes = cat.codes_atc(table.columns)
    table = table[["datum"] + codes].sort_values("datum").reset_index(drop=True)
    table["jour"] = table["datum"].dt.normalize()
    table["semaine"] = table["datum"].dt.weekday          # 0 = lundi
    table["heure"] = table["datum"].dt.hour
    return table, codes, horaire


def filtrer(table, debut, fin, jours=None, heures=None):
    """Garde les lignes entre deux dates (incluses), et si demandé certains jours de la semaine et heures."""
    garde = (table["jour"] >= debut) & (table["jour"] <= fin)
    if jours is not None:
        garde &= table["semaine"].isin(jours)
    if heures is not None:
        garde &= table["heure"].between(heures[0], heures[1])
    return table[garde]


def synthese_ventes(table, codes, debut, fin, mode, jours=None, heures=None):
    """Le tableau de synthèse : une ligne par médicament, une colonne par période.

    mode = "Mois" ou "Année" : ventes totales de chaque période ;
    mode = "Jour de la semaine" ou "Heure" : ventes moyennes par jour.
    Renvoie (colonnes, valeurs par code, total par code, an dernier par code, nb de jours, colonnes incomplètes).
    """
    choix = filtrer(table, debut, fin, jours, heures)
    nb_jours = choix["jour"].nunique()
    if mode in ("Mois", "Année"):
        cle = choix["datum"].dt.to_period("M" if mode == "Mois" else "Y")
        sommes = choix.groupby(cle)[codes].sum()
        etiquettes = [f"{MOIS_COURT[p.month - 1]} {p.year % 100:02d}" if mode == "Mois" else str(p.year)
                      for p in sommes.index]
        # une période est incomplète si les dates choisies (ou la fin des données) la coupent
        borne_fin = min(fin, table["jour"].max())
        incompletes = [etiquettes[i] for i, p in enumerate(sommes.index)
                       if p.start_time.normalize() < debut or p.end_time.normalize() > borne_fin]
        totaux = sommes.sum()
        avant = filtrer(table, debut - pd.DateOffset(years=1), fin - pd.DateOffset(years=1), jours, heures)
        an_dernier = avant[codes].sum() if avant["jour"].nunique() >= 0.9 * nb_jours else None
    else:
        cle = choix["semaine"] if mode == "Jour de la semaine" else choix["heure"]
        sommes = choix.groupby(cle)[codes].sum()
        jours_par_cle = choix.groupby(cle)["jour"].nunique()
        sommes = sommes.div(jours_par_cle, axis=0)                     # moyenne par jour
        if mode == "Heure":
            sommes = sommes[(sommes > 0).any(axis=1)]                  # on cache les heures où la pharmacie est fermée
            etiquettes = [f"{h} h" for h in sommes.index]
        else:
            etiquettes = [JOURS[j][:3] + "." for j in sommes.index]
        incompletes = []
        totaux = choix[codes].sum() / nb_jours if nb_jours else choix[codes].sum()
        avant = filtrer(table, debut - pd.DateOffset(years=1), fin - pd.DateOffset(years=1), jours, heures)
        n_avant = avant["jour"].nunique()
        an_dernier = avant[codes].sum() / n_avant if n_avant >= 0.9 * nb_jours and n_avant else None
    valeurs = {c: list(sommes[c]) for c in codes}
    return etiquettes, valeurs, totaux, an_dernier, nb_jours, incompletes


# ======================================================================= la prévision (sans fenêtre)

def mois_complets(table):
    """Ventes par mois, en ne gardant que les mois dont tous les jours sont présents."""
    par_jour = table.groupby("jour")[cat.codes_atc(table.columns)].sum()
    mois = par_jour.resample("MS").sum()
    nb = par_jour.resample("MS").size()
    return mois[nb == nb.index.days_in_month]


# Les 4 méthodes, identiques à prevision.py et classeur.ipynb
def _mois_dernier(h, cible):
    return h.iloc[-1]


def _an_dernier(h, cible):
    return h.loc[cible - pd.DateOffset(years=1)]


def _moyenne_3(h, cible):
    return h.iloc[-3:].mean()


def _saison(h, cible):
    niveau = h.iloc[-12:].mean()
    indice = (h / h.rolling(12, center=True).mean()).dropna()
    return niveau * indice[indice.index.month == cible.month].mean()


METHODES = [("Naïve : mois dernier", _mois_dernier), ("Naïve : même mois an dernier", _an_dernier),
            ("Moyenne 3 derniers mois", _moyenne_3), ("Saison x niveau", _saison)]


def charger_prevision(table):
    """Lit la prévision dans le classeur, et contrôle si le dernier mois complet était « inhabituel »."""
    wb = openpyxl.load_workbook(CLASSEUR, read_only=True, data_only=True)
    feuille = wb["Prévision"]
    lignes = {}
    for r in range(16, 24):                                   # le test : erreur des 4 méthodes et verdict
        code = feuille.cell(r, 2).value
        if code:
            erreurs = [feuille.cell(r, c).value for c in range(3, 7)]
            lignes[code] = {"code": code, "erreurs": erreurs, "notre": feuille.cell(r, 8).value,
                            "verdict": feuille.cell(r, 10).value}
    for r in range(27, 35):                                   # la prévision du mois suivant
        code = feuille.cell(r, 2).value
        if code in lignes:
            lignes[code].update({"prevision": feuille.cell(r, 3).value, "mois_dernier": feuille.cell(r, 4).value,
                                 "an_dernier": feuille.cell(r, 5).value})
    titre = feuille["A25"].value or "Prévision"
    wb.close()

    # Le dernier mois complet : la meilleure méthode l'aurait-elle bien prévu ? Sinon, il est « inhabituel ».
    mois = mois_complets(table)
    dernier = mois.index.max()
    historique = mois[mois.index < dernier]
    for code, l in lignes.items():
        erreurs = [e if isinstance(e, (int, float)) else 9 for e in l["erreurs"]]
        i = erreurs.index(min(erreurs))
        try:
            prevu = float(METHODES[i][1](historique[[code]], dernier)[code])
            reel = float(mois.loc[dernier, code])
            l["ecart"] = (reel - prevu) / prevu if prevu else None
        except (KeyError, IndexError, ValueError):
            l["ecart"] = None
        l["erreur_habituelle"] = min(erreurs)
        if l["ecart"] is None:
            l["vigilance"] = "–"
        elif abs(l["ecart"]) <= l["erreur_habituelle"] * 1.5:
            l["vigilance"] = "Normal"
        elif abs(l["ecart"]) <= l["erreur_habituelle"] * 3:
            l["vigilance"] = "À surveiller"
        else:
            l["vigilance"] = "Inhabituel"
    return {"lignes": lignes, "titre": titre, "dernier_mois": dernier,
            "mois_cible": dernier + pd.DateOffset(months=1)}


# ======================================================================= la grippe (Réseau Sentinelles)

def lire_grippe(chemin=FICHIER_GRIPPE):
    """Le taux de syndromes grippaux, semaine par semaine (cas pour 100 000 habitants)."""
    lignes = Path(chemin).read_text(encoding="utf-8", errors="replace").splitlines()
    debut = next(i for i, l in enumerate(lignes)                                       # la ligne des titres
                 if not l.startswith("#") and "week" in l and "inc100" in l)
    table = pd.read_csv(io.StringIO("\n".join(lignes[debut:])))
    table["inc100"] = pd.to_numeric(table["inc100"], errors="coerce")
    table = table.dropna(subset=["inc100"])
    dates = pd.to_datetime(table["week"].astype(int).astype(str) + "1", format="%G%V%u")   # le lundi de la semaine
    return pd.Series(table["inc100"].values, index=dates).sort_index()


def telecharger_grippe():
    """Télécharge les dernières données du Réseau Sentinelles. Renvoie un message."""
    with urlopen(URL_GRIPPE, timeout=20) as reponse:
        contenu = reponse.read().decode("utf-8", errors="replace")
    return remplacer_grippe(contenu)


def remplacer_grippe(contenu):
    """Vérifie un fichier grippe, puis remplace l'ancien. Renvoie un message."""
    essai = DOSSIER / "donnees_externes" / "_grippe_essai.csv"
    essai.write_text(contenu, encoding="utf-8")
    try:
        serie = lire_grippe(essai)
    except Exception as erreur:
        essai.unlink()
        raise ValueError(f"ce n'est pas un fichier du Réseau Sentinelles ({erreur})")
    essai.replace(FICHIER_GRIPPE)
    return f"✓ Données grippe à jour : {len(serie)} semaines, jusqu'à la semaine du {serie.index[-1]:%d/%m/%Y}."


def etat_grippe(serie, date_ref):
    """La grippe à la date des dernières ventes : niveau, niveau habituel à cette période, tendance."""
    connues = serie[serie.index <= date_ref]
    if connues.empty:
        return None
    semaine, niveau = connues.index[-1], float(connues.iloc[-1])
    precedente = float(connues.iloc[-2]) if len(connues) > 1 else niveau
    numero = semaine.isocalendar()[1]
    memes = serie[[d.isocalendar()[1] == numero and d.year < semaine.year for d in serie.index]]
    habituel = float(memes.mean()) if len(memes) else None
    if (date_ref - semaine).days > 21:
        statut = "Données anciennes"                       # pas de chiffre récent : on ne conclut rien
    elif niveau >= SEUIL_EPIDEMIE:
        statut = "Épidémie en cours"
    elif habituel and niveau > 1.5 * habituel and niveau > precedente:
        statut = "En hausse"
    else:
        statut = "Calme"
    return {"semaine": semaine, "niveau": niveau, "precedente": precedente, "habituel": habituel, "statut": statut}


def test_grippe(mois, serie, lignes):
    """Teste si la grippe du mois précédent rend la prévision mensuelle plus juste.

    Pour chaque mois testé (à partir de janvier 2018), on corrige la prévision de la meilleure méthode
    par l'écart de la grippe à son niveau habituel, avec un coefficient appris uniquement sur les mois d'avant.
    Renvoie {code: (erreur sans grippe, erreur avec grippe)}.
    """
    grippe_mois = serie.resample("MS").mean()
    resultats = {}
    for code, l in lignes.items():
        erreurs = [e if isinstance(e, (int, float)) else 9 for e in l["erreurs"]]
        methode = METHODES[erreurs.index(min(erreurs))][1]
        rangs = []
        for cible in mois.index:
            historique = mois[mois.index < cible][[code]]
            if len(historique) < 25:
                continue
            try:
                prevu = float(methode(historique, cible)[code])
            except (KeyError, IndexError):
                continue
            avant = cible - pd.DateOffset(months=1)
            passes = grippe_mois[(grippe_mois.index.month == avant.month) & (grippe_mois.index < avant)]
            ecart = grippe_mois[avant] - passes.mean() if avant in grippe_mois.index and len(passes) else None
            rangs.append((cible, prevu, float(mois.loc[cible, code]), ecart))
        table = pd.DataFrame(rangs, columns=["mois", "prevu", "reel", "grippe"]).set_index("mois")
        table["residu"] = table["reel"] / table["prevu"] - 1
        sans, avec = [], []
        for m in table.index[table.index >= "2018-01-01"]:
            appris = table[table.index < m].dropna()
            coef = (appris["residu"] * appris["grippe"]).sum() / (appris["grippe"] ** 2).sum() if len(appris) > 5 else 0
            g = table.loc[m, "grippe"]
            corrige = table.loc[m, "prevu"] * (1 + coef * (0 if pd.isna(g) else g))
            sans.append(abs(table.loc[m, "prevu"] - table.loc[m, "reel"]) / table.loc[m, "reel"])
            avec.append(abs(corrige - table.loc[m, "reel"]) / table.loc[m, "reel"])
        if sans:
            resultats[code] = (sum(sans) / len(sans), sum(avec) / len(avec))
    return resultats


# ======================================================================= les événements notés par l'équipe

def charger_evenements():
    try:
        return json.loads(FICHIER_EVENEMENTS.read_text(encoding="utf-8"))
    except (FileNotFoundError, ValueError):
        return []


def enregistrer_evenements(evenements):
    FICHIER_EVENEMENTS.write_text(json.dumps(evenements, ensure_ascii=False, indent=2), encoding="utf-8")


def evenements_du_mois(evenements, mois, code=None):
    """Les événements qui touchent ce mois (et ce médicament, si on le précise)."""
    cle = f"{mois:%Y-%m}"
    return [e for e in evenements if e["debut"] <= cle <= e["fin"]
            and (code is None or not e["codes"] or code in e["codes"])]


# ======================================================================= le suivi des prévisions

def suivi_test(mois, lignes, evenements):
    """La prévision qu'on aurait faite chaque mois depuis janvier 2018, comparée au réel.

    Si une prévision a été enregistrée par l'interface pour ce mois, c'est elle qui est utilisée.
    """
    enregistrees = {}
    if FICHIER_SUIVI.exists():
        for _, r in pd.read_csv(FICHIER_SUIVI).iterrows():
            enregistrees[(r["mois"], r["code"])] = r["prevision"]
    rangs = []
    for cible in mois.index[mois.index >= "2018-01-01"]:
        historique = mois[mois.index < cible]
        for code, l in lignes.items():
            erreurs = [e if isinstance(e, (int, float)) else 9 for e in l["erreurs"]]
            i_nous, i_naive = erreurs.index(min(erreurs)), erreurs.index(min(erreurs[:2]))
            try:
                prevu = float(METHODES[i_nous][1](historique[[code]], cible)[code])
                naif = float(METHODES[i_naive][1](historique[[code]], cible)[code])
            except (KeyError, IndexError):
                continue
            source = "test"
            if (f"{cible:%Y-%m}", code) in enregistrees:
                prevu, source = float(enregistrees[(f"{cible:%Y-%m}", code)]), "enregistrée"
            reel = float(mois.loc[cible, code])
            rangs.append({"mois": cible, "code": code, "prevu": prevu, "reel": reel, "naif": naif,
                          "erreur": abs(prevu - reel) / reel if reel else None,
                          "erreur_naive": abs(naif - reel) / reel if reel else None,
                          "habituelle": min(erreurs), "source": source,
                          "evenement": bool(evenements_du_mois(evenements, cible, code))})
    return pd.DataFrame(rangs)


def enregistrer_prevision(p):
    """Garde une trace de la prévision du mois cible, pour la comparer au réel le mois suivant."""
    cle = f"{p['mois_cible']:%Y-%m}"
    deja = pd.read_csv(FICHIER_SUIVI) if FICHIER_SUIVI.exists() else pd.DataFrame(columns=["mois", "code"])
    nouvelles = [{"mois": cle, "code": code, "prevision": l["prevision"], "enregistree_le": f"{datetime.now():%Y-%m-%d}"}
                 for code, l in p["lignes"].items() if isinstance(l.get("prevision"), (int, float))
                 and not ((deja["mois"] == cle) & (deja["code"] == code)).any()]
    if nouvelles:
        pd.concat([deja, pd.DataFrame(nouvelles)]).to_csv(FICHIER_SUIVI, index=False)


# ======================================================================= la synthèse en PDF

def exporter_pdf(contenu):
    """Une page A4 en graphiques : fiabilité, prévision, évolution des ventes, grippe et événements."""
    import textwrap

    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    DOSSIER_RAPPORTS.mkdir(exist_ok=True)
    chemin = DOSSIER_RAPPORTS / f"synthese_{contenu['mois_cible']:%Y-%m}.pdf"
    bleu, bleu_pale, gris, noir = "#2563EB", "#BFD3FB", "#64748B", "#0F172A"
    vert, jaune, rouge = "#16A34A", "#F59E0B", "#DC2626"
    plt.rcParams.update({"font.size": 9, "axes.edgecolor": "#CBD5E1", "axes.labelcolor": gris,
                         "xtick.color": gris, "ytick.color": gris})
    fig = plt.figure(figsize=(8.27, 11.69))
    fig.text(0.06, 0.958, "Pharmacie · Synthèse mensuelle", fontsize=18, weight="bold", color=noir)
    fig.text(0.06, 0.94, contenu["sous_titre"], fontsize=9, color=gris)

    # 1. Les chiffres clés, et un anneau : combien de prévisions sont fiables
    fig.text(0.06, 0.895, "En bref", fontsize=12, weight="bold", color=bleu)
    for i, (valeur, libelle) in enumerate(contenu["chiffres"]):
        fig.text(0.06, 0.855 - i * 0.045, valeur, fontsize=16, weight="bold", color=noir)
        fig.text(0.06, 0.84 - i * 0.045, libelle, fontsize=8.5, color=gris)
    ax = fig.add_axes([0.5, 0.74, 0.22, 0.16])
    noms, nombres, couleurs = zip(*[(n, v, c) for n, v, c in
                                    [("Fiables", contenu["fiabilite"][0], vert),
                                     ("Gain marginal", contenu["fiabilite"][1], jaune),
                                     ("Non fiables", contenu["fiabilite"][2], rouge)] if v])
    ax.pie(nombres, colors=couleurs, startangle=90, counterclock=False,
           wedgeprops={"width": 0.38, "edgecolor": "white"})
    ax.text(0, 0, f"{contenu['fiabilite'][0]}/{sum(contenu['fiabilite'])}", ha="center", va="center",
            fontsize=15, weight="bold", color=noir)
    ax.set_title("Prévisions fiables", fontsize=10, color=noir)
    ax.legend([f"{n} ({v})" for n, v in zip(noms, nombres)], loc="center left", bbox_to_anchor=(1, 0.5),
              frameon=False, fontsize=8)

    # 2. La prévision du mois, comparée au même mois l'an dernier : un petit graphique par médicament
    #    (chacun sa propre échelle, puisqu'on ne compare jamais deux médicaments entre eux)
    p = contenu["prevision"]
    fig.text(0.06, 0.70, contenu["titre_prevision"], fontsize=12, weight="bold", color=bleu)
    largeur = 0.88 / max(len(p), 1)
    for i, l in enumerate(p):
        ax = fig.add_axes([0.06 + i * largeur + 0.015, 0.50, largeur - 0.03, 0.17])
        ax.bar([0], [l["an_dernier"]], width=0.6, color="#CBD5E1")
        ax.bar([1], [l["prevision"]], width=0.6, color=bleu,
               yerr=[[l["prevision"] - l["bas"]], [l["haut"] - l["prevision"]]],
               error_kw={"ecolor": noir, "capsize": 3, "elinewidth": 1})
        ax.text(0, l["an_dernier"], f"{l['an_dernier']:,.0f}".replace(",", " "), ha="center", va="bottom",
                fontsize=7.5, color=gris)
        ax.text(1, l["haut"], f"{l['prevision']:,.0f}".replace(",", " "), ha="center", va="bottom",
                fontsize=8.5, color=noir, weight="bold")
        ax.set_ylim(0, max(l["haut"], l["an_dernier"]) * 1.25)
        ax.set_xticks([0, 1])
        ax.set_xticklabels([contenu["libelle_court_an_dernier"], "prévu"], fontsize=7.5)
        ax.set_yticks([])
        ax.set_title(l["nom"], fontsize=9, color=noir, weight="bold")
        for cote in ("top", "right", "left"):
            ax.spines[cote].set_visible(False)
    fig.text(0.06, 0.465, textwrap.fill(contenu["note_prevision"], 125), fontsize=8, color=gris, va="top",
             linespacing=1.4)

    # 3. Les ventes du dernier mois, comparées à un an plus tôt
    e = contenu["evolution"]
    fig.text(0.06, 0.395, contenu["titre_ventes"], fontsize=12, weight="bold", color=bleu)
    ax = fig.add_axes([0.24, 0.2, 0.68, 0.18])
    valeurs = [v * 100 for _, v in e]
    ax.barh(range(len(e)), valeurs, color=[vert if v > 5 else rouge if v < -5 else "#94A3B8" for v in valeurs],
            height=0.6)
    for i, v in enumerate(valeurs):
        ax.text(v + (1.5 if v >= 0 else -1.5), i, f"{v:+.0f} %".replace("-", "−"), va="center",
                ha="left" if v >= 0 else "right", fontsize=8, color=noir)
    ax.set_yticks(range(len(e)))
    ax.set_yticklabels([n for n, _ in e], fontsize=8.5)
    ax.invert_yaxis()
    ax.axvline(0, color=gris, linewidth=0.8)
    ax.set_xticks([])
    marge = max(abs(v) for v in valeurs) * 1.35 + 5
    ax.set_xlim(-marge, marge)
    for cote in ("top", "right", "bottom", "left"):
        ax.spines[cote].set_visible(False)

    # 4. Grippe et événements
    fig.text(0.06, 0.165, "Grippe et événements", fontsize=12, weight="bold", color=bleu)
    texte = "\n".join(textwrap.fill(n, 110) for n in contenu["notes"])
    fig.text(0.06, 0.15, texte, fontsize=8.5, color=noir, va="top", linespacing=1.5)
    fig.text(0.06, 0.03, "Quantités dans l'unité du logiciel de caisse (non documentée) : un médicament est comparé "
                         "à lui-même, jamais à un autre.\nDétail et historique : classeur Pharma_Analyse.xlsx.",
             fontsize=7, color=gris)
    fig.savefig(chemin)
    plt.close(fig)
    return chemin


# ======================================================================= la fenêtre

def lancer_fenetre():
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    if sys.platform.startswith("win"):                        # texte net sur les écrans haute définition
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass

    # ---------- la palette
    FOND, CARTE, BORD = "#F4F6FA", "#FFFFFF", "#E2E8F0"
    TEXTE, DOUX, TRES_DOUX = "#0F172A", "#64748B", "#94A3B8"
    PRIMAIRE, PRIMAIRE_FONCE, PRIMAIRE_PALE = "#2563EB", "#1D4ED8", "#EFF4FF"
    MENU, MENU_ACTIF, MENU_TEXTE = "#0F1B2D", "#1E2F4A", "#CBD5E1"
    VERT, VERT_TEXTE = "#DCFCE7", "#166534"
    JAUNE, JAUNE_TEXTE = "#FEF3C7", "#92400E"
    ROUGE, ROUGE_TEXTE = "#FEE2E2", "#991B1B"
    # échelle de couleurs du tableau des ventes : en dessous (orange) → normal (blanc) → au-dessus (bleu)
    ECHELLE = [(-0.20, "#FDBA74", "En dessous de −20 %"), (-0.05, "#FFEDD5", "−20 % à −5 %"),
               (0.05, "#FFFFFF", "Normal (±5 %)"), (0.20, "#DBEAFE", "+5 % à +20 %"),
               (9e9, "#93C5FD", "Au-dessus de +20 %")]
    WIN = sys.platform.startswith("win")
    POLICE = "Segoe UI" if WIN else "DejaVu Sans"
    POLICE_TITRE = "Segoe UI Semibold" if WIN else "DejaVu Sans"

    def f(taille, gras=False):
        return (POLICE_TITRE if gras else POLICE, taille, "bold" if gras and not WIN else "normal")

    racine = tk.Tk()
    racine.title("Pharmacie · Suivi des ventes")
    racine.geometry("1320x840")
    racine.minsize(1100, 700)
    if sys.platform.startswith("win"):
        racine.state("zoomed")                                # plein écran au démarrage
    racine.configure(bg=FOND)

    style = ttk.Style()
    style.theme_use("clam")
    from tkinter import font as tkfont
    hauteur_ligne = tkfont.Font(root=racine, font=f(10)).metrics("linespace") + 16
    style.configure("Treeview", font=f(10), rowheight=hauteur_ligne, background=CARTE, fieldbackground=CARTE,
                    foreground=TEXTE, borderwidth=0)
    style.configure("Treeview.Heading", font=f(10, True), background=FOND, foreground=DOUX,
                    relief="flat", borderwidth=0, padding=(8, 8))
    style.map("Treeview", background=[("selected", PRIMAIRE_PALE)], foreground=[("selected", PRIMAIRE_FONCE)])
    style.map("Treeview.Heading", background=[("active", FOND)])
    style.configure("TCombobox", padding=6, arrowsize=14, bordercolor=BORD, lightcolor=CARTE, darkcolor=CARTE,
                    arrowcolor=DOUX, background=CARTE)
    style.map("TCombobox", fieldbackground=[("readonly", CARTE)], selectbackground=[("readonly", CARTE)],
              selectforeground=[("readonly", TEXTE)], bordercolor=[("focus", PRIMAIRE)])
    style.layout("Treeview", [("Treeview.treearea", {"sticky": "nswe"})])
    style.configure("Horizontal.TProgressbar", troughcolor=PRIMAIRE_PALE, background=PRIMAIRE, thickness=6,
                    borderwidth=0)
    style.configure("Vertical.TScrollbar", background=BORD, troughcolor=FOND, borderwidth=0, arrowsize=12)
    style.configure("Horizontal.TScrollbar", background=BORD, troughcolor=FOND, borderwidth=0, arrowsize=12)

    # ---------- petits composants réutilisés
    def carte(parent, padx=0, pady=0):
        c = tk.Frame(parent, bg=CARTE, highlightbackground=BORD, highlightthickness=1)
        return c

    def bouton(parent, texte, action, sorte="principal", petit=False):
        couleurs = {"principal": (PRIMAIRE, "#FFFFFF", PRIMAIRE_FONCE),
                    "secondaire": (CARTE, TEXTE, FOND),
                    "danger": (CARTE, ROUGE_TEXTE, ROUGE)}[sorte]
        b = tk.Label(parent, text=texte, bg=couleurs[0], fg=couleurs[1], font=f(9 if petit else 10, True),
                     padx=12 if petit else 16, pady=4 if petit else 8, cursor="hand2",
                     highlightbackground=BORD, highlightthickness=0 if sorte == "principal" else 1)
        b.actif = True

        def clic(_):
            if b.actif:
                action()
        b.bind("<Button-1>", clic)
        b.bind("<Enter>", lambda _: b.actif and b.configure(bg=couleurs[2]))
        b.bind("<Leave>", lambda _: b.configure(bg=couleurs[0]))

        def activer(oui, texte=None):
            b.actif = oui
            b.configure(fg=couleurs[1] if oui else TRES_DOUX, cursor="hand2" if oui else "arrow")
            if texte:
                b.configure(text=texte)
        b.activer = activer
        return b

    class Puce(tk.Label):
        """Un bouton à bascule (on/off), utilisé pour les filtres."""

        def __init__(self, parent, texte, actif=False, commande=None, coche=False):
            super().__init__(parent, text=texte, font=f(9), padx=10, pady=5, cursor="hand2",
                             highlightthickness=1)
            self.actif, self.commande, self.libelle, self.coche = actif, commande, texte, coche
            self.bind("<Button-1>", lambda _: self.basculer())
            self.dessiner()

        def dessiner(self):
            if self.actif:
                self.configure(bg=PRIMAIRE, fg="#FFFFFF", highlightbackground=PRIMAIRE)
            else:
                self.configure(bg=CARTE, fg=DOUX if self.coche else TEXTE, highlightbackground=BORD)
            if self.coche:                                       # case à cocher : ✓ quand c'est sélectionné
                self.configure(text=("✓  " if self.actif else "○  ") + self.libelle)

        def basculer(self, valeur=None, appeler=True):
            self.actif = (not self.actif) if valeur is None else valeur
            self.dessiner()
            if appeler and self.commande:
                self.commande(self)

    def segments(parent, options, valeur, commande):
        """Un sélecteur à choix unique (comme des onglets) ; renvoie le cadre."""
        cadre = tk.Frame(parent, bg=FOND, highlightbackground=BORD, highlightthickness=1)
        puces = {}

        def choisir(p):
            for o, autre in puces.items():
                autre.basculer(o == p.cget("text"), appeler=False)
            valeur.set(p.cget("text"))
            commande()
        for o in options:
            puces[o] = Puce(cadre, o, actif=(o == valeur.get()), commande=choisir)
            puces[o].configure(highlightthickness=0)
            puces[o].pack(side="left")
        return cadre

    def titre_page(parent, titre, sous_titre):
        tk.Label(parent, text=titre, bg=FOND, fg=TEXTE, font=f(18, True)).pack(anchor="w", padx=28, pady=(22, 0))
        tk.Label(parent, text=sous_titre, bg=FOND, fg=DOUX, font=f(10), justify="left", wraplength=980).pack(anchor="w", padx=28,
                                                                                            pady=(2, 12))

    def grille(parent, entetes, lignes, largeurs, aligns=None):
        """Un tableau en couleurs. lignes = listes de cases (texte, fond, couleur du texte, gras ?)."""
        for w in parent.winfo_children():
            w.destroy()
        aligns = aligns or ["w"] + ["e"] * (len(entetes) - 1)
        for j, (t, larg) in enumerate(zip(entetes, largeurs)):
            parent.grid_columnconfigure(j, minsize=larg * 7)          # largeur minimale ; s'élargit si besoin
            tk.Label(parent, text=t, bg=FOND, fg=DOUX, font=f(9, True), anchor=aligns[j],
                     padx=8, pady=8).grid(row=0, column=j, sticky="nsew")
        for i, ligne in enumerate(lignes, start=1):
            for j, case in enumerate(ligne):
                texte, fond, encre, gras = (list(case) + [CARTE, TEXTE, False])[:4] if isinstance(case, tuple) \
                    else (case, CARTE, TEXTE, False)
                tk.Label(parent, text=texte, bg=fond, fg=encre, font=f(10, gras),
                         anchor=aligns[j], padx=8, pady=9).grid(row=i, column=j, sticky="nsew", padx=(0, 1),
                                                                pady=(0, 1))
        parent.configure(bg=BORD)

    def legende(parent, elements, titre=None):
        """Une légende : une pastille de couleur et son explication, sur une ligne."""
        cadre = tk.Frame(parent, bg=CARTE)
        if titre:
            tk.Label(cadre, text=titre, bg=CARTE, fg=TEXTE, font=f(9, True)).pack(side="left", padx=(0, 10))
        for couleur, texte in elements:
            tk.Label(cadre, text="  ", bg=couleur, highlightbackground=BORD, highlightthickness=1).pack(side="left")
            tk.Label(cadre, text=texte, bg=CARTE, fg=DOUX, font=f(9)).pack(side="left", padx=(5, 14))
        return cadre

    def defilable(parent, horizontal=True):
        """Une zone qui défile (molette, ou Maj + molette pour le côté) ; renvoie le cadre intérieur."""
        boite = tk.Frame(parent, bg=CARTE)
        toile = tk.Canvas(boite, bg=CARTE, highlightthickness=0)
        dedans = tk.Frame(toile, bg=CARTE)
        dedans.bind("<Configure>", lambda _: toile.configure(scrollregion=toile.bbox("all")))
        toile.create_window((0, 0), window=dedans, anchor="nw")
        vs = ttk.Scrollbar(boite, orient="vertical", command=toile.yview)
        toile.configure(yscrollcommand=vs.set)
        if horizontal:
            hs = ttk.Scrollbar(boite, orient="horizontal", command=toile.xview)
            toile.configure(xscrollcommand=hs.set)
            hs.pack(side="bottom", fill="x")
        vs.pack(side="right", fill="y")
        toile.pack(side="left", fill="both", expand=True)

        def molette(e):
            pas = -1 if (e.delta > 0 or e.num == 4) else 1
            (toile.xview_scroll if e.state & 1 else toile.yview_scroll)(pas * 2, "units")
        toile.bind("<Enter>", lambda _: (toile.bind_all("<MouseWheel>", molette),
                                         toile.bind_all("<Button-4>", molette), toile.bind_all("<Button-5>", molette)))
        toile.bind("<Leave>", lambda _: (toile.unbind_all("<MouseWheel>"), toile.unbind_all("<Button-4>"),
                                         toile.unbind_all("<Button-5>")))
        return boite, dedans

    # ---------- la structure : menu à gauche, page à droite
    menu = tk.Frame(racine, bg=MENU)
    menu.pack(side="left", fill="y")
    contenu = tk.Frame(racine, bg=FOND)
    contenu.pack(side="left", fill="both", expand=True)

    tk.Label(menu, text="✚  Pharmacie", bg=MENU, fg="#FFFFFF", font=f(15, True)).pack(anchor="w", padx=22,
                                                                                    pady=(26, 0))
    tk.Label(menu, text="Suivi des ventes", bg=MENU, fg=TRES_DOUX, font=f(10)).pack(anchor="w", padx=24,
                                                                                  pady=(0, 26))
    pages, boutons_menu = {}, {}

    def afficher(nom):
        for n, p in pages.items():
            p.pack_forget()
            boutons_menu[n].configure(bg=MENU, fg=MENU_TEXTE)
        pages[nom].pack(fill="both", expand=True)
        boutons_menu[nom].configure(bg=MENU_ACTIF, fg="#FFFFFF")

    for nom, icone in [("Mise à jour", "⟳"), ("Ventes", "▦"), ("Prévision", "◔"), ("Suivi", "✓"),
                       ("Événements", "⚑"), ("Catégories", "☰")]:
        pages[nom] = tk.Frame(contenu, bg=FOND)
        b = tk.Label(menu, text=f"   {icone}    {nom}", bg=MENU, fg=MENU_TEXTE, font=f(11), anchor="w",
                     pady=11, cursor="hand2")
        b.pack(fill="x", padx=12, pady=1)
        b.bind("<Button-1>", lambda _, n=nom: afficher(n))
        b.bind("<Enter>", lambda _, b=b: b.cget("bg") == MENU and b.configure(bg="#16243A"))
        b.bind("<Leave>", lambda _, b=b, n=nom: b.configure(bg=MENU_ACTIF if pages[n].winfo_ismapped() else MENU))
        boutons_menu[nom] = b

    pied_menu = tk.Label(menu, text="", bg=MENU, fg=TRES_DOUX, font=f(9), justify="left", anchor="w")
    pied_menu.pack(side="bottom", fill="x", padx=24, pady=20)

    donnees = {}          # ventes, codes, prévision… remplis par rafraichir()

    # ================================================================ page 1 : Mise à jour
    page = pages["Mise à jour"]
    titre_page(page, "Mise à jour du classeur",
               "Chaque mois : importez les exports reçus du client, puis remettez le classeur Excel à jour.")

    def etape(numero, titre, texte):
        c = carte(page)
        c.pack(fill="x", padx=28, pady=6)
        tk.Label(c, text=str(numero), bg=PRIMAIRE_PALE, fg=PRIMAIRE, font=f(15, True), width=3).pack(
            side="left", fill="y")
        corps = tk.Frame(c, bg=CARTE)
        corps.pack(side="left", fill="both", expand=True, padx=18, pady=14)
        tk.Label(corps, text=titre, bg=CARTE, fg=TEXTE, font=f(12, True)).pack(anchor="w")
        tk.Label(corps, text=texte, bg=CARTE, fg=DOUX, font=f(10), justify="left", wraplength=440).pack(anchor="w")
        zone = tk.Frame(c, bg=CARTE)
        zone.pack(side="right", padx=18)
        return zone

    zone1 = etape(1, "Importer les nouveaux exports",
                  "Choisissez les fichiers CSV reçus du client, quel que soit leur nom : ils sont reconnus et rangés.")
    zone2 = etape(2, "Mettre à jour le classeur",
                  "Fermez Excel, puis lancez la mise à jour (environ 30 secondes). Tout le classeur est recalculé, "
                  "les nouveaux médicaments sont ajoutés et une copie est archivée.")
    zone3 = etape(3, "Consulter et partager",
                  "Le classeur Excel garde le détail et l'historique (une copie par mois pendant 3 ans, dans Archives). "
                  "La synthèse PDF résume le mois pour le propriétaire.")

    progression = ttk.Progressbar(page, mode="indeterminate", style="Horizontal.TProgressbar")
    cadre_journal = carte(page)
    tk.Label(cadre_journal, text="Journal", bg=CARTE, fg=DOUX, font=f(9, True)).pack(anchor="w", padx=14,
                                                                                   pady=(10, 0))
    journal = tk.Text(cadre_journal, height=10, font=("Consolas" if WIN else "DejaVu Sans Mono", 10), bg=CARTE,
                      fg=TEXTE, relief="flat", highlightthickness=0, wrap="word", padx=14, pady=8)
    journal.pack(fill="both", expand=True)

    def ecrire(texte):
        journal.configure(state="normal")
        journal.insert("end", texte + "\n")
        journal.see("end")
        journal.configure(state="disabled")

    def action_importer():
        chemins = filedialog.askopenfilenames(title="Choisissez les nouveaux exports (CSV)",
                                              filetypes=[("Fichiers CSV", "*.csv")])
        if not chemins:
            return
        ecrire("— Import des exports —")
        for ligne in importer(chemins):
            ecrire(ligne)
        try:
            ecrire(verifier())
        except Exception as erreur:
            ecrire(f"! Contrôle impossible : {erreur}")
        ecrire("→ Étape suivante : Mettre à jour le classeur.")

    def action_mettre_a_jour():
        if excel_ouvert():
            messagebox.showwarning("Excel est ouvert", "Le classeur est ouvert dans Excel.\nFermez-le, puis cliquez à nouveau.")
            return
        bouton_maj.activer(False, "Mise à jour en cours…")
        progression.pack(fill="x", padx=28, pady=(8, 0), before=cadre_journal)
        progression.start(12)
        ecrire("— Mise à jour du classeur —")

        def travail():
            try:
                reussi, message = mettre_a_jour()
            except Exception as erreur:
                reussi, message = False, str(erreur)
            try:                                               # au passage : les derniers chiffres de la grippe
                message += "\n" + telecharger_grippe()
            except Exception:
                message += "\n   (Données grippe non mises à jour : pas d'accès à Internet. Ce n'est pas bloquant.)"

            def fin():
                progression.stop()
                progression.pack_forget()
                bouton_maj.activer(True, "Mettre à jour le classeur")
                if reussi:
                    ecrire(message)
                    rafraichir()
                    ecrire(f"   Données jusqu'au {donnees['dernier_jour']:%d/%m/%Y}. Voir la page Prévision.")
                else:
                    ecrire(f"✗ La mise à jour a échoué : {message}\n"
                           "   Le plus souvent : un CSV au mauvais format. Ouvrez classeur.ipynb dans Jupyter pour le détail.")
            racine.after(0, fin)
        threading.Thread(target=travail, daemon=True).start()

    bouton(zone1, "Choisir les fichiers CSV…", action_importer).pack()
    bouton_maj = bouton(zone2, "Mettre à jour le classeur", action_mettre_a_jour)
    bouton_maj.pack()
    bouton(zone3, "Ouvrir Excel", lambda: ouvrir(CLASSEUR)).pack(side="left", padx=4)
    bouton(zone3, "Synthèse PDF", lambda: action_pdf(), sorte="secondaire").pack(side="left", padx=4)

    def ouvrir_archives():
        dossier = CLASSEUR.parent / "archives"
        if not dossier.exists():
            messagebox.showinfo("Archives", "Pas encore d'archive : elles sont créées à chaque mise à jour du classeur.")
            return
        ouvrir(dossier)
    bouton(zone3, "Archives", ouvrir_archives, sorte="secondaire").pack(side="left", padx=4)

    cadre_journal.pack(fill="both", expand=True, padx=28, pady=(12, 24))

    # ================================================================ page 2 : Ventes
    page = pages["Ventes"]
    titre_page(page, "Ventes : la synthèse",
               "Choisissez une période et des médicaments ; ajoutez les filtres dont vous avez besoin. "
               "Le détail jour par jour reste dans Excel.")

    filtres = carte(page)
    filtres.pack(fill="x", padx=28)
    etat = {"debut": None, "fin": None, "codes": set(), "jours": None, "heures": None, "categorie": None}
    mode = tk.StringVar(value="Mois")

    def ligne_filtre(titre):
        ligne = tk.Frame(filtres, bg=CARTE)
        ligne.pack(fill="x", padx=18, pady=(12, 0))
        tk.Label(ligne, text=titre, bg=CARTE, fg=DOUX, font=f(9, True), width=16, anchor="w").pack(side="left")
        return ligne

    # -- période : raccourcis, un mois précis, ou deux dates
    l_periode = ligne_filtre("PÉRIODE")
    raccourcis = {}
    cadre_rac = tk.Frame(l_periode, bg=FOND, highlightbackground=BORD, highlightthickness=1)
    cadre_rac.pack(side="left")
    l_periode2 = ligne_filtre("")
    tk.Label(l_periode2, text="Un mois précis", bg=CARTE, fg=DOUX, font=f(9)).pack(side="left", padx=(0, 6))
    choix_mois = ttk.Combobox(l_periode2, state="readonly", width=16, font=f(10))
    choix_mois.pack(side="left")
    tk.Label(l_periode2, text="ou des dates : du", bg=CARTE, fg=DOUX, font=f(9)).pack(side="left", padx=(18, 6))
    saisie_debut = tk.Entry(l_periode2, width=11, font=f(10), relief="flat", highlightthickness=1,
                            highlightbackground=BORD, highlightcolor=PRIMAIRE)
    saisie_debut.pack(side="left", ipady=4)
    tk.Label(l_periode2, text="au", bg=CARTE, fg=DOUX, font=f(9)).pack(side="left", padx=6)
    saisie_fin = tk.Entry(l_periode2, width=11, font=f(10), relief="flat", highlightthickness=1,
                          highlightbackground=BORD, highlightcolor=PRIMAIRE)
    saisie_fin.pack(side="left", ipady=4)
    tk.Label(l_periode2, text="(JJ/MM/AAAA, puis Entrée)", bg=CARTE, fg=TRES_DOUX, font=f(8)).pack(side="left",
                                                                                                padx=6)

    def poser_periode(debut, fin, depuis=None):
        etat["debut"], etat["fin"] = debut, fin
        for s, d in ((saisie_debut, debut), (saisie_fin, fin)):
            s.delete(0, "end")
            s.insert(0, f"{d:%d/%m/%Y}")
            s.configure(highlightbackground=BORD)
        for nom, p in raccourcis.items():
            p.basculer(nom == depuis, appeler=False)
        if depuis != "mois":
            choix_mois.set("")
        afficher_ventes()

    def periode_raccourci(nom):
        dernier = donnees["dernier_mois"]                     # le dernier mois complet
        fin_mois = dernier + pd.offsets.MonthEnd(0)
        debut = {"Dernier mois": dernier, "3 derniers mois": dernier - pd.DateOffset(months=2),
                 "12 derniers mois": dernier - pd.DateOffset(months=11)}.get(nom)
        if nom == "Tout l'historique":
            poser_periode(donnees["premier_jour"], donnees["dernier_jour"], nom)
        else:
            poser_periode(debut, fin_mois, nom)

    for nom in ["Dernier mois", "3 derniers mois", "12 derniers mois", "Tout l'historique"]:
        raccourcis[nom] = Puce(cadre_rac, nom, commande=lambda p: periode_raccourci(p.cget("text")))
        raccourcis[nom].configure(highlightthickness=0)
        raccourcis[nom].pack(side="left")

    def mois_choisi(_=None):
        texte = choix_mois.get()
        if not texte:
            return
        nom_mois, annee = texte.rsplit(" ", 1)
        debut = pd.Timestamp(int(annee), MOIS_FR.index(nom_mois) + 1, 1)
        poser_periode(debut, min(debut + pd.offsets.MonthEnd(0), donnees["dernier_jour"]), "mois")
    choix_mois.bind("<<ComboboxSelected>>", mois_choisi)

    def dates_saisies(_=None):
        try:
            debut = pd.to_datetime(saisie_debut.get().strip(), format="%d/%m/%Y")
            fin = pd.to_datetime(saisie_fin.get().strip(), format="%d/%m/%Y")
            if fin < debut:
                raise ValueError
        except ValueError:
            for s in (saisie_debut, saisie_fin):
                s.configure(highlightbackground="#EF4444")
            message_ventes.configure(text="Dates non comprises : écrivez-les comme 01/09/2019, la fin après le début.",
                                     fg=ROUGE_TEXTE)
            return
        poser_periode(debut, fin)
    for s in (saisie_debut, saisie_fin):
        s.bind("<Return>", dates_saisies)

    # -- médicaments
    l_codes = ligne_filtre("MÉDICAMENTS")
    cadre_codes = tk.Frame(l_codes, bg=CARTE)
    cadre_codes.pack(side="left")
    l_codes2 = ligne_filtre("")
    l_codes2.pack_configure(pady=(6, 0))
    puces_codes = {}

    def code_bascule(p):
        etat["codes"] = {c for c, pc in puces_codes.items() if pc.actif}
        afficher_ventes()

    def tous_les_codes(oui):
        for pc in puces_codes.values():
            pc.basculer(oui, appeler=False)
        code_bascule(None)

    # -- affichage
    l_mode = ligne_filtre("AFFICHER PAR")
    segments(l_mode, ["Mois", "Année", "Jour de la semaine", "Heure"], mode, lambda: afficher_ventes()).pack(
        side="left")

    # -- filtres que l'on ajoute soi-même
    zone_ajouts = tk.Frame(filtres, bg=CARTE)
    zone_ajouts.pack(fill="x")
    l_ajout = tk.Frame(filtres, bg=CARTE)
    l_ajout.pack(fill="x", padx=18, pady=(10, 14))
    ajoutes = {}

    def ajouter_filtre(nom):
        if nom in ajoutes:
            return
        ligne = tk.Frame(zone_ajouts, bg=PRIMAIRE_PALE)
        ligne.pack(fill="x", padx=18, pady=(10, 0))
        etiquette = {"Jours de la semaine": "JOURS", "Plage horaire": "HEURES", "Catégorie": "CATÉGORIE"}[nom]
        tk.Label(ligne, text=etiquette, bg=PRIMAIRE_PALE, fg=PRIMAIRE_FONCE, font=f(9, True), width=16,
                 anchor="w").pack(side="left", padx=(10, 0), pady=8)
        if nom == "Jours de la semaine":
            puces = []

            def jours_change(_):
                etat["jours"] = [i for i, p in enumerate(puces) if p.actif] or None
                afficher_ventes()
            for j in JOURS:
                p = Puce(ligne, j[:3] + ".", actif=True, commande=jours_change)
                p.pack(side="left", padx=2)
                puces.append(p)
            etat["jours"] = list(range(7))
        elif nom == "Plage horaire":
            heures = [f"{h} h" for h in range(24)]
            de = ttk.Combobox(ligne, values=heures, state="readonly", width=6, font=f(10))
            a = ttk.Combobox(ligne, values=heures, state="readonly", width=6, font=f(10))
            de.set("8 h")
            a.set("20 h")
            tk.Label(ligne, text="de", bg=PRIMAIRE_PALE, fg=DOUX, font=f(9)).pack(side="left", padx=4)
            de.pack(side="left")
            tk.Label(ligne, text="à", bg=PRIMAIRE_PALE, fg=DOUX, font=f(9)).pack(side="left", padx=4)
            a.pack(side="left")
            tk.Label(ligne, text="(heure de fin incluse)", bg=PRIMAIRE_PALE, fg=TRES_DOUX, font=f(8)).pack(
                side="left", padx=6)

            def heures_change(_=None):
                h1, h2 = int(de.get().split()[0]), int(a.get().split()[0])
                etat["heures"] = (min(h1, h2), max(h1, h2))
                afficher_ventes()
            de.bind("<<ComboboxSelected>>", heures_change)
            a.bind("<<ComboboxSelected>>", heures_change)
            etat["heures"] = (8, 20)
        elif nom == "Catégorie":
            choix = ttk.Combobox(ligne, values=sorted({cat.categorie(c) for c in donnees["codes"]}),
                                 state="readonly", width=28, font=f(10))
            choix.pack(side="left")
            choix.current(0)

            def cat_change(_=None):
                etat["categorie"] = choix.get()
                afficher_ventes()
            choix.bind("<<ComboboxSelected>>", cat_change)
            etat["categorie"] = choix.get()

        def retirer():
            ligne.destroy()
            del ajoutes[nom]
            etat[{"Jours de la semaine": "jours", "Plage horaire": "heures", "Catégorie": "categorie"}[nom]] = None
            maj_menu_ajout()
            afficher_ventes()
        croix = tk.Label(ligne, text="✕  Retirer", bg=PRIMAIRE_PALE, fg=PRIMAIRE, font=f(9, True), cursor="hand2",
                         padx=10)
        croix.pack(side="right", padx=6)
        croix.bind("<Button-1>", lambda _: retirer())
        ajoutes[nom] = ligne
        maj_menu_ajout()
        afficher_ventes()

    bouton_ajout = tk.Menubutton(l_ajout, text="＋  Ajouter un filtre", bg=CARTE, fg=PRIMAIRE, font=f(10, True),
                                 relief="flat", cursor="hand2", activebackground=PRIMAIRE_PALE,
                                 activeforeground=PRIMAIRE, highlightthickness=1, highlightbackground=BORD,
                                 padx=12, pady=5)
    menu_ajout = tk.Menu(bouton_ajout, tearoff=0, font=f(10))
    bouton_ajout.configure(menu=menu_ajout)
    bouton_ajout.pack(side="left")
    bouton(l_ajout, "Tout réinitialiser", lambda: reinitialiser(), sorte="secondaire", petit=True).pack(
        side="left", padx=10)
    message_ventes = tk.Label(l_ajout, text="", bg=CARTE, fg=DOUX, font=f(9))
    message_ventes.pack(side="left", padx=10)

    def maj_menu_ajout():
        menu_ajout.delete(0, "end")
        possibles = ["Jours de la semaine", "Catégorie"] + (["Plage horaire"] if donnees.get("horaire") else [])
        for nom in possibles:
            menu_ajout.add_command(label=nom, command=lambda n=nom: ajouter_filtre(n),
                                   state="disabled" if nom in ajoutes else "normal")

    def reinitialiser():
        for nom in list(ajoutes):
            ajoutes[nom].destroy()
            del ajoutes[nom]
        etat.update({"jours": None, "heures": None, "categorie": None})
        for pc in puces_codes.values():
            pc.basculer(True, appeler=False)
        etat["codes"] = set(puces_codes)
        mode.set("Mois")
        maj_menu_ajout()
        periode_raccourci("12 derniers mois")

    # -- le résultat
    resultat = carte(page)
    resultat.pack(fill="both", expand=True, padx=28, pady=(12, 0))
    resume = tk.Label(resultat, text="", bg=CARTE, fg=TEXTE, font=f(11, True), anchor="w", justify="left")
    resume.pack(fill="x", padx=18, pady=(14, 8))
    boite_ventes, grille_ventes = defilable(resultat)
    boite_ventes.pack(fill="both", expand=True, padx=18)
    bas_ventes = tk.Frame(resultat, bg=CARTE)
    bas_ventes.pack(fill="x", padx=18, pady=(10, 14))
    legende(bas_ventes, [(c, t) for _, c, t in ECHELLE], "Couleur :").pack(anchor="w")
    note_ventes = tk.Label(bas_ventes, text="", bg=CARTE, fg=DOUX, font=f(9), justify="left", anchor="w",
                           wraplength=980)
    note_ventes.pack(fill="x", pady=(6, 0))

    def couleur(ecart):
        return next(c for borne, c, _ in ECHELLE if ecart <= borne)

    def afficher_ventes():
        if not donnees or etat["debut"] is None:
            return
        codes = [c for c in donnees["codes"] if c in etat["codes"]
                 and (etat["categorie"] is None or cat.categorie(c) == etat["categorie"])]
        message_ventes.configure(text="", fg=DOUX)
        if not codes:
            grille(grille_ventes, ["Aucun médicament sélectionné"], [], [40])
            resume.configure(text="Choisissez au moins un médicament.")
            return
        m = mode.get()
        if m == "Heure" and not donnees["horaire"]:
            message_ventes.configure(text="L'export horaire manque : l'affichage par heure est impossible.",
                                     fg=ROUGE_TEXTE)
            return
        cols, valeurs, totaux, an_dernier, nb_jours, incompletes = synthese_ventes(
            donnees["ventes"], codes, etat["debut"], etat["fin"], m, etat["jours"], etat["heures"])
        if nb_jours == 0:
            grille(grille_ventes, ["Aucune vente avec ces filtres"], [], [40])
            resume.configure(text="Aucun jour ne correspond à ces filtres.")
            return
        moyenne = m in ("Jour de la semaine", "Heure")
        evenements = charger_evenements()
        periodes = []
        if m == "Mois":                                           # le 1er jour de chaque mois affiché
            periodes = list(pd.period_range(etat["debut"], etat["fin"], freq="M").to_timestamp())
            periodes = [p for p in periodes if any(f"{MOIS_COURT[p.month - 1]} {p.year % 100:02d}" == col
                                                   for col in cols)]
        lignes = []
        for c in codes:
            vals = valeurs[c]
            plus_grand = max(vals + [totaux[c]]) if vals else 0
            dec = 0 if plus_grand >= 100 or not moyenne else 1 if plus_grand >= 5 else 2
            ref = sum(vals) / len(vals) if vals else 0
            ligne = [(nom_court(c), CARTE, TEXTE, True)]
            for k, v in enumerate(vals):
                ecart = (v / ref - 1) if ref else 0
                drapeau = " ⚑" if m == "Mois" and evenements_du_mois(evenements, periodes[k], c) else ""
                ligne.append((nombre(v, dec) + drapeau, couleur(ecart), TEXTE, False))
            ligne.append((nombre(totaux[c], dec), FOND, TEXTE, True))
            if an_dernier is not None and an_dernier[c]:
                evol = totaux[c] / an_dernier[c] - 1
                fond, encre = (VERT, VERT_TEXTE) if evol > 0.05 else (ROUGE, ROUGE_TEXTE) if evol < -0.05 \
                    else (CARTE, DOUX)
                ligne.append((f"{'▲' if evol > 0 else '▼'} {evol * 100:+.0f} %".replace("-", "−"), fond, encre, True))
            else:
                ligne.append(("–", CARTE, TRES_DOUX, False))
            lignes.append(ligne)
        entetes = ["Médicament"] + [c + ("*" if c in incompletes else "") for c in cols] + \
                  ["Moyenne / jour" if moyenne else "Total", "vs an dernier"]
        largeurs = [16] + [7] * len(cols) + [12, 12]
        grille(grille_ventes, entetes, lignes, largeurs)

        filtres_actifs = []
        if etat["jours"] is not None and len(etat["jours"]) < 7:
            filtres_actifs.append(", ".join(JOURS[j][:3] + "." for j in etat["jours"]))
        if etat["heures"]:
            filtres_actifs.append(f"de {etat['heures'][0]} h à {etat['heures'][1]} h")
        if etat["categorie"]:
            filtres_actifs.append(etat["categorie"])
        resume.configure(text=f"Du {etat['debut']:%d/%m/%Y} au {etat['fin']:%d/%m/%Y}  ·  {nb_jours} jours de ventes"
                              + (f"  ·  {' · '.join(filtres_actifs)}" if filtres_actifs else ""))
        notes = ["Chaque case est comparée à la moyenne du même médicament sur les colonnes affichées : "
                 "on compare un médicament à lui-même, jamais deux médicaments entre eux (l'unité n'est pas connue)."]
        if moyenne:
            notes.append("Valeurs : ventes moyennes par jour.")
        if incompletes:
            notes.append("* Période coupée par les dates choisies : son total ne couvre pas tous ses jours.")
        notes.append("« vs an dernier » : même période et mêmes filtres, un an plus tôt.")
        notes_ev = [e for e in evenements if e["debut"] <= f"{etat['fin']:%Y-%m}" and e["fin"] >= f"{etat['debut']:%Y-%m}"]
        for e in notes_ev[:4]:
            qui = ", ".join(nom_court(c) for c in e["codes"]) if e["codes"] else "tous les médicaments"
            notes.append(f"⚑ {e['debut']} → {e['fin']} · {e['type']} ({qui}) : {e['texte']}")
        note_ventes.configure(text="\n".join(notes))

    # ================================================================ page 3 : Prévision
    page = pages["Prévision"]
    titre_prev = tk.Label(page, text="Prévision", bg=FOND, fg=TEXTE, font=f(18, True))
    titre_prev.pack(anchor="w", padx=28, pady=(22, 0))
    tk.Label(page, text="Combien commander le mois prochain, et peut-on s'y fier ? "
                        "Les méthodes ont été testées sur les mois passés (onglet Prévision du classeur).",
             bg=FOND, fg=DOUX, font=f(10)).pack(anchor="w", padx=28, pady=(2, 12))

    bandeau = tk.Frame(page, bg=FOND)
    bandeau.pack(fill="x", padx=24)
    chiffres = {}
    for cle, libelle in [("cible", "Mois prévu"), ("fiables", "Prévisions fiables"),
                         ("grippe", "Grippe (Sentinelles)"), ("alertes", "Dernier mois inhabituel")]:
        c = carte(bandeau)
        c.pack(side="left", fill="both", expand=True, padx=4)
        tk.Label(c, text=libelle.upper(), bg=CARTE, fg=DOUX, font=f(8, True)).pack(anchor="w", padx=16,
                                                                                  pady=(9, 0))
        chiffres[cle] = tk.Label(c, text="–", bg=CARTE, fg=TEXTE, font=f(15, True))
        chiffres[cle].pack(anchor="w", padx=16, pady=(0, 9))
        if cle == "grippe":                                    # le détail de la grippe, en petit
            chiffres[cle].pack_configure(pady=0)
            detail_grippe = tk.Label(c, text="", bg=CARTE, fg=DOUX, font=f(8))
            detail_grippe.pack(anchor="w", padx=16, pady=(0, 8))

    filtres_prev = carte(page)
    filtres_prev.pack(fill="x", padx=28, pady=(12, 0))
    lp = tk.Frame(filtres_prev, bg=CARTE)
    lp.pack(fill="x", padx=18, pady=(12, 4))
    tk.Label(lp, text="FILTRES", bg=CARTE, fg=DOUX, font=f(9, True), width=10, anchor="w").pack(side="left")
    tk.Label(lp, text="Catégorie", bg=CARTE, fg=DOUX, font=f(9)).pack(side="left")
    choix_cat_prev = ttk.Combobox(lp, state="readonly", width=22, font=f(10))
    choix_cat_prev.pack(side="left", padx=(6, 16))
    fiabilite = tk.StringVar(value="Toutes")
    tk.Label(lp, text="Fiabilité", bg=CARTE, fg=DOUX, font=f(9)).pack(side="left")
    segments(lp, ["Toutes", "Fiables", "Non fiables"], fiabilite, lambda: afficher_prevision()).pack(
        side="left", padx=(6, 16))
    puce_alertes = Puce(lp, "⚠  Mois inhabituels seulement", commande=lambda p: afficher_prevision())
    puce_alertes.pack(side="left")
    choix_cat_prev.bind("<<ComboboxSelected>>", lambda _: afficher_prevision())
    lp2 = tk.Frame(filtres_prev, bg=CARTE)
    lp2.pack(fill="x", padx=18, pady=(6, 12))
    tk.Label(lp2, text="PRUDENCE", bg=CARTE, fg=DOUX, font=f(9, True), width=10, anchor="w").pack(side="left")
    prudence = tk.StringVar(value="Prudent")
    segments(lp2, ["Au plus juste", "Prudent", "Très prudent"], prudence, lambda: afficher_prevision()).pack(
        side="left")
    tk.Label(lp2, text="= prévision + 0, 1 ou 2 fois l'erreur habituelle",
             bg=CARTE, fg=TRES_DOUX, font=f(8)).pack(side="left", padx=10)
    bouton(lp2, "⤓  Synthèse PDF", lambda: action_pdf(), petit=True).pack(side="right")


    res_prev = carte(page)
    res_prev.pack(fill="both", expand=True, padx=28, pady=(12, 24))
    grille_prev = tk.Frame(res_prev, bg=BORD)
    grille_prev.pack(anchor="w", padx=18, pady=(16, 8))
    leg_prev = tk.Frame(res_prev, bg=CARTE)
    leg_prev.pack(fill="x", padx=18, pady=(4, 14))
    legende(leg_prev, [(VERT, "Fiable"), (JAUNE, "Gain marginal"), (ROUGE, "Non fiable")],
            "Fiabilité :").pack(anchor="w")
    legende(leg_prev, [(VERT, "Normal"), (JAUNE, "À surveiller"), (ROUGE, "Inhabituel"),
                       (PRIMAIRE_PALE, "⚑ Expliqué par un événement noté")],
            "Dernier mois :").pack(anchor="w", pady=(6, 0))
    tk.Label(leg_prev, text="« Chaque semaine » : pas de chiffre fiable, on commande chaque semaine selon les ventes.",
             bg=CARTE, fg=DOUX, font=f(9)).pack(anchor="w", pady=(8, 0))

    def a_prevoir(l):
        """La quantité à prévoir selon le niveau de prudence choisi."""
        k = {"Au plus juste": 0, "Prudent": 1, "Très prudent": 2}[prudence.get()]
        prev = l.get("prevision")
        return prev * (1 + k * (l["notre"] or 0)) if isinstance(prev, (int, float)) else None

    def case_ecart(code, l):
        dernier = donnees["prevision"]["dernier_mois"]
        ev = evenements_du_mois(charger_evenements(), dernier, code)
        if ev:
            return (f"⚑ {ev[0]['type']}", PRIMAIRE_PALE, PRIMAIRE_FONCE, True)
        vig = {"Normal": (VERT, VERT_TEXTE), "À surveiller": (JAUNE, JAUNE_TEXTE),
               "Inhabituel": (ROUGE, ROUGE_TEXTE)}.get(l["vigilance"], (CARTE, TRES_DOUX))
        texte = f"{l['vigilance']} ({l['ecart'] * 100:+.0f} %)".replace("-", "−") if l["ecart"] is not None else "–"
        return (texte, vig[0], vig[1], True)

    def badge_fiabilite(l):
        return {"Oui": ("Fiable", VERT, VERT_TEXTE), "Gain marginal": ("Gain marginal", JAUNE, JAUNE_TEXTE)}.get(
            l["verdict"], ("Non fiable", ROUGE, ROUGE_TEXTE))

    def afficher_prevision():
        if "prevision" not in donnees:
            return
        p = donnees["prevision"]
        lignes = []
        for code in donnees["codes"]:
            l = p["lignes"].get(code)
            if not l:                                          # un médicament apparu dans un nouvel export
                if fiabilite.get() == "Toutes" and not puce_alertes.actif and \
                        choix_cat_prev.get() in ("", "Toutes", cat.categorie(code)):
                    lignes.append([(nom_court(code), CARTE, TEXTE, True), ("Nouveau", JAUNE, JAUNE_TEXTE, True),
                                   ("pas encore", CARTE, TRES_DOUX, False), ("chaque semaine", CARTE, DOUX, False),
                                   ("–", CARTE, TRES_DOUX, False), ("Pas d'historique", JAUNE, JAUNE_TEXTE, True),
                                   ("–", CARTE, TRES_DOUX, False)])
                continue
            if choix_cat_prev.get() not in ("", "Toutes") and cat.categorie(code) != choix_cat_prev.get():
                continue
            fiable = l["verdict"] == "Oui"
            if (fiabilite.get() == "Fiables" and not fiable) or (fiabilite.get() == "Non fiables" and fiable):
                continue
            if puce_alertes.actif and l["vigilance"] != "Inhabituel":
                continue
            prev = l.get("prevision")
            if isinstance(prev, (int, float)):
                err = l["notre"] or 0
                case_prev = (nombre(prev), CARTE, TEXTE, True)
                fourchette = (f"{nombre(prev * (1 - err))} à {nombre(prev * (1 + err))}", CARTE, DOUX, False)
                case_a_prevoir = (nombre(a_prevoir(l)), PRIMAIRE_PALE, PRIMAIRE_FONCE, True)
            else:
                case_prev = ("Pas de chiffre", CARTE, TRES_DOUX, False)
                fourchette = ("voir le calendrier", CARTE, TRES_DOUX, False)
                case_a_prevoir = ("chaque semaine", CARTE, DOUX, False)
            badge = badge_fiabilite(l)
            lignes.append([(nom_court(code), CARTE, TEXTE, True), case_prev, fourchette, case_a_prevoir,
                           (nombre(l.get("an_dernier")), CARTE, DOUX, False), (badge[0], badge[1], badge[2], True),
                           case_ecart(code, l)])
        cible, dernier = p["mois_cible"], p["dernier_mois"]
        entetes = ["Médicament", f"Prévision {MOIS_COURT[cible.month - 1]} {cible.year % 100:02d}", "Fourchette",
                   f"À prévoir ({prudence.get().lower()})", f"Réel {MOIS_COURT[cible.month - 1]} {(cible.year - 1) % 100:02d}",
                   "Fiabilité", f"Écart {MOIS_COURT[dernier.month - 1]} {dernier.year % 100:02d}"]
        if not lignes:
            lignes = [[("Aucun médicament ne correspond aux filtres.", CARTE, DOUX, False)] + [""] * 6]
        grille(grille_prev, entetes, lignes, [14, 12, 13, 16, 10, 11, 15],
               ["w", "e", "e", "e", "e", "center", "center"])

    def contenu_pdf():
        """Rassemble ce qu'il faut pour la synthèse PDF (en graphiques)."""
        p = donnees["prevision"]
        cible, dernier = p["mois_cible"], p["dernier_mois"]
        prevision, sans_chiffre = [], []
        for code in donnees["codes"]:
            l = p["lignes"].get(code)
            if not l:
                continue
            prev = l.get("prevision")
            if isinstance(prev, (int, float)):
                err = l["notre"] or 0
                prevision.append({"nom": nom_court(code), "prevision": prev, "bas": prev * (1 - err),
                                  "haut": prev * (1 + err), "an_dernier": l.get("an_dernier") or 0})
            else:
                sans_chiffre.append(nom_court(code))
        verdicts = [l["verdict"] for l in p["lignes"].values()]
        fiabilite = (verdicts.count("Oui"), verdicts.count("Gain marginal"),
                     len(verdicts) - verdicts.count("Oui") - verdicts.count("Gain marginal"))
        mois = mois_complets(donnees["ventes"])
        an_avant = dernier - pd.DateOffset(years=1)
        evolution = []
        for code in donnees["codes"]:
            if an_avant in mois.index and mois.loc[an_avant, code]:
                evolution.append((nom_court(code), mois.loc[dernier, code] / mois.loc[an_avant, code] - 1))
        para = p["lignes"].get("N02BE", {})
        nb = sum(1 for l in p["lignes"].values() if l["vigilance"] == "Inhabituel")
        g = donnees.get("grippe")
        chiffres = [(nombre(para.get("prevision")) if isinstance(para.get("prevision"), (int, float)) else "–",
                     f"paracétamol prévu en {MOIS_FR[cible.month - 1]} (la moitié des ventes)"),
                    (str(nb) if nb else "Aucun", f"médicament au comportement inhabituel en {MOIS_FR[dernier.month - 1]}"),
                    (g["statut"] if g else "–", "grippe (Réseau Sentinelles)")]
        notes = []
        if g:
            notes.append(f"Grippe : {nombre(g['niveau'])} cas pour 100 000 habitants la semaine du {g['semaine']:%d/%m/%Y}"
                         + (f" (habituel à cette période : {nombre(g['habituel'])})" if g["habituel"] else "")
                         + f". {g['statut']}.")
        evs = [e for e in charger_evenements() if e["fin"] >= f"{dernier - pd.DateOffset(months=2):%Y-%m}"]
        notes += [f"⚑ {e['debut']} → {e['fin']} · {e['type']} : {e['texte']}" for e in evs[:4]] or \
                 ["Aucun événement noté sur les trois derniers mois."]
        k = {"Au plus juste": 0, "Prudent": 1, "Très prudent": 2}[prudence.get()]
        note_prev = ("Gris = même mois l'an dernier, bleu = prévision, trait noir = fourchette probable."
                     + (f" Quantité à prévoir (niveau « {prudence.get().lower()} ») : le haut de la fourchette." if k == 1
                        else f" Quantité à prévoir (niveau « {prudence.get().lower()} ») : voir l'interface." if k else ""))
        if sans_chiffre:
            note_prev += f" Pas de chiffre fiable pour {', '.join(sans_chiffre)} : commander chaque semaine."
        return {"mois_cible": cible,
                "sous_titre": f"Prévision pour {MOIS_FR[cible.month - 1]} {cible.year} · données jusqu'au "
                              f"{donnees['dernier_jour']:%d/%m/%Y} · éditée le {datetime.now():%d/%m/%Y}",
                "chiffres": chiffres, "fiabilite": fiabilite,
                "titre_prevision": f"Prévision pour {MOIS_FR[cible.month - 1]} {cible.year}",
                "libelle_court_an_dernier": f"{MOIS_COURT[cible.month - 1]} {cible.year - 1}",
                "prevision": prevision, "note_prevision": note_prev,
                "titre_ventes": f"Ventes de {MOIS_FR[dernier.month - 1]} {dernier.year} : évolution sur un an",
                "evolution": evolution, "notes": notes}

    def action_pdf():
        if "prevision" not in donnees:
            messagebox.showwarning("Synthèse PDF", "Lancez d'abord une mise à jour du classeur.")
            return
        try:
            chemin = exporter_pdf(contenu_pdf())
        except Exception as erreur:
            messagebox.showerror("Synthèse PDF", f"Le PDF n'a pas pu être créé :\n{erreur}")
            return
        ecrire(f"✓ Synthèse PDF : rapports/{chemin.name}")
        ouvrir(chemin)

    def texte_grippe(court=False):
        g = donnees.get("grippe")
        if not g:
            return "Grippe : données du Réseau Sentinelles indisponibles (page Mise à jour, étape 2)."
        habituel = f", habituel à cette période : {nombre(g['habituel'])}" if g["habituel"] else ""
        texte = (f"Grippe (Réseau Sentinelles) : {nombre(g['niveau'])} cas pour 100 000 habitants la semaine du "
                 f"{g['semaine']:%d/%m/%Y}{habituel}. {g['statut']}.")
        if g["statut"] != "Calme":
            texte += " Renforcer paracétamol, asthme et ibuprofène dès cette semaine (+51 % de paracétamol en forte grippe)."
        test = donnees.get("test_grippe") or {}
        if test and not court:
            gagnants = [c for c, (sans, avec) in test.items() if avec < sans - 0.01]
            para = test.get("N02BE")
            if gagnants:
                texte += ("\nTesté sur 21 mois : tenir compte de la grippe améliore la prévision de "
                          + ", ".join(nom_court(c) for c in gagnants) + ".")
            elif para:
                texte += (f"\nTesté sur 21 mois : tenir compte de la grippe du mois précédent ne rend pas la prévision "
                          f"mensuelle plus juste (paracétamol : {nombre(para[0] * 100, 1)} % d'erreur sans, {nombre(para[1] * 100, 1)} % "
                          "avec). La grippe sert donc d'alerte en cours de mois, pas à corriger le chiffre.")
        return texte

    # ================================================================ page : Suivi des prévisions
    page = pages["Suivi"]
    titre_page(page, "Suivi des prévisions",
               "Chaque mois, la prévision est comparée au réel. Jusqu'au début de la mission, les mois sont recalculés "
               "(test) ; ensuite, ce sont les prévisions réellement enregistrées par cette interface.")
    bandeau_suivi = tk.Frame(page, bg=FOND)
    bandeau_suivi.pack(fill="x", padx=24)
    chiffres_suivi = {}
    for cle, libelle in [("mois", "Mois suivis"), ("dans", "Dans la fourchette"), ("gain", "Gain vs prévision simple"),
                         ("enreg", "Prévisions enregistrées")]:
        c = carte(bandeau_suivi)
        c.pack(side="left", fill="both", expand=True, padx=4)
        tk.Label(c, text=libelle.upper(), bg=CARTE, fg=DOUX, font=f(8, True)).pack(anchor="w", padx=16, pady=(12, 0))
        chiffres_suivi[cle] = tk.Label(c, text="–", bg=CARTE, fg=TEXTE, font=f(15, True))
        chiffres_suivi[cle].pack(anchor="w", padx=16, pady=(0, 9))
    filtres_suivi = carte(page)
    filtres_suivi.pack(fill="x", padx=28, pady=(12, 0))
    ls = tk.Frame(filtres_suivi, bg=CARTE)
    ls.pack(fill="x", padx=18, pady=12)
    tk.Label(ls, text="PÉRIODE", bg=CARTE, fg=DOUX, font=f(9, True), width=10, anchor="w").pack(side="left")
    periode_suivi = tk.StringVar(value="12 derniers mois")
    segments(ls, ["6 derniers mois", "12 derniers mois", "Tout le suivi"], periode_suivi,
             lambda: afficher_suivi()).pack(side="left", padx=(0, 16))
    puce_sans_ev = Puce(ls, "Exclure les mois avec un événement noté", commande=lambda p: afficher_suivi(), coche=True)
    puce_sans_ev.pack(side="left")
    res_suivi = carte(page)
    res_suivi.pack(fill="both", expand=True, padx=28, pady=(12, 24))
    boite_suivi, grille_suivi = defilable(res_suivi)
    boite_suivi.pack(fill="both", expand=True, padx=18, pady=(16, 8))
    leg_suivi = tk.Frame(res_suivi, bg=CARTE)
    leg_suivi.pack(fill="x", padx=18, pady=(4, 14))
    legende(leg_suivi, [(VERT, "Dans l'erreur habituelle"), (JAUNE, "Jusqu'à 2 fois l'erreur habituelle"),
                        (ROUGE, "Plus de 2 fois"), (FOND, "⚑ événement noté ce mois-là")],
            "Erreur du mois :").pack(anchor="w")
    tk.Label(leg_suivi, text="Chaque case : écart entre la prévision et le réel, en % du réel. « Prévision simple » : "
                             "la meilleure des deux prévisions naïves (le mois dernier, ou le même mois l'an dernier). "
                             "Le gain montre ce que la méthode apporte au client, mois après mois.",
             bg=CARTE, fg=DOUX, font=f(9), justify="left", wraplength=980).pack(anchor="w", pady=(8, 0))

    def afficher_suivi():
        t = donnees.get("suivi")
        if t is None or t.empty:
            grille(grille_suivi, ["Suivi indisponible : lancez une mise à jour."], [], [40])
            return
        tous_mois = sorted(t["mois"].unique())
        n = {"6 derniers mois": 6, "12 derniers mois": 12}.get(periode_suivi.get(), len(tous_mois))
        mois_vus = tous_mois[-n:]
        t = t[t["mois"].isin(mois_vus)]
        compte = t[~t["evenement"]] if puce_sans_ev.actif else t
        lignes, dans, total = [], 0, 0
        for code in donnees["codes"]:
            tc = t[t["code"] == code].set_index("mois")
            if tc.empty:
                continue
            ligne = [(nom_court(code), CARTE, TEXTE, True)]
            for m in mois_vus:
                if m not in tc.index:
                    ligne.append(("–", CARTE, TRES_DOUX, False))
                    continue
                r = tc.loc[m]
                ratio = r["erreur"] / r["habituelle"] if r["habituelle"] else 0
                fond, encre = (VERT, VERT_TEXTE) if ratio <= 1 else (JAUNE, JAUNE_TEXTE) if ratio <= 2 else (ROUGE, ROUGE_TEXTE)
                ligne.append((f"{r['erreur'] * 100:.0f} %" + (" ⚑" if r["evenement"] else "")
                              + ("*" if r["source"] == "enregistrée" else ""), fond, encre, False))
            cc = compte[compte["code"] == code]
            nous, naif = cc["erreur"].mean(), cc["erreur_naive"].mean()
            gain = (naif - nous) * 100
            ligne += [(f"{nous * 100:.1f} %", FOND, TEXTE, True), (f"{naif * 100:.1f} %", FOND, DOUX, False),
                      (f"{gain:+.1f} pt".replace("-", "−"), VERT if gain >= 2 else CARTE if gain >= 0 else ROUGE,
                       VERT_TEXTE if gain >= 2 else DOUX if gain >= 0 else ROUGE_TEXTE, True)]
            lignes.append(ligne)
            dans += int((cc["erreur"] <= cc["habituelle"]).sum())
            total += len(cc)
        entetes = ["Médicament"] + [f"{MOIS_COURT[m.month - 1]} {m.year % 100:02d}" for m in mois_vus] + \
                  ["Notre erreur", "Prév. simple", "Gain"]
        grille(grille_suivi, entetes, lignes, [14] + [7] * len(mois_vus) + [11, 11, 9])
        chiffres_suivi["mois"].configure(text=f"{len(mois_vus)} mois")
        chiffres_suivi["dans"].configure(text=f"{dans / total * 100:.0f} %" if total else "–")
        gain_moyen = (compte["erreur_naive"].mean() - compte["erreur"].mean()) * 100
        chiffres_suivi["gain"].configure(text=("+" if gain_moyen > 0 else "") + nombre(gain_moyen, 1).replace("-", "−")
                                              + " points",
                                         fg=VERT_TEXTE if gain_moyen > 0 else ROUGE_TEXTE)
        nb_enreg = len(pd.read_csv(FICHIER_SUIVI)) if FICHIER_SUIVI.exists() else 0
        chiffres_suivi["enreg"].configure(text=str(nb_enreg))

    # ================================================================ page : Événements
    page = pages["Événements"]
    titre_page(page, "Événements",
               "Notez ce qui sort de l'ordinaire : rupture de stock, fermeture, épidémie, crise sanitaire, promotion… "
               "Ces mois sont signalés ⚑ dans les pages Ventes, Prévision et Suivi, pour ne pas les confondre avec une vraie tendance.")
    corps_ev = tk.Frame(page, bg=FOND)
    corps_ev.pack(fill="both", expand=True, padx=24, pady=(0, 24))
    droite_ev = carte(corps_ev)
    droite_ev.pack(side="right", fill="y", padx=4)
    gauche_ev = carte(corps_ev)
    gauche_ev.pack(side="left", fill="both", expand=True, padx=4)
    tk.Label(gauche_ev, text="Événements notés", bg=CARTE, fg=TEXTE, font=f(12, True)).pack(anchor="w", padx=14,
                                                                                          pady=(14, 8))
    cols_ev = ("debut", "fin", "type", "qui", "texte")
    tableau_ev = ttk.Treeview(gauche_ev, columns=cols_ev, show="headings", height=12, selectmode="browse")
    for col, titre, larg in zip(cols_ev, ["Du", "Au", "Type", "Médicaments", "Description"], [70, 70, 170, 150, 230]):
        tableau_ev.heading(col, text=titre, anchor="w")
        tableau_ev.column(col, width=larg, anchor="w")
    tableau_ev.pack(fill="both", expand=True, padx=14, pady=(0, 14))
    vide_ev = tk.Label(gauche_ev, text="", bg=CARTE, fg=DOUX, font=f(9))
    vide_ev.pack(anchor="w", padx=14, pady=(0, 14))

    form_ev = tk.Frame(droite_ev, bg=CARTE)
    form_ev.pack(fill="both", padx=20, pady=18)
    tk.Label(form_ev, text="Noter un événement", bg=CARTE, fg=TEXTE, font=f(12, True)).pack(anchor="w")
    tk.Label(form_ev, text="Cliquez sur une ligne pour la modifier.", bg=CARTE, fg=DOUX, font=f(9)).pack(
        anchor="w", pady=(2, 8))
    tk.Label(form_ev, text="Type", bg=CARTE, fg=TEXTE, font=f(9, True)).pack(anchor="w", pady=(6, 2))
    tk.Label(form_ev, text="Choisissez dans la liste, ou écrivez votre propre type.", bg=CARTE, fg=TRES_DOUX,
             font=f(8)).pack(anchor="w")
    ev_type = ttk.Combobox(form_ev, values=TYPES_EVENEMENTS, width=30, font=f(10))
    ev_type.pack(anchor="w")
    l_dates_ev = tk.Frame(form_ev, bg=CARTE)
    l_dates_ev.pack(anchor="w", pady=(10, 0))
    tk.Label(l_dates_ev, text="Du", bg=CARTE, fg=TEXTE, font=f(9, True)).grid(row=0, column=0, sticky="w")
    tk.Label(l_dates_ev, text="Au", bg=CARTE, fg=TEXTE, font=f(9, True)).grid(row=0, column=1, sticky="w", padx=(10, 0))
    ev_debut = ttk.Combobox(l_dates_ev, state="readonly", width=14, font=f(10))
    ev_fin = ttk.Combobox(l_dates_ev, state="readonly", width=14, font=f(10))
    ev_debut.grid(row=1, column=0)
    ev_fin.grid(row=1, column=1, padx=(10, 0))
    ev_debut.bind("<<ComboboxSelected>>", lambda _: ev_fin.get() < ev_debut.get() and ev_fin.set(ev_debut.get()))
    tk.Label(form_ev, text="Médicaments concernés", bg=CARTE, fg=TEXTE, font=f(9, True)).pack(anchor="w", pady=(10, 2))
    tk.Label(form_ev, text="Aucun sélectionné = tous les médicaments.", bg=CARTE, fg=TRES_DOUX, font=f(8)).pack(anchor="w")
    cadre_ev_codes = tk.Frame(form_ev, bg=CARTE)
    cadre_ev_codes.pack(anchor="w", pady=(4, 0))
    puces_ev = {}
    tk.Label(form_ev, text="Description", bg=CARTE, fg=TEXTE, font=f(9, True)).pack(anchor="w", pady=(10, 2))
    ev_texte = tk.Entry(form_ev, width=34, font=f(10), relief="flat", highlightthickness=1, highlightbackground=BORD,
                        highlightcolor=PRIMAIRE)
    ev_texte.pack(anchor="w", ipady=5)
    message_ev = tk.Label(form_ev, text="", bg=CARTE, fg=DOUX, font=f(9), justify="left", wraplength=280)
    message_ev.pack(anchor="w", pady=(8, 0))
    boutons_ev = tk.Frame(form_ev, bg=CARTE)
    boutons_ev.pack(anchor="w", pady=(10, 0))
    selection_ev = {"index": None}

    def vider_ev():
        ev_type.set("")
        ev_texte.delete(0, "end")
        for pc in puces_ev.values():
            pc.basculer(False, appeler=False)
        dernier = f"{donnees['dernier_mois']:%Y-%m}" if donnees else ""
        ev_debut.set(dernier)
        ev_fin.set(dernier)
        selection_ev["index"] = None
        tableau_ev.selection_remove(tableau_ev.selection())

    def choisir_ev(_=None):
        choix = tableau_ev.selection()
        if not choix:
            return
        i = int(choix[0])
        e = charger_evenements()[i]
        vider_ev()
        selection_ev["index"] = i
        ev_type.set(e["type"])
        ev_debut.set(e["debut"])
        ev_fin.set(e["fin"])
        ev_texte.insert(0, e["texte"])
        for c, pc in puces_ev.items():
            pc.basculer(c in e["codes"], appeler=False)
        message_ev.configure(text="")
    tableau_ev.bind("<<TreeviewSelect>>", choisir_ev)

    def enregistrer_ev():
        if not ev_type.get().strip() or not ev_debut.get() or not ev_fin.get():
            message_ev.configure(text="✗ Choisissez un type et les mois.", fg=ROUGE_TEXTE)
            return
        e = {"debut": min(ev_debut.get(), ev_fin.get()), "fin": max(ev_debut.get(), ev_fin.get()),
             "type": ev_type.get().strip(), "codes": [c for c, pc in puces_ev.items() if pc.actif],
             "texte": ev_texte.get().strip()}
        evenements = charger_evenements()
        if selection_ev["index"] is None:
            evenements.append(e)
        else:
            evenements[selection_ev["index"]] = e
        enregistrer_evenements(sorted(evenements, key=lambda x: x["debut"]))
        message_ev.configure(text="✓ Événement enregistré.", fg=VERT_TEXTE)
        vider_ev()
        remplir_evenements()

    def supprimer_ev():
        if selection_ev["index"] is None:
            message_ev.configure(text="Cliquez d'abord sur l'événement à supprimer.", fg=DOUX)
            return
        if not messagebox.askyesno("Supprimer", "Supprimer cet événement ?"):
            return
        evenements = charger_evenements()
        del evenements[selection_ev["index"]]
        enregistrer_evenements(evenements)
        message_ev.configure(text="✓ Événement supprimé.", fg=VERT_TEXTE)
        vider_ev()
        remplir_evenements()

    bouton(boutons_ev, "Enregistrer", enregistrer_ev, petit=True).pack(side="left")
    bouton(boutons_ev, "Supprimer", supprimer_ev, sorte="danger", petit=True).pack(side="left", padx=6)
    bouton(boutons_ev, "Vider", vider_ev, sorte="secondaire", petit=True).pack(side="left")

    def remplir_evenements():
        evenements = charger_evenements()
        tableau_ev.delete(*tableau_ev.get_children())
        for i, e in enumerate(evenements):
            qui = ", ".join(nom_court(c) for c in e["codes"]) if e["codes"] else "Tous"
            tableau_ev.insert("", "end", iid=str(i), values=(e["debut"], e["fin"], e["type"], qui, e["texte"]))
        vide_ev.configure(text="" if evenements else "Aucun événement noté pour l'instant.")
        types_perso = sorted({e["type"] for e in evenements} - set(TYPES_EVENEMENTS))   # les types écrits à la main
        ev_type.configure(values=TYPES_EVENEMENTS + types_perso)
        # les autres pages tiennent compte des événements
        if "suivi" in donnees and "prevision" in donnees:
            donnees["suivi"] = suivi_test(mois_complets(donnees["ventes"]), donnees["prevision"]["lignes"], evenements)
            afficher_suivi()
        afficher_prevision()
        afficher_ventes()

    # ================================================================ page 4 : Catégories
    page = pages["Catégories"]
    titre_page(page, "Catégories de médicaments",
               "Chaque médicament est rangé grâce à son code ATC. Vous pouvez en ajouter ou en reclasser à la main : "
               "vos choix s'appliquent aussi au classeur, à la prochaine mise à jour.")
    bas_cat = carte(page)
    bas_cat.pack(side="bottom", fill="x", padx=28, pady=(8, 24))
    corps_cat = tk.Frame(page, bg=FOND)
    corps_cat.pack(fill="both", expand=True, padx=24)
    droite = carte(corps_cat)
    droite.pack(side="right", fill="y", padx=4)
    gauche = carte(corps_cat)
    gauche.pack(side="left", fill="both", expand=True, padx=4)
    alerte = tk.Label(gauche, text="", font=f(10, True), anchor="w", padx=14, pady=10, justify="left",
                      wraplength=560)
    alerte.pack(fill="x", padx=14, pady=(14, 8))
    cols_cat = ("code", "nom", "cat", "origine", "export")
    tableau_cat = ttk.Treeview(gauche, columns=cols_cat, show="headings", height=8, selectmode="browse")
    for col, titre, larg in zip(cols_cat, ["Code", "Nom", "Catégorie", "Origine", "Dans l'export"],
                                [55, 125, 175, 125, 85]):
        tableau_cat.heading(col, text=titre, anchor="w")
        tableau_cat.column(col, width=larg, anchor="w")
    tableau_cat.tag_configure("main", background=PRIMAIRE_PALE)
    tableau_cat.tag_configure("attente", foreground=DOUX)
    tableau_cat.pack(fill="both", expand=True, padx=14, pady=(0, 14))
    tk.Label(bas_cat, text="Ventes par catégorie : les 12 derniers mois complets contre les 12 d'avant", bg=CARTE,
             fg=TEXTE, font=f(11, True)).pack(anchor="w", padx=14, pady=(12, 6))
    grille_cat = tk.Frame(bas_cat, bg=BORD)
    grille_cat.pack(anchor="w", padx=14)
    tk.Label(bas_cat, text="Les quantités de médicaments différents sont additionnées : "
                           "à suivre dans le temps, pas à comparer d'une catégorie à l'autre.",
             bg=CARTE, fg=DOUX, font=f(9), justify="left").pack(anchor="w", padx=14, pady=(6, 12))

    form = tk.Frame(droite, bg=CARTE)
    form.pack(fill="both", padx=20, pady=18)
    tk.Label(form, text="Ajouter ou modifier", bg=CARTE, fg=TEXTE, font=f(12, True)).pack(anchor="w")
    tk.Label(form, text="Cliquez sur une ligne du tableau pour la modifier,\nou remplissez les cases pour en ajouter une.",
             bg=CARTE, fg=DOUX, font=f(9), justify="left").pack(anchor="w", pady=(2, 12))

    def champ(titre, aide):
        tk.Label(form, text=titre, bg=CARTE, fg=TEXTE, font=f(9, True)).pack(anchor="w", pady=(8, 2))
        return aide

    champ("Code ATC", None)
    saisie_code = tk.Entry(form, width=24, font=f(10), relief="flat", highlightthickness=1, highlightbackground=BORD,
                           highlightcolor=PRIMAIRE)
    saisie_code.pack(anchor="w", ipady=5)
    tk.Label(form, text="Exemples : R05 (toux), N02BB, A02 (estomac)", bg=CARTE, fg=TRES_DOUX, font=f(8)).pack(
        anchor="w")
    champ("Nom", None)
    saisie_nom = tk.Entry(form, width=24, font=f(10), relief="flat", highlightthickness=1, highlightbackground=BORD,
                          highlightcolor=PRIMAIRE)
    saisie_nom.pack(anchor="w", ipady=5)
    champ("Catégorie", None)
    saisie_cat = ttk.Combobox(form, width=22, font=f(10))
    saisie_cat.pack(anchor="w")
    tk.Label(form, text="Choisissez dans la liste,\nou tapez une nouvelle catégorie.", bg=CARTE, fg=TRES_DOUX, justify="left",
             font=f(8)).pack(anchor="w")
    message_cat = tk.Label(form, text="", bg=CARTE, fg=DOUX, font=f(9), justify="left", wraplength=230)
    message_cat.pack(anchor="w", pady=(10, 0))
    boutons_cat = tk.Frame(form, bg=CARTE)
    boutons_cat.pack(anchor="w", pady=(12, 0))

    def vider():
        for s in (saisie_code, saisie_nom):
            s.delete(0, "end")
        saisie_cat.set("")
        tableau_cat.selection_remove(tableau_cat.selection())

    def selection_cat(_=None):
        choix = tableau_cat.selection()
        if not choix:
            return
        code, nom, categorie, *_ = tableau_cat.item(choix[0], "values")
        vider()
        saisie_code.insert(0, code)
        saisie_nom.insert(0, nom)
        saisie_cat.set(categorie)
        message_cat.configure(text="")
    tableau_cat.bind("<<TreeviewSelect>>", selection_cat)

    def enregistrer():
        code = saisie_code.get().strip().upper()
        nom, categorie = saisie_nom.get().strip(), saisie_cat.get().strip()
        if not cat.code_valide(code):
            message_cat.configure(text="✗ Ce code n'a pas la forme d'un code ATC (une lettre, deux chiffres, "
                                       "puis éventuellement des lettres : R05, N02BB…).", fg=ROUGE_TEXTE)
            return
        if not categorie:
            message_cat.configure(text="✗ Choisissez ou tapez une catégorie.", fg=ROUGE_TEXTE)
            return
        perso = cat.charger_perso()
        perso[code] = {"nom": nom or nom_court(code), "categorie": categorie}
        cat.enregistrer_perso(perso)
        message_cat.configure(text=f"✓ {code} enregistré dans « {categorie} ». "
                                   "Le classeur en tiendra compte à la prochaine mise à jour.", fg=VERT_TEXTE)
        remplir_categories()

    def supprimer():
        code = saisie_code.get().strip().upper()
        perso = cat.charger_perso()
        if code not in perso:
            message_cat.configure(text="Ce médicament n'a pas été ajouté à la main : rien à supprimer.", fg=DOUX)
            return
        if not messagebox.askyesno("Supprimer", f"Retirer l'ajout fait à la main pour {code} ?\n"
                                                "S'il est dans l'export, il reprendra son classement automatique."):
            return
        del perso[code]
        cat.enregistrer_perso(perso)
        message_cat.configure(text=f"✓ Ajout retiré pour {code}.", fg=VERT_TEXTE)
        vider()
        remplir_categories()

    bouton(boutons_cat, "Enregistrer", enregistrer, petit=True).pack(side="left")
    bouton(boutons_cat, "Supprimer l'ajout", supprimer, sorte="danger", petit=True).pack(side="left", padx=6)
    bouton(boutons_cat, "Vider", vider, sorte="secondaire", petit=True).pack(side="left")

    def remplir_categories():
        perso = cat.charger_perso()
        codes = list(donnees["codes"]) + [c for c in sorted(perso) if c not in donnees["codes"]]
        tableau_cat.delete(*tableau_cat.get_children())
        for c in codes:
            dans_export = c in donnees["codes"]
            if c in perso:
                origine = "Reclassé à la main" if dans_export and c in cat.GROUPES_CONNUS else "Ajouté à la main"
            else:
                origine = "Classification ATC"
            tags = (("main",) if c in perso else ()) + (() if dans_export else ("attente",))
            tableau_cat.insert("", "end", values=(c, nom_court(c) if c in NOMS or c in perso else cat.sous_groupe(c),
                                                  cat.categorie(c), origine,
                                                  "Oui" if dans_export else "Pas encore"), tags=tags)
        saisie_cat.configure(values=cat.toutes_les_categories())
        nouveaux = donnees["nouveaux"]
        if nouveaux:
            alerte.configure(text=f"★  Nouveaux médicaments dans l'export : {', '.join(nouveaux)}. "
                                  "Vérifiez leur catégorie ci-dessous.", bg=JAUNE, fg=JAUNE_TEXTE)
        else:
            alerte.configure(text="✓  Aucun nouveau médicament dans l'export : les mêmes 8 groupes qu'au départ.",
                             bg=VERT, fg=VERT_TEXTE)
        # ventes par catégorie : 12 derniers mois complets contre les 12 précédents
        mois = mois_complets(donnees["ventes"])
        par_cat = mois.T.groupby(cat.categorie).sum().T
        lignes = []
        for nom in par_cat.columns:
            recent, avant = par_cat[nom].iloc[-12:].sum(), par_cat[nom].iloc[-24:-12].sum()
            evol = recent / avant - 1 if avant else None
            if evol is None:
                case = ("nouveau", CARTE, DOUX, False)
            else:
                fond, encre = (VERT, VERT_TEXTE) if evol > 0.05 else (ROUGE, ROUGE_TEXTE) if evol < -0.05 \
                    else (CARTE, DOUX)
                case = (f"{'▲' if evol > 0 else '▼'} {evol * 100:+.0f} %".replace("-", "−"), fond, encre, True)
            membres = ", ".join(nom_court(c) for c in mois.columns if cat.categorie(c) == nom)
            lignes.append([(nom, CARTE, TEXTE, True), (membres, CARTE, DOUX, False), nombre(recent), nombre(avant),
                           case])
        grille(grille_cat, ["Catégorie", "Médicaments", "Sur 12 mois", "12 mois avant", "Évolution"], lignes,
               [22, 40, 12, 12, 10], ["w", "w", "e", "e", "center"])
        # les listes des autres pages suivent les catégories
        choix_cat_prev.configure(values=["Toutes"] + sorted({cat.categorie(c) for c in donnees["codes"]}))
        if not choix_cat_prev.get():
            choix_cat_prev.set("Toutes")
        afficher_prevision()

    # ================================================================ chargement des données
    def rafraichir():
        try:
            ventes, codes, horaire = charger_ventes()
        except Exception as erreur:
            pied_menu.configure(text=f"Données illisibles :\n{erreur}")
            return
        donnees.update({"ventes": ventes, "codes": codes, "horaire": horaire,
                        "premier_jour": ventes["jour"].min(), "dernier_jour": ventes["jour"].max(),
                        "nouveaux": cat.nouveaux_codes(codes)})
        mois = mois_complets(ventes)
        donnees["dernier_mois"] = mois.index.max()
        try:
            donnees["prevision"] = charger_prevision(ventes)
        except Exception as erreur:
            donnees.pop("prevision", None)
            titre_prev.configure(text=f"Prévision indisponible : {erreur}. Lancez une mise à jour.")
        try:                                                   # la grippe, si le fichier est là
            serie = lire_grippe()
            donnees["grippe"] = etat_grippe(serie, donnees["dernier_jour"])
            donnees["test_grippe"] = test_grippe(mois, serie, donnees["prevision"]["lignes"]) \
                if "prevision" in donnees else {}
        except Exception:
            donnees["grippe"], donnees["test_grippe"] = None, {}
        if "prevision" in donnees:                             # le suivi, et la trace de la prévision du mois
            try:
                enregistrer_prevision(donnees["prevision"])
            except Exception:
                pass
            donnees["suivi"] = suivi_test(mois, donnees["prevision"]["lignes"], charger_evenements())
        maj = datetime.fromtimestamp(CLASSEUR.stat().st_mtime) if CLASSEUR.exists() else None
        pied_menu.configure(text=f"Données jusqu'au\n{donnees['dernier_jour']:%d/%m/%Y}\n\n"
                                 + (f"Classeur mis à jour le\n{maj:%d/%m/%Y à %H:%M}" if maj else ""))

        # page Ventes : médicaments et liste des mois
        for w in cadre_codes.winfo_children():
            w.destroy()
        puces_codes.clear()
        rangee = None
        for i, c in enumerate(codes):                         # 6 par ligne : avec de nouveaux médicaments, ça passe à la ligne
            if i % 6 == 0:
                rangee = tk.Frame(cadre_codes, bg=CARTE)
                rangee.pack(anchor="w", pady=(0, 4))
            puces_codes[c] = Puce(rangee, nom_court(c), actif=True, commande=code_bascule, coche=True)
            puces_codes[c].pack(side="left", padx=2)
        for w in l_codes2.winfo_children()[1:]:
            w.destroy()
        bouton(l_codes2, "✓  Tout sélectionner", lambda: tous_les_codes(True), sorte="secondaire", petit=True).pack(
            side="left")
        bouton(l_codes2, "○  Tout désélectionner", lambda: tous_les_codes(False), sorte="secondaire",
               petit=True).pack(side="left", padx=6)
        tk.Label(l_codes2, text="Cliquez sur un médicament pour l'afficher ou le masquer.", bg=CARTE, fg=TRES_DOUX,
                 font=f(8)).pack(side="left", padx=8)
        etat["codes"] = set(codes)
        mois_dispo = pd.period_range(donnees["premier_jour"], donnees["dernier_jour"], freq="M")[::-1]
        choix_mois.configure(values=[f"{MOIS_FR[p.month - 1]} {p.year}" for p in mois_dispo])
        maj_menu_ajout()
        periode_raccourci("12 derniers mois")

        # page Prévision : chiffres clés
        if "prevision" in donnees:
            p = donnees["prevision"]
            cible = p["mois_cible"]
            titre_prev.configure(text=f"Prévision pour {MOIS_FR[cible.month - 1]} {cible.year}")
            chiffres["cible"].configure(text=f"{MOIS_FR[cible.month - 1].capitalize()} {cible.year}")
            g = donnees.get("grippe")
            chiffres["grippe"].configure(text=g["statut"] if g else "–",
                                         fg=(ROUGE_TEXTE if g["statut"] == "Épidémie en cours" else JAUNE_TEXTE
                                             if g["statut"] == "En hausse" else VERT_TEXTE) if g else DOUX)
            detail_grippe.configure(text=f"{nombre(g['niveau'])} cas / 100 000 · habituel {nombre(g['habituel'])}"
                                    if g and g["habituel"] else "")
            nb_fiables = sum(1 for l in p["lignes"].values() if l["verdict"] == "Oui")
            chiffres["fiables"].configure(text=f"{nb_fiables} sur {len(p['lignes'])}")
            nb_alertes = sum(1 for l in p["lignes"].values() if l["vigilance"] == "Inhabituel")
            chiffres["alertes"].configure(text=f"{nb_alertes} médicament{'s' if nb_alertes > 1 else ''}" if nb_alertes
                                          else "Aucun",
                                          fg=ROUGE_TEXTE if nb_alertes else VERT_TEXTE)
        remplir_categories()

        # page Événements : listes des mois et des médicaments
        mois_ev = [p.strftime("%Y-%m") for p in pd.period_range(donnees["premier_jour"],
                                                          donnees["dernier_jour"] + pd.DateOffset(months=3), freq="M")]
        ev_debut.configure(values=mois_ev[::-1])
        ev_fin.configure(values=mois_ev[::-1])
        for w in cadre_ev_codes.winfo_children():
            w.destroy()
        puces_ev.clear()
        for i, c in enumerate(codes):
            puces_ev[c] = Puce(cadre_ev_codes, nom_court(c), coche=True)
            puces_ev[c].grid(row=i // 2, column=i % 2, sticky="w", padx=(0, 4), pady=2)
        vider_ev()
        remplir_evenements()

    ecrire("Prêt. Pour un nouvel export : étape 1, puis étape 2.")
    rafraichir()
    afficher("Prévision")
    racine.mainloop()


if __name__ == "__main__":
    lancer_fenetre()
