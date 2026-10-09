import pandas as pd

pd.set_option("display.max_columns", None)
pd.set_option("display.width", 200)

GROUPES = ["M01AB", "M01AE", "N02BA", "N02BE", "N05B", "N05C", "R03", "R06"]

# 1. Lire les 4 fichiers, chacun avec son propre format de date
horaire = pd.read_csv("data/Pharma_Ventes_Hourly.csv")
horaire["datum"] = pd.to_datetime(horaire["datum"], format="%m/%d/%Y %H:%M")

journalier = pd.read_csv("data/Pharma_Ventes_Daily.csv")
journalier["datum"] = pd.to_datetime(journalier["datum"], format="%m/%d/%Y")

hebdo = pd.read_csv("data/Pharma_Ventes_Weekly.csv")
hebdo["datum"] = pd.to_datetime(hebdo["datum"], format="%m/%d/%Y")

mensuel = pd.read_csv("data/Pharma_Ventes_Monthly.csv")
mensuel["datum"] = pd.to_datetime(mensuel["datum"], format="%Y-%m-%d")

# 2. Fiche d'identité de chaque fichier
for nom, df in [("Horaire", horaire), ("Journalier", journalier),
                ("Hebdo", hebdo), ("Mensuel", mensuel)]:
    print(f"--- {nom} ---")
    print("Du", df["datum"].min().date(), "au", df["datum"].max().date(),
          "|", len(df), "lignes",
          "| doublons :", df["datum"].duplicated().sum(),
          "| valeurs vides :", df[GROUPES].isna().sum().sum(),
          "| valeurs négatives :", (df[GROUPES] < 0).sum().sum())


# 3. Fonction qui compare deux tableaux ligne à ligne
def comparer(nom, reference, a_verifier):
    ecart = (reference - a_verifier).abs()
    lignes_differentes = (ecart > 0.01).any(axis=1)
    print(f"\n=== {nom} ===")
    print("Périodes comparées :", len(ecart))
    print("Périodes avec un écart :", lignes_differentes.sum())
    print("Écart maximum par groupe :")
    print(ecart.max().round(2).to_string())
    if lignes_differentes.sum() > 0:
        print("Exemples de périodes qui ne collent pas :")
        print(ecart[lignes_differentes].head().round(2))


jour = journalier.set_index("datum")[GROUPES]

# 4. Horaire additionné par jour  vs  journalier
horaire_par_jour = horaire.set_index("datum")[GROUPES].resample("D").sum()
communs = jour.index.intersection(horaire_par_jour.index)
comparer("Horaire additionné par jour vs Journalier",
         jour.loc[communs], horaire_par_jour.loc[communs])

# 5. Journalier additionné par semaine (lundi -> dimanche)  vs  hebdo
jour_par_semaine = jour.resample("W-SUN").sum()
semaine = hebdo.set_index("datum")[GROUPES]
communs = semaine.index.intersection(jour_par_semaine.index)
comparer("Journalier additionné par semaine vs Hebdo",
         semaine.loc[communs], jour_par_semaine.loc[communs])

# 6. Journalier additionné par mois  vs  mensuel
jour_par_mois = jour.resample("ME").sum()
mois = mensuel.set_index("datum")[GROUPES]
communs = mois.index.intersection(jour_par_mois.index)
comparer("Journalier additionné par mois vs Mensuel",
         mois.loc[communs], jour_par_mois.loc[communs])