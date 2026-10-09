import pandas as pd

pd.set_option("display.max_columns", None)
pd.set_option("display.width", 200)

GROUPES = ["M01AB", "M01AE", "N02BA", "N02BE", "N05B", "N05C", "R03", "R06"]

# 1. Lire le fichier de référence (journalier) et le fichier horaire
jour = pd.read_csv("data/Pharma_Ventes_Daily.csv")
jour["datum"] = pd.to_datetime(jour["datum"], format="%m/%d/%Y")
jour = jour.set_index("datum")[GROUPES]

horaire = pd.read_csv("data/Pharma_Ventes_Hourly.csv")
horaire["datum"] = pd.to_datetime(horaire["datum"], format="%m/%d/%Y %H:%M")

# 2. Ventes par mois (en additionnant les jours) et on retire octobre 2019, incomplet
mois = jour.resample("ME").sum()
jours_par_mois = jour.resample("ME").size()
mois = mois[jours_par_mois == jours_par_mois.index.days_in_month]
# 3. Tableau de synthèse par groupe
synthese = pd.DataFrame({
    "Total ventes": mois.sum(),
    "Part du total (%)": mois.sum() / mois.sum().sum() * 100,
    "Moyenne par mois": mois.mean(),
    "Écart-type mensuel": mois.std(),
    "Mois le plus bas": mois.min(),
    "Mois le plus haut": mois.max(),
})
# Coefficient de variation : écart-type / moyenne. Plus il est haut, plus les ventes sont irrégulières
synthese["Irrégularité (%)"] = synthese["Écart-type mensuel"] / synthese["Moyenne par mois"] * 100
synthese = synthese.round(1).sort_values("Total ventes", ascending=False)
print("=== Synthèse par groupe (févr. 2014 - sept. 2019) ===")
print(synthese)

# 4. Évolution d'une année sur l'autre (années complètes seulement)
annees = jour.resample("YE").sum()
annees.index = annees.index.year
annees = annees.loc[2014:2018]
evolution = ((annees.loc[2018] / annees.loc[2014] - 1) * 100).round(1)
print("\n=== Ventes par année ===")
print(annees.round(0))
print("\n=== Évolution 2014 -> 2018 (%) ===")
print(evolution.to_string())

# 5. Ventes moyennes par jour de la semaine
noms_jours = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]
par_jour = jour.groupby(jour.index.dayofweek).mean()
par_jour.index = noms_jours
print("\n=== Ventes moyennes par jour de la semaine ===")
print(par_jour.round(1))

# 6. Ventes moyennes par heure d'ouverture
par_heure = horaire.groupby(horaire["datum"].dt.hour)[GROUPES].mean()
par_heure.index.name = "Heure"
print("\n=== Ventes moyennes par heure ===")
print(par_heure.round(2))

# 7. Tout enregistrer dans un fichier Excel
with pd.ExcelWriter("statistiques.xlsx") as fichier:
    synthese.to_excel(fichier, sheet_name="Synthèse groupes")
    annees.to_excel(fichier, sheet_name="Ventes par année")
    mois.to_excel(fichier, sheet_name="Ventes par mois")
    par_jour.to_excel(fichier, sheet_name="Par jour semaine")
    par_heure.to_excel(fichier, sheet_name="Par heure")
print("\nFichier statistiques.xlsx créé.")