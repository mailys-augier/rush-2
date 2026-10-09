"""Croise les ventes de la pharmacie avec l'épidémie de grippe (Réseau Sentinelles).

Source : Réseau Sentinelles (IPLESP, INSERM, Sorbonne Université), https://www.sentiweb.fr
Fichier : incidence hebdomadaire des syndromes grippaux, France entière (inc-3-PAY-ds2.csv).
La colonne inc100 = nombre de cas pour 100 000 habitants dans la semaine.

Question du propriétaire : quelle part des ventes vient de l'environnement (ici la grippe),
et quelle part vient de la pharmacie elle-même ?
"""

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd

pd.set_option("display.max_columns", None)
pd.set_option("display.width", 200)

GROUPES = ["M01AB", "M01AE", "N02BA", "N02BE", "N05B", "N05C", "R03", "R06"]
NOMS = {
    "M01AB": "Anti-inflammatoires (diclofénac)", "M01AE": "Anti-inflammatoires (ibuprofène)",
    "N02BA": "Aspirine", "N02BE": "Paracétamol", "N05B": "Anxiolytiques",
    "N05C": "Somnifères", "R03": "Asthme", "R06": "Antihistaminiques",
}
SEUIL_FORTE_GRIPPE = 150  # cas pour 100 000 habitants dans la semaine

# 1. Ventes de la pharmacie, additionnées par semaine (semaine ISO, du lundi au dimanche)
jour = pd.read_csv("data/Pharma_Ventes_Daily.csv")
jour["datum"] = pd.to_datetime(jour["datum"], format="%m/%d/%Y")
iso = jour["datum"].dt.isocalendar()
jour["semaine"] = iso["year"] * 100 + iso["week"]
nb_jours = jour.groupby("semaine").size()
ventes = jour.groupby("semaine")[GROUPES].sum()
ventes = ventes[nb_jours == 7]  # on garde seulement les semaines complètes

# 2. Grippe : la première ligne du fichier est une description, on la saute
grippe = pd.read_csv("donnees_externes/inc-3-PAY-ds2.csv", skiprows=1)
grippe = grippe.rename(columns={"week": "semaine", "inc100": "grippe"}).set_index("semaine")["grippe"]

# 3. On met les deux côte à côte, semaine par semaine (période commune aux deux fichiers)
tableau = ventes.join(grippe, how="inner")
print(f"Semaines communes : {len(tableau)} (de {tableau.index.min()} à {tableau.index.max()})")

# 4. Corrélation brute : les ventes montent-elles quand la grippe monte ?
#    (1 = suivent parfaitement la grippe, 0 = aucun lien, -1 = font l'inverse)
brute = tableau[GROUPES].corrwith(tableau["grippe"])

# 5. Corrélation hors saison : on retire l'effet « hiver » pour isoler l'effet de l'épidémie.
#    Pour chaque semaine de l'année, on compare à la moyenne de cette même semaine sur les 4 ans :
#    une grippe plus forte que d'habitude fait-elle vendre plus que d'habitude ?
tableau["num_semaine"] = tableau.index % 100
ecarts = tableau[GROUPES + ["grippe"]] - tableau.groupby("num_semaine")[GROUPES + ["grippe"]].transform("mean")
hors_saison = ecarts[GROUPES].corrwith(ecarts["grippe"])

# 6. Ventes moyennes par semaine : forte grippe contre le reste du temps
forte = tableau["grippe"] >= SEUIL_FORTE_GRIPPE
moy_forte = tableau.loc[forte, GROUPES].mean()
moy_autre = tableau.loc[~forte, GROUPES].mean()

resultat = pd.DataFrame({
    "Groupe": [NOMS[g] for g in GROUPES],
    "Corrélation avec la grippe": brute.round(2).values,
    "Corrélation hors effet saison": hors_saison.round(2).values,
    "Ventes / semaine en forte grippe": moy_forte.round(1).values,
    "Ventes / semaine le reste du temps": moy_autre.round(1).values,
    "Écart (%)": ((moy_forte / moy_autre - 1) * 100).round(0).values,
}, index=GROUPES)
print(f"\n=== Ventes et grippe ({forte.sum()} semaines de forte grippe, seuil {SEUIL_FORTE_GRIPPE} cas / 100 000) ===")
print(resultat.drop(columns="Groupe"))

# 7. Graphique pour le deck : paracétamol et grippe, semaine par semaine
dates = pd.to_datetime(tableau.index.astype(str) + "1", format="%G%V%u")
fig, ax1 = plt.subplots(figsize=(11, 4))
ax1.plot(dates, tableau["N02BE"], color="#2b6cb0", label="Paracétamol vendu (par semaine)")
ax1.set_ylabel("Paracétamol vendu", color="#2b6cb0")
ax2 = ax1.twinx()
ax2.fill_between(dates, tableau["grippe"], color="#e53e3e", alpha=0.25, label="Grippe (cas / 100 000)")
ax2.set_ylabel("Grippe : cas pour 100 000 habitants", color="#e53e3e")
ax1.set_title("Paracétamol vendu et épidémies de grippe, semaine par semaine (2016-2019)")
fig.tight_layout()
fig.savefig("donnees_externes/paracetamol_grippe.png", dpi=150)

# 8. Enregistrer
with pd.ExcelWriter("donnees_externes/ventes_et_grippe.xlsx") as fichier:
    resultat.to_excel(fichier, sheet_name="Ventes et grippe")
    tableau.drop(columns="num_semaine").to_excel(fichier, sheet_name="Données par semaine")
print("\nFichiers créés dans donnees_externes/ : ventes_et_grippe.xlsx et paracetamol_grippe.png")
