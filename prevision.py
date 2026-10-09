import pandas as pd

pd.set_option("display.max_columns", None)
pd.set_option("display.width", 200)

GROUPES = ["M01AB", "M01AE", "N02BA", "N02BE", "N05B", "N05C", "R03", "R06"]

# 1. Ventes mensuelles reconstruites à partir du journalier (octobre 2019 retiré : incomplet)
jour = pd.read_csv("data/Pharma_Ventes_Daily.csv")
jour["datum"] = pd.to_datetime(jour["datum"], format="%m/%d/%Y")
mois = jour.set_index("datum")[GROUPES].resample("MS").sum()
mois = mois[mois.index < "2019-10-01"]


# 2. Les 4 méthodes de prévision.
#    Chaque méthode reçoit UNIQUEMENT l'historique connu avant le mois à prévoir.

def naive_mois_dernier(historique, mois_cible):
    """Naïve 1 : le mois prochain = le mois dernier."""
    return historique.iloc[-1]

def naive_annee_derniere(historique, mois_cible):
    """Naïve 2 : le mois prochain = le même mois l'an dernier."""
    return historique.loc[mois_cible - pd.DateOffset(years=1)]

def moyenne_3_mois(historique, mois_cible):
    """Modèle A : moyenne des 3 derniers mois."""
    return historique.iloc[-3:].mean()

def saison_et_niveau(historique, mois_cible):
    """Modèle B : niveau actuel x effet saisonnier du mois.
    - niveau = moyenne des 12 derniers mois
    - effet saisonnier = en moyenne, ce mois-là est X fois au-dessus ou en dessous de son année
    """
    niveau = historique.iloc[-12:].mean()
    moyenne_glissante = historique.rolling(12, center=True).mean()
    indice = (historique / moyenne_glissante).dropna()
    effet_saison = indice[indice.index.month == mois_cible.month].mean()
    return niveau * effet_saison


METHODES = {
    "Naïve : mois dernier": naive_mois_dernier,
    "Naïve : même mois an dernier": naive_annee_derniere,
    "Moyenne 3 derniers mois": moyenne_3_mois,
    "Saison x niveau": saison_et_niveau,
}

# 3. Le test : on se replace chaque mois de janv. 2018 à sept. 2019,
#    on prévoit le mois suivant avec ce qu'on savait à ce moment-là, puis on compare au réel.
mois_testes = mois.loc["2018-01-01":].index
lignes = []
for mois_cible in mois_testes:
    historique = mois[mois.index < mois_cible]
    reel = mois.loc[mois_cible]
    for nom, methode in METHODES.items():
        prevu = methode(historique, mois_cible)
        for g in GROUPES:
            lignes.append({"mois": mois_cible, "methode": nom, "groupe": g,
                           "reel": reel[g], "prevu": prevu[g]})
resultats = pd.DataFrame(lignes)
resultats["erreur"] = (resultats["prevu"] - resultats["reel"]).abs()
resultats["erreur %"] = resultats["erreur"] / resultats["reel"] * 100

# 4. Erreur moyenne en % (MAPE) par groupe et par méthode
mape = resultats.pivot_table(index="groupe", columns="methode", values="erreur %", aggfunc="mean")
mape = mape[list(METHODES)].round(1)
print(f"=== Erreur moyenne en % sur {len(mois_testes)} mois testés (plus c'est bas, mieux c'est) ===")
print(mape)

# 5. Pour chaque groupe : notre meilleur modèle bat-il la meilleure naïve ?
naives = ["Naïve : mois dernier", "Naïve : même mois an dernier"]
modeles = ["Moyenne 3 derniers mois", "Saison x niveau"]
verdict = pd.DataFrame({
    "Meilleure naïve (%)": mape[naives].min(axis=1),
    "Meilleur modèle (%)": mape[modeles].min(axis=1),
    "Modèle retenu": mape[modeles].idxmin(axis=1),
})
verdict["Gain (points)"] = (verdict["Meilleure naïve (%)"] - verdict["Meilleur modèle (%)"]).round(1)
# On exige au moins 2 points de gain pour dire "Oui" : en dessous, la différence peut être due au hasard
def conclure(gain):
    if gain >= 2:
        return "Oui"
    if gain > 0:
        return "Gain marginal"
    return "Non"

verdict["Prévisible ?"] = verdict["Gain (points)"].apply(conclure)
print("\n=== Verdict par groupe ===")
print(verdict)

# 6. Prévision d'octobre 2019 (mois qui suit le dernier mois complet)
octobre = pd.Timestamp("2019-10-01")
prevision = {}
for g in GROUPES:
    if verdict.loc[g, "Prévisible ?"] == "Oui":
        methode = METHODES[verdict.loc[g, "Modèle retenu"]]
    else:
        methode = METHODES[mape.loc[g, naives].idxmin()]
    prevision[g] = methode(mois, octobre)[g]
prevision = pd.Series(prevision).round(0)
print("\n=== Prévision octobre 2019 ===")
print(prevision.to_string())

# 7. Enregistrer les résultats
with pd.ExcelWriter("prevision.xlsx") as fichier:
    mape.to_excel(fichier, sheet_name="Erreur par méthode")
    verdict.to_excel(fichier, sheet_name="Verdict")
    prevision.to_frame("Prévision oct. 2019").to_excel(fichier, sheet_name="Prévision")
    resultats.to_excel(fichier, sheet_name="Détail du test", index=False)
print("\nFichier prevision.xlsx créé.")