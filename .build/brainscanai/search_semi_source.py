SOURCE = '''# Réutiliser le train et la validation réservés à l’étape 3.
resultats_recherche = []

for taux_essai in [0.0001, 0.001]:
    for epoques_essai in [5, 10]:
        print(f"Essai : taux = {taux_essai}, époques fortes = {epoques_essai}")

        # Chaque essai repart des mêmes poids pré-entraînés.
        modele_essai = creer_cnn()
        optimiseur_essai = torch.optim.Adam(
            modele_essai.fc.parameters(), lr=taux_essai
        )

        entrainer_cnn(
            modele_essai, chargeur_faible, optimiseur_essai, 5
        )
        entrainer_cnn(
            modele_essai, chargeur_fort, optimiseur_essai, epoques_essai
        )
        mesures_validation, _ = evaluer_cnn(modele_essai, chargeur_validation)

        resultats_recherche.append({
            "taux_apprentissage": taux_essai,
            "epoques_fortes": epoques_essai,
            **mesures_validation,
        })

tableau_recherche = pd.DataFrame(resultats_recherche)
display(tableau_recherche.rename(columns={
    "Accuracy": "Accuracy validation",
    "F1": "F1 cancer validation",
    "Précision": "Précision cancer validation",
    "Rappel": "Rappel cancer validation",
    "F2": "F2 cancer validation",
}).round(4))

meilleurs_parametres = tableau_recherche.loc[tableau_recherche["F2"].idxmax()]
taux_apprentissage_semi = float(meilleurs_parametres["taux_apprentissage"])
epoques_fortes_semi = int(meilleurs_parametres["epoques_fortes"])

# Comparer 5 et 10 époques faibles avec le taux et les époques fortes retenus.
resultats_recherche_faibles = []

for epoques_faibles_essai in [5, 10]:
    print(
        f"Essai faible : époques faibles = {epoques_faibles_essai}, "
        f"taux = {taux_apprentissage_semi}, époques fortes = {epoques_fortes_semi}"
    )

    modele_essai_faible = creer_cnn()
    optimiseur_essai_faible = torch.optim.Adam(
        modele_essai_faible.fc.parameters(), lr=taux_apprentissage_semi
    )

    entrainer_cnn(
        modele_essai_faible, chargeur_faible,
        optimiseur_essai_faible, epoques_faibles_essai
    )
    entrainer_cnn(
        modele_essai_faible, chargeur_fort,
        optimiseur_essai_faible, epoques_fortes_semi
    )
    mesures_validation_faible, _ = evaluer_cnn(
        modele_essai_faible, chargeur_validation
    )

    resultats_recherche_faibles.append({
        "epoques_faibles": epoques_faibles_essai,
        "taux_apprentissage": taux_apprentissage_semi,
        "epoques_fortes": epoques_fortes_semi,
        **mesures_validation_faible,
    })

tableau_recherche_faibles = pd.DataFrame(resultats_recherche_faibles)
display(tableau_recherche_faibles.rename(columns={
    "Accuracy": "Accuracy validation",
    "F1": "F1 cancer validation",
    "Précision": "Précision cancer validation",
    "Rappel": "Rappel cancer validation",
    "F2": "F2 cancer validation",
}).round(4))

meilleurs_parametres_faibles = tableau_recherche_faibles.loc[
    tableau_recherche_faibles["F2"].idxmax()
]
epoques_faibles = int(meilleurs_parametres_faibles["epoques_faibles"])

# Préparer le modèle final pour l'entraînement faible, puis fort.
modele_semi = creer_cnn()
optimiseur_semi = torch.optim.Adam(
    modele_semi.fc.parameters(), lr=taux_apprentissage_semi
)

print("Taux retenu pour le semi-supervisé :", taux_apprentissage_semi)
print("Époques faibles retenues :", epoques_faibles)
print("Époques fortes retenues :", epoques_fortes_semi)
print("Meilleur F2 cancer sur la validation :", round(meilleurs_parametres_faibles["F2"], 3))
'''
