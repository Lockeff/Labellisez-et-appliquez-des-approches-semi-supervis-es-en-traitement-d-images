SOURCE = '''from copy import deepcopy

# 1. Choisir les époques faibles avec les 16 images de validation existantes.
# Pour cette phase, les cibles sont leurs groupes K-Means, pas les labels médicaux.
groupes_validation_faible = kmeans.predict(X_validation_pca)
taux_apprentissage_faible = 0.001
resultats_recherche_faibles = []
meilleure_accuracy_faible = -1.0

for epoques_faibles_essai in [5, 10]:
    print(f"Essai faible : taux = {taux_apprentissage_faible}, époques = {epoques_faibles_essai}")

    modele_essai_faible = creer_cnn()
    optimiseur_essai_faible = torch.optim.Adam(
        modele_essai_faible.fc.parameters(), lr=taux_apprentissage_faible
    )
    entrainer_cnn(
        modele_essai_faible, chargeur_faible,
        optimiseur_essai_faible, epoques_faibles_essai
    )

    predictions_validation = []
    modele_essai_faible.eval()
    with torch.no_grad():
        for images_lot, _ in chargeur_validation:
            sorties = modele_essai_faible(images_lot.to(appareil_etape4))
            predictions_validation.extend(sorties.argmax(dim=1).cpu().tolist())

    accuracy_faible = float((
        np.array(predictions_validation) == groupes_validation_faible
    ).mean())
    resultats_recherche_faibles.append({
        "epoques_faibles": epoques_faibles_essai,
        "taux_apprentissage": taux_apprentissage_faible,
        "Accuracy groupes K-Means validation": accuracy_faible,
    })

    # Garder le meilleur modèle faible et son optimiseur pour la suite.
    # À score égal, conserver le premier essai : 5 époques.
    if accuracy_faible > meilleure_accuracy_faible:
        meilleure_accuracy_faible = accuracy_faible
        modele_faible_retenu = modele_essai_faible
        optimiseur_faible_retenu = optimiseur_essai_faible

tableau_recherche_faibles = pd.DataFrame(resultats_recherche_faibles)
display(tableau_recherche_faibles.round(4))
meilleurs_parametres_faibles = tableau_recherche_faibles.loc[
    tableau_recherche_faibles["Accuracy groupes K-Means validation"].idxmax()
]
epoques_faibles = int(meilleurs_parametres_faibles["epoques_faibles"])
print("Époques faibles retenues :", epoques_faibles)

# 2. Poursuivre le meilleur modèle faible pour rechercher les paramètres forts.
resultats_recherche = []

for taux_essai in [0.0001, 0.001]:
    for epoques_essai in [5, 10]:
        print(f"Essai fort : taux = {taux_essai}, époques = {epoques_essai}")

        # Chaque essai reprend les mêmes poids après la meilleure phase faible.
        modele_essai = creer_cnn()
        modele_essai.load_state_dict(modele_faible_retenu.state_dict())
        optimiseur_essai = torch.optim.Adam(
            modele_essai.fc.parameters(), lr=taux_essai
        )
        # Copier l'état de l'optimiseur pour garder les essais indépendants.
        optimiseur_essai.load_state_dict(deepcopy(optimiseur_faible_retenu.state_dict()))
        for groupe in optimiseur_essai.param_groups:
            groupe["lr"] = taux_essai

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

# 3. Préparer le CNN initial pour appliquer les paramètres dans les cellules suivantes.
modele_semi = creer_cnn()
optimiseur_semi = torch.optim.Adam(
    modele_semi.fc.parameters(), lr=taux_apprentissage_semi
)

print("Taux utilisé pour la recherche faible :", taux_apprentissage_faible)
print("Époques faibles retenues :", epoques_faibles)
print("Taux retenu pour la phase forte :", taux_apprentissage_semi)
print("Époques fortes retenues :", epoques_fortes_semi)
print("Meilleur F2 cancer sur la validation forte :", round(meilleurs_parametres["F2"], 3))
'''
