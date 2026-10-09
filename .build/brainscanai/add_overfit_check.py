import ast
import copy
import json
from pathlib import Path

root = Path.cwd()
path = root / '01_exploration_dataset.ipynb'
raw = path.read_bytes()
before = json.loads(raw)
after = copy.deepcopy(before)
build = root / '.build/brainscanai/overfit-check'
build.mkdir(exist_ok=True)
assert not any(c.get('id') == 'cnn-comparaison-train-test' for c in before['cells'])
markdown = '''### Comparer le train et le test pour repérer un surapprentissage

Évaluer les deux CNN déjà entraînés sur les 63 images du train fort et les 20 images du test fort, avec les vrais labels normal/cancer. `evaluer_cnn` réalise les prédictions sans modifier les poids. Le chargeur d'évaluation réutilise les images du train fort, sans les mélanger.

Afficher les mêmes métriques sur les deux jeux, les écarts train moins test en points pour l'accuracy et le F2, puis un graphique du F2. Un score beaucoup plus élevé sur le train peut indiquer un surapprentissage. L'écart seul ne le prouve pas, surtout avec seulement 20 images de test. Cette comparaison sert au bilan final, pas au choix des hyperparamètres.'''
code = '''chargeur_train_evaluation = DataLoader(
    chargeur_fort.dataset,
    batch_size=taille_lot_etape4,
    shuffle=False,
    generator=torch.Generator().manual_seed(42),
)

resultats_train_test = []
for nom, cnn, mesures_test in [
    ("Semi-supervisé", modele_semi, mesures_semi),
    ("Supervisé", modele_supervise, mesures_supervise),
]:
    mesures_train, _ = evaluer_cnn(cnn, chargeur_train_evaluation)

    resultats_train_test.append({
        "Modèle": nom, "Jeu": "Train fort",
        "Images": len(chargeur_train_evaluation.dataset), **mesures_train,
    })
    resultats_train_test.append({
        "Modèle": nom, "Jeu": "Test fort",
        "Images": len(chargeur_test.dataset), **mesures_test,
    })

comparaison_train_test = pd.DataFrame(resultats_train_test)
display(comparaison_train_test.rename(columns={
    "F1": "F1 cancer", "Précision": "Précision cancer",
    "Rappel": "Rappel cancer", "F2": "F2 cancer",
}).round(3))

for nom in ["Semi-supervisé", "Supervisé"]:
    scores = comparaison_train_test.loc[
        comparaison_train_test["Modèle"] == nom
    ].set_index("Jeu")
    ecart_accuracy = 100 * (scores.loc["Train fort", "Accuracy"] - scores.loc["Test fort", "Accuracy"])
    ecart_f2 = 100 * (scores.loc["Train fort", "F2"] - scores.loc["Test fort", "F2"])
    print(f"{nom} : écart train - test = {ecart_accuracy:+.1f} points d'accuracy et {ecart_f2:+.1f} points de F2 cancer.")

scores_f2_train_test = comparaison_train_test.pivot(
    index="Modèle", columns="Jeu", values="F2"
).reindex(index=["Semi-supervisé", "Supervisé"], columns=["Train fort", "Test fort"])
axe = scores_f2_train_test.plot.bar(
    figsize=(7, 4), color=["#3978A9", "#D58A3D"], rot=0,
)
axe.set_title("F2 cancer : entraînement et test des deux CNN")
axe.set_xlabel("")
axe.set_ylabel("F2 cancer")
axe.set_ylim(0, 1.05)
axe.legend(title="Jeu")
plt.tight_layout()
plt.show()'''
ast.parse(code)
new_cells = [
    {'cell_type': 'markdown', 'id': 'cnn-surapprentissage-explication', 'metadata': {},
     'source': markdown.splitlines(keepends=True)},
    {'cell_type': 'code', 'id': 'cnn-comparaison-train-test', 'metadata': {},
     'source': code.splitlines(keepends=True), 'execution_count': None, 'outputs': []},
]
position = next(i for i,c in enumerate(after['cells']) if c['id']=='ee07cebf')+1
after['cells'][position:position] = new_cells
assert [c for c in after['cells'] if c['id'] not in {n['id'] for n in new_cells}] == before['cells']
assert after['metadata'] == before['metadata']
(build/'before.ipynb').write_bytes(raw)
assert path.read_bytes() == raw, 'Concurrent notebook edit'
path.write_text(json.dumps(after, ensure_ascii=False, indent=1)+'\n',encoding='utf-8')
print('Added Markdown/code pair after the final CNN comparison. All existing cells preserved.')
