"""Update only the ARI Markdown/code pair requested by the user."""
from pathlib import Path
from contextlib import redirect_stdout
from io import StringIO
import json

import numpy as np
import pandas as pd
from IPython.core.formatters import DisplayFormatter
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import StandardScaler
from sklearn.decomposition import PCA
from sklearn.cluster import KMeans
from sklearn.metrics import adjusted_rand_score

root = Path(__file__).resolve().parents[2]
path = root / '01_exploration_dataset.ipynb'
original_bytes = path.read_bytes()
original = original_bytes.decode('utf-8')
before = json.loads(original)
cells = {c['id']: c for c in before['cells']}
markdown_id = 'd6cd3bd7'
code_id = '1454a0b8'

markdown = """### 3.9. Calcul du score ARI de K-Means

Réunir les 63 images fortes d'entraînement et les 16 de validation pour évaluer les groupes K-Means sur 79 images. Leurs coordonnées PCA sont déjà calculées. Les 20 images fortes de test sont exclues.

Résultat : un tableau avec le nombre d'images évaluées et le score ARI.
"""
code = """# Réunir uniquement les données fortes d'entraînement et de validation.
donnees_ari = pd.concat([fortes_train, fortes_validation])
X_ari = np.concatenate([X_fort_pca, X_validation_pca], axis=0)

labels_reference = donnees_ari["label_fort"].to_numpy()
groupes_kmeans_forts = kmeans.predict(X_ari)
ari_kmeans = adjusted_rand_score(labels_reference, groupes_kmeans_forts)

scores_ari = pd.DataFrame({
    "Méthode": ["K-Means"],
    "Nombre d'images fortes": [len(donnees_ari)],
    "ARI train + validation": [ari_kmeans],
})
display(scores_ari)
"""

# Evaluate the existing prerequisites in memory without saving their outputs.
env = {
    'pd': pd, 'np': np, 'train_test_split': train_test_split,
    'StandardScaler': StandardScaler, 'PCA': PCA, 'KMeans': KMeans,
    'adjusted_rand_score': adjusted_rand_score,
    'chemin_features': root / 'resultats/etape2/features_resnet18.csv',
    'display': lambda value: None,
}
columns = pd.read_csv(env['chemin_features'], nrows=0).columns
env['colonnes_features'] = [c for c in columns if c.startswith('feature_')]
with redirect_stdout(StringIO()):
    for id in ['5597b309', 'd5c99912', '1815ac75', 'd4bb2089']:
        exec(''.join(cells[id]['source']), env)
outputs = []
formatter = DisplayFormatter()

def capture_display(value):
    data, metadata = formatter.format(value)
    outputs.append({'data': data, 'metadata': metadata, 'output_type': 'display_data'})

env['display'] = capture_display
exec(code, env)
assert len(env['donnees_ari']) == len(env['X_ari']) == 79
assert len(env['fortes_validation']) == 16
assert len(env['fortes_test']) == 20
assert set(env['donnees_ari'].index).isdisjoint(env['fortes_test'].index)
assert np.array_equal(env['X_ari'][:63], env['X_fort_pca'])
assert np.array_equal(env['X_ari'][63:], env['X_validation_pca'])
assert np.array_equal(env['labels_reference'][:63], env['fortes_train']['label_fort'])
assert np.array_equal(env['labels_reference'][63:], env['fortes_validation']['label_fort'])

replacements = {}
for id, source in [(markdown_id, markdown), (code_id, code)]:
    changed = json.loads(json.dumps(cells[id]))
    changed['source'] = source.splitlines(keepends=True)
    if id == code_id:
        changed['outputs'] = outputs
        changed['execution_count'] = None
    replacements[id] = changed

# Replace only the two JSON cell spans; all other bytes remain untouched.
decoder = json.JSONDecoder()
pos = original.index('[', original.index('"cells"')) + 1
spans = []
while True:
    while original[pos].isspace() or original[pos] == ',': pos += 1
    if original[pos] == ']': break
    cell, end = decoder.raw_decode(original, pos)
    if cell['id'] in replacements:
        padding = original[original.rfind('\n', 0, pos) + 1:pos]
        replacement = json.dumps(replacements[cell['id']], ensure_ascii=False, indent=1)
        spans.append((pos, end, replacement.replace('\n', '\n' + padding)))
    pos = end
assert len(spans) == 2
updated = original
for start, end, replacement in reversed(spans):
    updated = updated[:start] + replacement + updated[end:]
after = json.loads(updated)
changed_indices = [i for i, (a, b) in enumerate(zip(before['cells'], after['cells'])) if a != b]
assert len(before['cells']) == len(after['cells'])
assert {before['cells'][i]['id'] for i in changed_indices} == {markdown_id, code_id}
assert before['metadata'] == after['metadata']
assert path.read_bytes() == original_bytes, 'Notebook changed during verification; stop.'
path.write_bytes(updated.encode('utf-8'))
print('Changed only ARI Markdown and code cells:', changed_indices)
print(env['scores_ari'].to_string(index=False).encode('ascii', 'backslashreplace').decode())
print('Preserved validation: 16 images; final test: 20 images.')
