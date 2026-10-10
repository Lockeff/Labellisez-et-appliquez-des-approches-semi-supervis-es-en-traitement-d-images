import copy
import json
import sys
from contextlib import nullcontext
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8')
import numpy as np
import pandas as pd
from sklearn.preprocessing import StandardScaler
from sklearn.decomposition import PCA
from sklearn.cluster import KMeans
from sklearn.model_selection import train_test_split

path = Path('01_exploration_dataset.ipynb')
before = json.loads(Path('.build/brainscanai/align-labels/before.ipynb').read_bytes())
after = json.loads(path.read_bytes())
cells = {c.get('id'): c for c in after['cells']}
env = {
    'np': np, 'pd': pd, 'StandardScaler': StandardScaler, 'PCA': PCA,
    'KMeans': KMeans, 'train_test_split': train_test_split,
    'mesurer_ressources': lambda *a, **k: nullcontext(), 'display': lambda *a: None,
    'chemin_features': Path('resultats/etape2/features_resnet18.csv'),
    'colonnes_features': [f'feature_{i:03d}' for i in range(512)],
}
for cid in ['5597b309', 'd5c99912', '1815ac75', 'd4bb2089', 'bca25a5c']:
    exec(''.join(cells[cid]['source']), env)

source = ''.join(cells['ba2b3b8a']['source'])
mapping_code = source[:source.index('    taille_lot_etape4 =')]
validation_source = ''.join(cells['hyperparametres-semi-recherche']['source'])
validation_code = validation_source[:validation_source.index('    taux_apprentissage_faible =')]
original_kmeans = env['kmeans']
original_train_groups = env['faibles_train']['label_faible'].copy()
original_test_groups = env['faibles_test']['label_faible'].copy()
exec(mapping_code, env)
exec(validation_code, env)
mapping = env['correspondance_groupes_cibles']
assert mapping == {0: 1, 1: 0}, mapping
assert env['faibles_train']['label_faible'].equals(original_train_groups)
assert env['faibles_test']['label_faible'].equals(original_test_groups)
assert set(env['faibles_train']['cible']) == {0, 1}
train_targets = env['faibles_train']['cible'].copy()
test_targets = env['faibles_test']['cible'].copy()
validation_targets = env['groupes_validation_faible'].copy()
np.testing.assert_array_equal(validation_targets, 1 - original_kmeans.predict(env['X_validation_pca']))
assert env['fortes_train']['cible'].equals(env['fortes_train']['label_fort'].map({'normal': 0, 'cancer': 1}))

# The same code must also work if K-Means renumbers its two groups.
class RenumberedKMeans:
    def predict(self, values):
        return 1 - original_kmeans.predict(values)

env['kmeans'] = RenumberedKMeans()
env['faibles_train']['label_faible'] = 1 - original_train_groups
env['faibles_test']['label_faible'] = 1 - original_test_groups
exec(mapping_code, env)
exec(validation_code, env)
assert env['correspondance_groupes_cibles'] == {0: 0, 1: 1}
assert env['faibles_train']['cible'].equals(train_targets)
assert env['faibles_test']['cible'].equals(test_targets)
np.testing.assert_array_equal(env['groupes_validation_faible'], validation_targets)

for old, new in zip(before['cells'], after['cells']):
    if old != new:
        assert new['id'] in {'ba2b3b8a', 'hyperparametres-semi-recherche', 'b184ba1a'}
        assert {k:v for k,v in old.items() if k != 'source'} == {k:v for k,v in new.items() if k != 'source'}
    if new['cell_type'] == 'markdown':
        assert old == new
assert json.loads(path.read_bytes()) == after
print('Both K-Means numberings give identical aligned CNN targets.')
print('Weak train/test and weak validation agree on the same encoding.')
print('Markdown, outputs and out-of-scope cells unchanged. No CNN retraining performed.')
