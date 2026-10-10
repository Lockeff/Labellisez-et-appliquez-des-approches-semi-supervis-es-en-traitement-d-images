import json
from pathlib import Path

import pandas as pd
from PIL import Image

before = json.loads(Path('.build/brainscanai/correction-compteur/before.ipynb').read_text(encoding='utf-8'))
after = json.loads(Path('01_exploration_dataset.ipynb').read_text(encoding='utf-8'))
changed = []
for old, new in zip(before['cells'], after['cells']):
    if old != new:
        changed.append(new['id'])
        old_other = {k: v for k, v in old.items() if k != 'source'}
        new_other = {k: v for k, v in new.items() if k != 'source'}
        assert old_other == new_other
assert changed == ['60838683', 'eacb5d90'], changed
assert len(before['cells']) == len(after['cells'])
assert {k: v for k, v in before.items() if k != 'cells'} == {k: v for k, v in after.items() if k != 'cells'}

dataset = Path('mri_dataset_brain_cancer_oc')
source = ''.join(next(c for c in after['cells'] if c.get('id') == '60838683')['source'])
env = {
    'dossier_dataset': dataset, 'Image': Image, 'pd': pd,
    'display': lambda *args: None,
    'caracteristiques': [{'Fichier': 'ancienne ligne'}] * 100,
    'erreurs': [{'Erreur': 'ancienne erreur'}],
}
exec(source, env)
assert len(env['controle']) == len(list(dataset.rglob('*.jpg'))) == 1506
assert len(env['erreurs']) == 0
print('Verification : 1506 images, sans recompter les 100 images precedentes.')
print('Seules les deux sources demandees ont change ; sorties et metadonnees conservees.')
