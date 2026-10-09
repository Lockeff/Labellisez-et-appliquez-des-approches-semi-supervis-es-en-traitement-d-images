from pathlib import Path
import json
import ast
import copy
import base64
import re

ROOT=Path(__file__).resolve().parents[2]
PATH=ROOT/'01_exploration_dataset.ipynb'
BUILD=ROOT/'.build/brainscanai/mentor-notebook'
n=json.loads(PATH.read_bytes())
o=json.loads((BUILD/'original.ipynb').read_bytes())
report=json.loads((BUILD/'runtime-results.json').read_bytes())
cells={c['id']:c for c in n['cells']}
old={c['id']:c for c in o['cells']}

def replace_source(cell_id, before, after):
    value=''.join(cells[cell_id]['source'])
    assert before in value
    cells[cell_id]['source']=value.replace(before,after).splitlines(keepends=True)

replace_source('mentor-ressources-explication',
    'Le tableau récapitulatif sera affiché après le traitement final.',
    'Résultat : la fonction de mesure et le dictionnaire de résultats sont prêts.')
replace_source('mentor-intermediaire-explication',
    'Cet essai reste en mémoire ; le tableau CSV suivant conserve les embeddings habituels de 512 dimensions.',
    'Cet essai reste en mémoire et ne remplace pas `matrice_features`, qui contient les embeddings habituels de 512 dimensions.')
replace_source('421c9b11', '0,346', f'{report["ari_layer4"]:.3f}'.replace('.',','))
cells['30295dde']['source']=[
    'Avec `eps=0.5` et `min_samples=10`, DBSCAN produit deux groupes, mais laisse '
    '1 028 des 1 048 images hors groupe. Les autres essais produisent un nombre de groupes différent de deux. '
    'Je conserve donc K-Means pour la labellisation faible : il répartit toutes les images en deux groupes.\n'
]

# Keep the existing ARI interpretation next to its original calculation.
trial_ids={'mentor-intermediaire-ari-explication','mentor-intermediaire-ari'}
trial=[c for c in n['cells'] if c['id'] in trial_ids]
n['cells']=[c for c in n['cells'] if c['id'] not in trial_ids]
position=next(i for i,c in enumerate(n['cells']) if c['id']=='421c9b11')+1
n['cells'][position:position]=trial

# Verify source preservation outside the explicitly affected computations.
edited=json.loads((BUILD/'edited-cell-ids.json').read_bytes())
allowed=set(edited)|{'421c9b11','30295dde'}
for c in o['cells']:
    if c['id'] not in allowed:
        assert cells[c['id']]['source']==c['source'],c['id']
    assert cells[c['id']]['metadata']==c['metadata'],c['id']
assert n['cells'][:28]==o['cells'][:28]
assert all(n[k]==v for k,v in o.items() if k!='cells')
assert ''.join(cells['mentor-egalisation-explication']['source'])=='Egalisation des histogrammes'

for c in n['cells']:
    if c['cell_type']=='code':
        ast.parse(''.join(c['source']))
        assert not any(out.get('output_type')=='error' for out in c.get('outputs',[])),c['id']
for i,c in enumerate(n['cells']):
    if c['id'].startswith('mentor-') and c['cell_type']=='markdown':
        assert n['cells'][i+1]['cell_type']=='code'
for cell_id in edited:
    if cell_id in old and old[cell_id]['cell_type']=='code' and cell_id!='a3f5cc5d':
        tree=ast.parse(''.join(cells[cell_id]['source']))
        assert len(tree.body)==1 and isinstance(tree.body[0],ast.With)
        unwrapped=ast.Module(body=tree.body[0].body,type_ignores=[])
        assert ast.dump(unwrapped)==ast.dump(ast.parse(''.join(old[cell_id]['source']))),cell_id

# Keep the inference example and the two-row report next to each other.
assert cells['mentor-bilan-ressources']['outputs']
assert len(report['details_ressources'])==22
assert abs(sum(r['Durée (s)'] for r in report['details_ressources'] if r['Groupe']=='Création du modèle')
           -report['bilan_ressources']['Création du modèle']['Durée (s)'])<1e-6
assert report['bilan_ressources']['Création du modèle']['Pic RAM (Mio)']==max(
    r['Pic RAM (Mio)'] for r in report['details_ressources'] if r['Groupe']=='Création du modèle')

# Extract changed plot outputs for a visual check.
for cell_id,name in [('f59cda16','kmeans-equalized'),('ee07cebf','cnn-confusion-equalized')]:
    for out in cells[cell_id]['outputs']:
        value=out.get('data',{}).get('image/png')
        if value:
            if isinstance(value,list):value=''.join(value)
            (BUILD/f'{name}.png').write_bytes(base64.b64decode(value))

PATH.write_text(json.dumps(n,ensure_ascii=False,indent=1)+'\n',encoding='utf-8')
print('Scope, compilation, paired cells, fresh outputs, arithmetic and model code verified.')
print('Step 1, shared imports, existing metadata and all learning logic preserved.')
