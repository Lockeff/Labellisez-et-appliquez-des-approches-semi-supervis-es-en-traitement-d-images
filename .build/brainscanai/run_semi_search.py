"""Execute the requested search and edit only its existing code cell."""
from pathlib import Path
from contextlib import redirect_stdout
from io import StringIO
import json
import sys
import time

import pandas as pd
from IPython.core.formatters import DisplayFormatter

from search_semi_weak_first import SOURCE

ROOT = Path(__file__).resolve().parents[2]
PATH = ROOT / '01_exploration_dataset.ipynb'
TARGET = 'hyperparametres-semi-recherche'
initial = json.loads(PATH.read_bytes())
cells = {c['id']: c for c in initial['cells']}
env = {'__name__': '__main__'}
started = time.perf_counter()

def run(id):
    exec(''.join(cells[id]['source']), env)

print('Preparing the existing notebook prerequisites.', flush=True)
with redirect_stdout(StringIO()):
    run('db1c0311')
    # Use the persisted feature table to select exactly the retained images.
    env['dossier_dataset'] = ROOT / 'mri_dataset_brain_cancer_oc'
    env['chemin_features'] = ROOT / 'resultats/etape2/features_resnet18.csv'
    paths = pd.read_csv(env['chemin_features'], usecols=['chemin'])['chemin']
    env['chemins_images'] = [env['dossier_dataset'] / p for p in paths]
    env['colonnes_features'] = [f'feature_{i:03d}' for i in range(512)]
    for id in ['5597b309','d5c99912','1815ac75','d4bb2089','bca25a5c']:
        run(id)
    run('1ef4c71e')

# Limit CPU threading for this standalone execution; notebook settings are preserved.
env['torch'].set_num_threads(4)
print('Preparing images with the existing ImageNet preprocessing.', flush=True)
with redirect_stdout(StringIO()):
    run('375b1310')
    for id in ['ba2b3b8a','cc73ef88','f32fb809']:
        run(id)

# The search may evaluate only the reserved strong validation loader.
original_evaluate = env['evaluer_cnn']

def validation_only(cnn, loader):
    assert loader is env['chargeur_validation'], 'Only validation is allowed during search.'
    return original_evaluate(cnn, loader)

env['evaluer_cnn'] = validation_only
assert len(env['chargeur_faible'].dataset) == 1048
assert len(env['chargeur_fort'].dataset) == 63
assert len(env['chargeur_validation'].dataset) == 16
assert set(env['fortes_test'].index).isdisjoint(env['fortes_train'].index)
assert set(env['fortes_test'].index).isdisjoint(env['fortes_validation'].index)
outputs = []
formatter = DisplayFormatter()
pending = StringIO()

def flush_stream():
    text = pending.getvalue()
    if text:
        outputs.append({'name':'stdout','output_type':'stream','text':text.splitlines(keepends=True)})
        pending.seek(0)
        pending.truncate(0)

class Tee:
    def write(self, text):
        pending.write(text)
        sys.__stdout__.write(text)
        sys.__stdout__.flush()
        return len(text)
    def flush(self):
        sys.__stdout__.flush()

def display(value):
    flush_stream()
    data, metadata = formatter.format(value)
    outputs.append({'output_type':'display_data','data':data,'metadata':metadata})
    sys.__stdout__.write(value.to_string(index=False) + '\n')
    sys.__stdout__.flush()

env['display'] = display
print('Running two weak-only trials, then four strong trials from the best weak model.', flush=True)
with redirect_stdout(Tee()):
    exec(SOURCE, env)
    flush_stream()

assert len(env['tableau_recherche']) == 4
assert env['tableau_recherche_faibles']['epoques_faibles'].tolist() == [5,10]
assert env['epoques_faibles'] in [5,10]
assert env['optimiseur_semi'].param_groups[0]['lr'] == env['taux_apprentissage_semi']
assert env['taux_apprentissage_semi'] == env['taux_apprentissage_faible'], 'Selected phase rates differ; downstream cells require an explicit rate switch.'
assert len(env['fortes_validation']) == 16
assert len(env['faibles_train']) == 1048
assert 'train_test_split(' not in SOURCE

# Preserve any unrelated edits made by the user while training was running.
raw = PATH.read_bytes()
current = json.loads(raw)
target = next(c for c in current['cells'] if c['id'] == TARGET)
assert target['source'] == cells[TARGET]['source'], 'Search source changed during execution; stop.'
changed = json.loads(json.dumps(target))
changed['source'] = SOURCE.splitlines(keepends=True)
changed['outputs'] = outputs
changed['execution_count'] = None
text = raw.decode('utf-8')
decoder = json.JSONDecoder()
pos = text.index('[', text.index('"cells"')) + 1
while True:
    while text[pos].isspace() or text[pos] == ',': pos += 1
    assert text[pos] != ']', 'Target code cell not found.'
    cell, end = decoder.raw_decode(text, pos)
    if cell['id'] == TARGET:
        padding = text[text.rfind('\n', 0, pos) + 1:pos]
        replacement = json.dumps(changed, ensure_ascii=False, indent=1).replace('\n','\n'+padding)
        updated = text[:pos] + replacement + text[end:]
        break
    pos = end
after = json.loads(updated)
assert len(current['cells']) == len(after['cells'])
assert [c['id'] for c,d in zip(current['cells'],after['cells']) if c != d] == [TARGET]
assert current['metadata'] == after['metadata']
assert all(c == d for c,d in zip(current['cells'],after['cells']) if c['cell_type']=='markdown')
assert PATH.read_bytes() == raw, 'Concurrent notebook change; stop.'
PATH.write_bytes(updated.encode('utf-8'))
result = {
    'weak_epochs':env['epoques_faibles'],
    'strong_epochs':env['epoques_fortes_semi'],
    'learning_rate':env['taux_apprentissage_semi'],
    'validation_F2':float(env['meilleurs_parametres']['F2']),
    'weak_validation_accuracy':float(env['meilleurs_parametres_faibles']['Accuracy groupes K-Means validation']),
    'validation_images':16,
    'additional_split':False,
    'weak_search':env['tableau_recherche_faibles'].to_dict(orient='records'),
    'original_search':env['tableau_recherche'].to_dict(orient='records'),
    'elapsed_seconds':time.perf_counter()-started,
    'changed_cell_id':TARGET,
    'markdown_unchanged':True,
    'test_used':False,
}
(ROOT/'.build/brainscanai/weak-first-search-results.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
print('Saved only the requested search code cell and its actual outputs.', flush=True)
print(json.dumps(result,ensure_ascii=True), flush=True)
