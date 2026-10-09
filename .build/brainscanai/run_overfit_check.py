import json
import sys
import time
from pathlib import Path

import nbformat
from nbclient import NotebookClient
from jupyter_client import KernelManager

root = Path.cwd()
path = root / '01_exploration_dataset.ipynb'
build = root / '.build/brainscanai/overfit-check'
raw = path.read_bytes()
original = json.loads(raw)
notebook = nbformat.reads(raw.decode('utf-8'),as_version=4)
target_index = next(i for i,c in enumerate(notebook.cells) if c.id == 'cnn-comparaison-train-test')
notebook.cells = notebook.cells[:target_index+1]
# Only the new cell's outputs will be saved. Existing feature CSV is read as-is.
for cell in notebook.cells:
    if cell.id == 'sauvegarde-features-csv':
        cell.source = 'chemin_features = Path("resultats/etape2/features_resnet18.csv")\nassert chemin_features.exists()'
check = '''
poids_semi_avant = {k:v.clone() for k,v in modele_semi.state_dict().items()}
poids_supervise_avant = {k:v.clone() for k,v in modele_supervise.state_dict().items()}
'''
notebook.cells[target_index].source = check + notebook.cells[target_index].source
notebook.cells.append(nbformat.v4.new_code_cell('''
assert all(torch.equal(v,modele_semi.state_dict()[k]) for k,v in poids_semi_avant.items())
assert all(torch.equal(v,modele_supervise.state_dict()[k]) for k,v in poids_supervise_avant.items())
assert len(chargeur_train_evaluation.dataset)==63
assert len(chargeur_test.dataset)==20
assert set(fortes_train.index).isdisjoint(fortes_test.index)
assert (comparaison_train_test.loc[comparaison_train_test.Jeu=="Test fort", "Accuracy"].to_numpy() == [mesures_semi["Accuracy"],mesures_supervise["Accuracy"]]).all()
Path(".build/brainscanai/overfit-check/results.json").write_text(comparaison_train_test.to_json(orient="records", force_ascii=False, indent=2),encoding="utf-8")
print("Check passed: unchanged model weights, separate train/test, consistent scores.")
'''))
start = time.perf_counter()
def progress(cell,cell_index,**kwargs):
    if cell.cell_type=='code':
        print(f'{time.perf_counter()-start:.0f}s: {cell.id}',flush=True)
km = KernelManager(kernel_name='python3')
km.kernel_spec.argv = [sys.executable,'-m','ipykernel_launcher','-f','{connection_file}']
client = NotebookClient(notebook,km=km,timeout=1800,
                        resources={'metadata':{'path':str(root)}},on_cell_start=progress)
try:
    client.execute()
finally:
    nbformat.write(notebook,build/'executed-private.ipynb')
assert path.read_bytes()==raw,'Notebook changed during execution; results kept privately.'
result = notebook.cells[target_index]
target = next(c for c in original['cells'] if c['id']=='cnn-comparaison-train-test')
target['outputs']=result['outputs']
target['execution_count']=result['execution_count']
before=json.loads((build/'before.ipynb').read_bytes())
new_ids={'cnn-surapprentissage-explication','cnn-comparaison-train-test'}
assert [c for c in original['cells'] if c['id'] not in new_ids] == before['cells']
path.write_text(json.dumps(original,ensure_ascii=False,indent=1)+'\n',encoding='utf-8')
print((build/'results.json').read_text(encoding='utf-8'),flush=True)
print('Saved only the new diagnostic outputs. Existing cells and results preserved.',flush=True)
