"""Execute in a fresh local kernel; preserve step 1 and save only affected outputs."""
from pathlib import Path
import hashlib
import json
import sys
import time
import copy

import nbformat
from nbclient import NotebookClient
from jupyter_client import KernelManager

ROOT = Path(__file__).resolve().parents[2]
PATH = ROOT / '01_exploration_dataset.ipynb'
BUILD = ROOT / '.build/brainscanai/mentor-notebook'
initial_bytes = PATH.read_bytes()
initial = json.loads(initial_bytes)
notebook = nbformat.reads(initial_bytes.decode('utf-8'), as_version=4)
started = time.perf_counter()

def on_start(cell, cell_index, **kwargs):
    if cell.cell_type == 'code':
        print(f'{time.perf_counter()-started:.1f}s | executing {cell_index}: {cell.get("id")}', flush=True)

def on_done(cell, cell_index, **kwargs):
    for output in cell.get('outputs', []):
        if output.get('output_type') == 'stream':
            text = output.get('text', '')
            for line in text.splitlines()[-3:]:
                print('  '+line, flush=True)

notebook.cells.append(nbformat.v4.new_code_cell('''
import json
assert matrice_features.shape == (len(chemins_images), 512)
assert matrice_features_intermediaires.shape == (len(chemins_images), 256)
assert len(fortes_test) == 20
assert set(fortes_test.index).isdisjoint(donnees_ari.index)
assert len(donnees_ari) == 79
image_verification = read_image(str(chemins_images[0]), mode=ImageReadMode.RGB)
assert torch.equal(
    pretraitement(image_verification),
    poids.transforms()(equalize(image_verification)),
)
assert features["chemin"].tolist() == [p.relative_to(dossier_dataset).as_posix() for p in chemins_images]
assert len(labels_predits_mesure) == 1000
assert set(bilan_ressources.index) == {"Création du modèle", "Traitement final"}
assert (details_ressources["Pic RAM (Mio)"] >= details_ressources["RAM au départ (Mio)"]).all()
rapport = {
    "threads_cpu": torch.get_num_threads(),
    "appareil": str(appareil_etape4),
    "ari_layer4": float(ari_kmeans),
    "ari_layer3": float(ari_intermediaire),
    "epoques_faibles": int(epoques_faibles),
    "epoques_fortes_semi": int(epoques_fortes_semi),
    "epoques_fortes_supervise": int(epoques_fortes_supervise),
    "resultats_semi": mesures_semi,
    "resultats_supervise": mesures_supervise,
    "confusion_semi": confusion_semi.tolist(),
    "confusion_supervise": confusion_supervise.tolist(),
    "groupes_dbscan": [int(e["nombre_groupes"]) for e in essais_dbscan],
    "heures_estimees_4m": float(heures_estimees),
    "bilan_ressources": bilan_ressources.to_dict(orient="index"),
    "details_ressources": details_ressources.to_dict(orient="records"),
}
Path(".build/brainscanai/mentor-notebook/runtime-results.json").write_text(
    json.dumps(rapport, ensure_ascii=False, indent=2), encoding="utf-8"
)
print("Checks passed: preprocessing, layer shapes, preserved test, row alignment, RAM, timing.")
'''))

km = KernelManager(kernel_name='python3')
km.kernel_spec.argv = [sys.executable, '-m', 'ipykernel_launcher', '-f', '{connection_file}']
client = NotebookClient(
    notebook, km=km, timeout=1800, resources={'metadata': {'path': str(ROOT)}},
    on_cell_start=on_start, on_cell_executed=on_done,
)
try:
    client.execute()
finally:
    nbformat.write(notebook, BUILD / 'executed-private.ipynb')

if PATH.read_bytes() != initial_bytes:
    raise RuntimeError('Notebook changed during execution: outputs retained privately; no overwrite.')
executed = {cell['id']: cell for cell in notebook.cells[:-1]}
step_two = next(i for i, c in enumerate(initial['cells']) if c['id'] == 'ba54e8fb')
for cell in initial['cells'][step_two:]:
    if cell['cell_type'] == 'code':
        result = executed[cell['id']]
        cell['outputs'] = result['outputs']
        cell['execution_count'] = result['execution_count']
before = json.loads((BUILD / 'original.ipynb').read_bytes())
assert initial['cells'][:28] == before['cells'][:28]
PATH.write_text(json.dumps(initial, ensure_ascii=False, indent=1)+'\n', encoding='utf-8')
print('Execution complete; results saved; step 1 preserved.', flush=True)
