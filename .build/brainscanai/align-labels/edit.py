import json
from pathlib import Path

path = Path('01_exploration_dataset.ipynb')
before = path.read_bytes()
raw = before.decode('utf-8')
notebook = json.loads(raw)
changes = {}
cells = {c.get('id'): c for c in notebook['cells']}

cid = 'ba2b3b8a'
source = ''.join(cells[cid]['source'])
old = '''    faibles_train["cible"] = faibles_train["label_faible"].astype(int)
    faibles_test["cible"] = faibles_test["label_faible"].astype(int)
'''
new = '''    # Aligner les groupes K-Means sur normal = 0 et cancer = 1.
    # Utiliser uniquement le train fort pour choisir la correspondance.
    groupes_train_fort = kmeans.predict(X_fort_pca)
    cibles_train_fort = fortes_train["label_fort"].map({"normal": 0, "cancer": 1})
    correspondance_train = pd.crosstab(
        groupes_train_fort, cibles_train_fort.to_numpy(),
        rownames=["Groupe K-Means"], colnames=["Cible forte"],
    ).reindex(index=[0, 1], columns=[0, 1], fill_value=0)

    # Comparer les deux correspondances possibles entre les deux groupes.
    accord_direct = correspondance_train.loc[0, 0] + correspondance_train.loc[1, 1]
    accord_inverse = correspondance_train.loc[0, 1] + correspondance_train.loc[1, 0]
    correspondance_groupes_cibles = {0: 0, 1: 1}
    if accord_inverse > accord_direct:
        correspondance_groupes_cibles = {0: 1, 1: 0}

    display(correspondance_train)
    print("Correspondance groupe K-Means / cible CNN (normal = 0, cancer = 1) :",
          correspondance_groupes_cibles)

    # Garder les numéros K-Means dans label_faible et aligner les cibles du CNN.
    faibles_train["cible"] = faibles_train["label_faible"].map(
        correspondance_groupes_cibles
    ).astype(int)
    faibles_test["cible"] = faibles_test["label_faible"].map(
        correspondance_groupes_cibles
    ).astype(int)
'''
assert source.count(old) == 1
changes[cid] = source.replace(old, new)

cid = 'hyperparametres-semi-recherche'
source = ''.join(cells[cid]['source'])
old = '    groupes_validation_faible = kmeans.predict(X_validation_pca)\n'
new = '''    # Appliquer la même correspondance aux cibles faibles de validation.
    groupes_validation_faible = pd.Series(
        kmeans.predict(X_validation_pca)
    ).map(correspondance_groupes_cibles).to_numpy()
'''
assert source.count(old) == 1
changes[cid] = source.replace(old, new)

cid = 'b184ba1a'
source = ''.join(cells[cid]['source'])
assert 'index=["Test faible : groupes K-Means"]' in source
source = source.replace('index=["Test faible : groupes K-Means"]', 'index=["Test faible : cibles K-Means alignées"]')
for metric in ['F1', 'Précision', 'Rappel', 'F2']:
    assert f'"{metric} groupe 1"' in source
    source = source.replace(f'"{metric} groupe 1"', f'"{metric} cible faible 1"')
changes[cid] = source

for cid, source in changes.items():
    compile(source, cid, 'exec')

# Replace only the source JSON values, keeping every other byte untouched.
decoder = json.JSONDecoder()
position = raw.index('[', raw.index('"cells"')) + 1
replacements = []
while True:
    while raw[position].isspace() or raw[position] == ',':
        position += 1
    if raw[position] == ']':
        break
    cell, end = decoder.raw_decode(raw, position)
    cid = cell.get('id')
    if cid in changes:
        member_pos = position + 1
        while member_pos < end:
            while raw[member_pos].isspace() or raw[member_pos] == ',':
                member_pos += 1
            if raw[member_pos] == '}':
                break
            key, member_pos = decoder.raw_decode(raw, member_pos)
            while raw[member_pos].isspace() or raw[member_pos] == ':':
                member_pos += 1
            value_start = member_pos
            _, member_pos = decoder.raw_decode(raw, member_pos)
            if key == 'source':
                value = json.dumps(changes[cid].splitlines(keepends=True), ensure_ascii=False, indent=1)
                value = value.replace('\n', '\n   ')
                if '\r\n' in raw:
                    value = value.replace('\n', '\r\n')
                replacements.append((value_start, member_pos, value))
                break
    position = end
assert len(replacements) == len(changes)
for start, end, value in reversed(replacements):
    raw = raw[:start] + value + raw[end:]
after = json.loads(raw)
assert len(notebook['cells']) == len(after['cells'])
changed = []
for i, (old_cell, new_cell) in enumerate(zip(notebook['cells'], after['cells'])):
    if old_cell != new_cell:
        assert old_cell['cell_type'] == 'code' and old_cell['id'] in changes
        assert {k: v for k, v in old_cell.items() if k != 'source'} == {k: v for k, v in new_cell.items() if k != 'source'}
        changed.append(i + 1)
assert len(changed) == 3
assert {k: v for k, v in notebook.items() if k != 'cells'} == {k: v for k, v in after.items() if k != 'cells'}
build = Path('.build/brainscanai/align-labels')
(build / 'before.ipynb').write_bytes(before)
assert path.read_bytes() == before, 'Notebook changed during the edit; retry from latest version.'
path.write_bytes(raw.encode('utf-8'))
print('Code sources changed in cells:', changed)
print('All Markdown, outputs, metadata and other cells preserved.')
