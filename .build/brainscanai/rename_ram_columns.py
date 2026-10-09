import ast
import copy
import json
from datetime import datetime
from html.parser import HTMLParser
from pathlib import Path

import pandas as pd

path = Path('01_exploration_dataset.ipynb')
original_text = path.read_text(encoding='utf-8')
original = json.loads(original_text)
notebook = copy.deepcopy(original)
cells = {cell['id']: cell for cell in notebook['cells']}
names = {
    'RAM au départ (Mio)': 'RAM totale avant la cellule (Mio)',
    'Pic RAM (Mio)': 'RAM totale maximale pendant la cellule (Mio)',
    'Hausse maximale (Mio)': 'Augmentation maximale pendant la cellule (Mio)',
}

cell = cells['mentor-bilan-ressources']
old_source = ''.join(cell['source'])
new_source = '''details_ressources = pd.DataFrame(mesures_ressources.values()).rename(columns={
    "RAM au départ (Mio)": "RAM totale avant la cellule (Mio)",
    "Pic RAM (Mio)": "RAM totale maximale pendant la cellule (Mio)",
    "Hausse maximale (Mio)": "Augmentation maximale pendant la cellule (Mio)",
})
display(details_ressources.round(2))

bilan_ressources = details_ressources.groupby("Groupe", sort=False).agg({
    "Durée (s)": "sum",
    "RAM totale maximale pendant la cellule (Mio)": "max",
}).rename(columns={
    "RAM totale maximale pendant la cellule (Mio)": "RAM totale maximale (Mio)",
})
bilan_ressources["Durée (min)"] = bilan_ressources["Durée (s)"] / 60
display(bilan_ressources.round(2))'''
ast.parse(new_source)
cell['source'] = new_source.splitlines(keepends=True)

markdown = cells['mentor-bilan-ressources-explication']
description = ''.join(markdown['source'])
description = description.replace(
    "La RAM au départ comprend les tableaux et modèles déjà en mémoire ; la hausse maximale distingue les allocations supplémentaires.",
    "La RAM totale avant la cellule comprend les tableaux et modèles déjà en mémoire. "
    "La RAM totale maximale pendant la cellule inclut cette mémoire déjà occupée. "
    "L'augmentation maximale indique la différence entre ce maximum et la RAM avant la cellule."
)
markdown['source'] = description.splitlines(keepends=True)


class TableReader(HTMLParser):
    def __init__(self):
        super().__init__()
        self.rows = []
        self.row = None
        self.cell = None

    def handle_starttag(self, tag, attrs):
        if tag == 'tr':
            self.row = []
        elif tag in ('td', 'th'):
            self.cell = []

    def handle_data(self, data):
        if self.cell is not None:
            self.cell.append(data)

    def handle_endtag(self, tag):
        if tag in ('td', 'th'):
            self.row.append(''.join(self.cell))
            self.cell = None
        elif tag == 'tr':
            self.rows.append(self.row)
            self.row = None


display_outputs = [o for o in cell['outputs'] if 'text/html' in o.get('data', {})]
assert len(display_outputs) == 2
for index, output in enumerate(display_outputs):
    html = ''.join(output['data']['text/html'])
    parser = TableReader()
    parser.feed(html)
    replacement = names if index == 0 else {'Pic RAM (Mio)': 'RAM totale maximale (Mio)'}
    headers = [replacement.get(h, h) for h in parser.rows[0][1:]]
    data_rows = parser.rows[1 if index == 0 else 2:]
    records = [[float(v) if j >= (2 if index == 0 else 0) else v
                for j, v in enumerate(row[1:])] for row in data_rows]
    frame = pd.DataFrame(records, columns=headers,
                         index=[int(row[0]) if index == 0 else row[0] for row in data_rows])
    if index == 1:
        frame.index.name = 'Groupe'
    for old, new in replacement.items():
        html = html.replace(old, new)
    output['data']['text/html'] = html.splitlines(keepends=True)
    with pd.option_context('display.max_columns', 20, 'display.width', 80):
        output['data']['text/plain'] = repr(frame.round(2)).splitlines(keepends=True)

# Verify that only display labels and their description have changed.
allowed = {'mentor-bilan-ressources', 'mentor-bilan-ressources-explication'}
assert notebook['metadata'] == original['metadata']
assert len(notebook['cells']) == len(original['cells'])
for before, after in zip(original['cells'], notebook['cells']):
    if before['id'] not in allowed:
        assert before == after, before['id']
    else:
        assert before['metadata'] == after['metadata']
        assert before.get('execution_count') == after.get('execution_count')
for before, after in zip(original['cells'][-1]['outputs'], cell['outputs']):
    if 'text/html' in before.get('data', {}):
        p1, p2 = TableReader(), TableReader()
        p1.feed(''.join(before['data']['text/html']))
        p2.feed(''.join(after['data']['text/html']))
        assert p1.rows[1:] == p2.rows[1:], 'Measured values changed'

backup = Path('.build/brainscanai/ram-column-rename')
backup.mkdir(exist_ok=True)
(backup / f'before-{datetime.now():%Y%m%d-%H%M%S}.ipynb').write_text(original_text, encoding='utf-8')
assert path.read_text(encoding='utf-8') == original_text, 'Concurrent notebook edit'
path.write_text(json.dumps(notebook, ensure_ascii=False, indent=1) + '\n', encoding='utf-8')
print('Renamed RAM columns in the final cell and saved displays; measured values preserved.')
