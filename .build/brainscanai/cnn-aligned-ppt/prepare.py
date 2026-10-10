import base64
import json
from html.parser import HTMLParser
from pathlib import Path

build = Path('.build/brainscanai/cnn-aligned-ppt')
n = json.loads(Path('01_exploration_dataset.ipynb').read_bytes())
cells = {c.get('id'): c for c in n['cells']}
class TableReader(HTMLParser):
    def __init__(self):
        super().__init__()
        self.rows, self.row, self.value = [], [], None
    def handle_starttag(self, tag, attrs):
        if tag == 'tr':
            self.row = []
        if tag in ['th', 'td']:
            self.value = ''
    def handle_data(self, data):
        if self.value is not None:
            self.value += data
    def handle_endtag(self, tag):
        if tag in ['th', 'td'] and self.value is not None:
            self.row.append(' '.join(self.value.split()))
            self.value = None
        if tag == 'tr':
            self.rows.append(self.row)

def table(cid, header):
    for o in cells[cid].get('outputs', []):
        html = o.get('data', {}).get('text/html')
        if html:
            parser = TableReader()
            parser.feed(''.join(html) if isinstance(html, list) else html)
            if header in parser.rows[0]:
                return [{key: value for key, value in zip(parser.rows[0], row) if key} for row in parser.rows[1:]]
    raise AssertionError((cid, header))

facts = {
    'semi_grid': table('hyperparametres-semi-recherche', 'F2 cancer validation'),
    'super_grid': table('hyperparametres-supervise-recherche', 'F2 cancer validation'),
    'weak_grid': table('hyperparametres-semi-recherche', 'Accuracy groupes K-Means validation'),
    'weak_test': table('b184ba1a', 'Accuracy'),
    'semi_test': table('82faeff6', 'Accuracy'),
    'super_test': table('9dfb6484', 'Accuracy'),
    'train_test': table('cnn-comparaison-train-test', 'F2 cancer'),
}
assert all(float(row[metric]) == .9 for field in ['semi_test', 'super_test'] for row in facts[field] for metric in ['Accuracy', 'F1 cancer', 'Précision cancer', 'Rappel cancer', 'F2 cancer'])
assert [float(row['F2 cancer validation']) for row in facts['semi_grid']] == [.4054, .5263, .5263, .8537]
assert [float(row['F2 cancer validation']) for row in facts['super_grid']] == [.5, .625, .6579, .9756]
assert float(facts['weak_test'][0]['Accuracy']) == .973
(build / 'facts.json').write_text(json.dumps(facts, ensure_ascii=False, indent=2), encoding='utf-8')
for o in cells['ee07cebf'].get('outputs', []):
    png = o.get('data', {}).get('image/png')
    if png:
        (build / 'new-confusions.png').write_bytes(base64.b64decode(''.join(png) if isinstance(png, list) else png))
print('Saved notebook results extracted and checked; notebook untouched.')
