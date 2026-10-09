"""Remove console-only duplicate tables from the captured cell output."""
from pathlib import Path
import json
import pandas as pd

ROOT = Path(__file__).resolve().parents[2]
PATH = ROOT/'01_exploration_dataset.ipynb'
results = json.loads((ROOT/'.build/brainscanai/weak-search-results.json').read_text(encoding='utf-8'))
raw = PATH.read_bytes()
text = raw.decode('utf-8')
before = json.loads(text)
target = next(c for c in before['cells'] if c['id']=='hyperparametres-semi-recherche')
cell = json.loads(json.dumps(target))
rename = {
    'Accuracy':'Accuracy validation', 'F1':'F1 cancer validation',
    'Précision':'Précision cancer validation', 'Rappel':'Rappel cancer validation',
    'F2':'F2 cancer validation',
}
duplicates = [pd.DataFrame(results[key]).rename(columns=rename).round(4).to_string(index=False)+'\n'
              for key in ['original_search','weak_search']]
removed = 0
for out in cell['outputs']:
    if out['output_type']=='stream':
        content = ''.join(out['text'])
        for duplicate in duplicates:
            if content.startswith(duplicate):
                content = content[len(duplicate):]
                removed += 1
        out['text'] = content.splitlines(keepends=True)
assert removed == 2
cell['outputs'] = [out for out in cell['outputs'] if out['output_type']!='stream' or out['text']]
assert sum(out['output_type']=='display_data' for out in cell['outputs'])==2
assert results['weak_epochs']==5
assert all(row['F2']==results['validation_F2'] for row in results['weak_search'])
assert results['test_used'] is False
decoder = json.JSONDecoder()
pos = text.index('[',text.index('"cells"'))+1
while True:
    while text[pos].isspace() or text[pos]==',': pos+=1
    found,end=decoder.raw_decode(text,pos)
    if found['id']==cell['id']:
        padding=text[text.rfind('\n',0,pos)+1:pos]
        replacement=json.dumps(cell,ensure_ascii=False,indent=1).replace('\n','\n'+padding)
        updated=text[:pos]+replacement+text[end:]
        break
    pos=end
after=json.loads(updated)
assert [c['id'] for c,d in zip(before['cells'],after['cells']) if c!=d]==[cell['id']]
assert all(c==d for c,d in zip(before['cells'],after['cells']) if c['cell_type']=='markdown')
assert target['source']==cell['source']
assert PATH.read_bytes()==raw
PATH.write_bytes(updated.encode('utf-8'))
print('Verified: two result tables, equal F2 scores, 5 weak epochs selected.')
print('All Markdown, other code cells and final CNN outputs preserved.')
