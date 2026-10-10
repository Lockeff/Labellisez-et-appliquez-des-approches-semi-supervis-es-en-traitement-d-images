import hashlib
import json
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
sys.stdout.reconfigure(encoding='utf-8')
build = Path('.build/brainscanai/cnn-aligned-ppt')
nb = Path('01_exploration_dataset.ipynb')
ppt = Path('output/presentations/BrainScanAI_presentation_V1_Final.pptx')
n = json.loads(nb.read_bytes())
ids = {'hyperparametres-semi-recherche','b184ba1a','82faeff6','hyperparametres-supervise-recherche','9dfb6484','ee07cebf','cnn-comparaison-train-test','a3f5cc5d','mentor-bilan-ressources'}
excerpts = []
for i, cell in enumerate(n['cells']):
    if cell.get('id') not in ids:
        continue
    lines = [f'CELL {i+1} {cell.get("id")}', ''.join(cell['source']), 'OUTPUTS']
    for o in cell.get('outputs', []):
        value = o.get('text', o.get('data', {}).get('text/plain', []))
        lines.append(''.join(value) if isinstance(value, list) else value)
        if o.get('output_type') == 'error':
            lines.append('ERROR: ' + o.get('ename', ''))
    excerpts.append('\n'.join(lines))
(build / 'notebook-results.txt').write_text('\n\n'.join(excerpts), encoding='utf-8')
ns = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main', 'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'}
slides = []
with zipfile.ZipFile(ppt) as z:
    for i in range(1, 14):
        root = ET.fromstring(z.read(f'ppt/slides/slide{i}.xml'))
        shapes = []
        for s in root.findall('./p:cSld/p:spTree/*', ns):
            props = s.find('.//p:cNvPr', ns)
            text = '\n'.join(''.join(t.text or '' for t in p.findall('.//a:t', ns)) for p in s.findall('.//a:p', ns))
            shapes.append({'id': props.get('id') if props is not None else None, 'name': props.get('name') if props is not None else None, 'text': text})
        slides.append({'slide': i, 'shapes': shapes})
(build / 'slides.json').write_text(json.dumps(slides, ensure_ascii=False, indent=2), encoding='utf-8')
(build / 'source-hashes.json').write_text(json.dumps({str(p): hashlib.sha256(p.read_bytes()).hexdigest() for p in [nb,ppt]}, indent=2), encoding='utf-8')
print('\n\n'.join(excerpts))
print('PPT shape inventory saved.')
