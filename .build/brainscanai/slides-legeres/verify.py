import hashlib
import json
import zipfile
from decimal import Decimal as D
from pathlib import Path
from xml.etree import ElementTree as ET

build = Path('.build/brainscanai/slides-legeres')
source = Path('output/presentations/BrainScanAI_presentation_13_slides_CPU_AWS.pptx')
candidate = build / 'candidate.pptx'
ns = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main', 'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'}
def note_body(xml):
    for shape in ET.fromstring(xml).findall('.//p:sp', ns):
        placeholder = shape.find('.//p:ph', ns)
        if placeholder is not None and placeholder.get('type') == 'body':
            return ''.join(t.text or '' for t in shape.findall('.//a:t', ns))
    raise AssertionError('Missing speaker notes')
with zipfile.ZipFile(source) as original, zipfile.ZipFile(candidate) as revised:
    for i in range(1, 10):
        p = f'ppt/slides/slide{i}.xml'
        assert ET.tostring(ET.fromstring(original.read(p))) == ET.tostring(ET.fromstring(revised.read(p)))
        png = f'slide-{i:02}.png'
        assert (build / 'source-renders' / png).read_bytes() == (build / 'candidate-renders' / png).read_bytes()
    for name in original.namelist():
        if name.startswith(('ppt/charts/', 'ppt/embeddings/')):
            assert original.read(name) == revised.read(name)
    for i in range(10, 14):
        p = f'ppt/notesSlides/notesSlide{i}.xml'
        old = note_body(original.read(p))
        new = note_body(revised.read(p))
        assert old in new, f'Detailed notes lost on slide {i}'
    visible = {}
    for i in range(10, 14):
        root = ET.fromstring(revised.read(f'ppt/slides/slide{i}.xml'))
        assert not root.findall('.//a:tbl', ns), 'Dense table remains'
        visible[i] = ' '.join(t.text or '' for t in root.findall('.//a:t', ns))
    for text in ['3,2 Gio', '96 Gio', '2,28', '24 cœurs physiques']:
        assert text in visible[10]
    for text in ['72,55', '68,39', '4,16', '30 heures', '300']:
        assert text in visible[11]
    for text in ['67,03', '101,78', '8,2', '41 minutes', '333 333', '12 mois', 'Hypothèse']:
        assert text in visible[12]
    assert D('68.39') + D('4.16') == D('72.55')
    costs = json.loads(Path('.build/brainscanai/aws-i9/costs-fixed-fx.json').read_text(encoding='utf-8'))
    rounded = lambda k: D(costs['display'][k].replace(',', '.'))
    assert rounded('prototype_total') == D('72.55')
    assert sum(rounded(k) for k in ['compute_4m', 'storage_year', 'disk_month', 'requests']) == rounded('one_shot') == D('67.03')
    assert sum(rounded(k) for k in ['compute_4m', 'storage_rolling', 'disk_year', 'requests']) == rounded('rolling_year1') == D('101.78')
receipt = json.loads((build / 'edit-receipt.json').read_text(encoding='utf-8-sig'))
assert hashlib.sha256(source.read_bytes()).hexdigest().upper() == receipt['SourceHash']
assert hashlib.sha256(Path('01_exploration_dataset.ipynb').read_bytes()).hexdigest().upper() == receipt['NotebookHash']
(build / 'scope-math-validation.json').write_text(json.dumps({'passed': True, 'candidate_sha256': hashlib.sha256(candidate.read_bytes()).hexdigest()}, indent=2), encoding='utf-8')
print('Slides 1-9 and all original notes preserved. RAM and financial figures checked. Notebook/source unchanged.')
