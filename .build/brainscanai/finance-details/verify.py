import hashlib
import json
import zipfile
from decimal import Decimal as D, ROUND_HALF_UP
from pathlib import Path
from xml.etree import ElementTree as ET

build = Path('.build/brainscanai/finance-details')
source = Path('output/presentations/BrainScanAI_presentation_V1.pptx')
candidate = build / 'candidate.pptx'
ns = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
url = 'https://docs.aws.amazon.com/awsaccountbilling/latest/aboutv2/using-the-aws-price-list-bulk-api-fetching-price-list-files-manually.html'
sha = lambda p: hashlib.sha256(p.read_bytes()).hexdigest()
with zipfile.ZipFile(source) as original, zipfile.ZipFile(candidate) as revised:
    for i in [*range(1, 11), 13]:
        for member in [f'ppt/slides/slide{i}.xml', f'ppt/notesSlides/notesSlide{i}.xml']:
            assert ET.tostring(ET.fromstring(original.read(member))) == ET.tostring(ET.fromstring(revised.read(member))), member
        png = f'slide-{i:02}.png'
        assert (build / 'source-renders' / png).read_bytes() == (build / 'candidate-renders' / png).read_bytes(), png
    for member in original.namelist():
        if member.startswith(('ppt/charts/', 'ppt/embeddings/')):
            assert original.read(member) == revised.read(member), member
    visible = {}
    for i in [11, 12]:
        visible[i] = ' '.join(t.text or '' for t in ET.fromstring(revised.read(f'ppt/slides/slide{i}.xml')).findall('.//a:t', ns))
        rels = ET.fromstring(revised.read(f'ppt/slides/_rels/slide{i}.xml.rels'))
        hyperlinks = [r.get('Target') for r in rels if r.get('Type', '').endswith('/hyperlink')]
        assert hyperlinks == [url], hyperlinks
        notes = ' '.join(t.text or '' for t in ET.fromstring(revised.read(f'ppt/notesSlides/notesSlide{i}.xml')).findall('.//a:t', ns))
        assert notes.count(url) == 1
    for text in ['72,55', '30 h × 2,2798 €/h = 68,39', '68,39 € + 4,15 € + 0,01 € = 72,55']:
        assert text in visible[11], text
    for text in ['67,03', '101,78', '8,2 h × 2,2798 €/h = 18,69', '12 mois × 4,148 € = 49,78', 'Dernière mesure : 9 h']:
        assert text in visible[12], text
round2 = lambda x: x.quantize(D('.01'), rounding=ROUND_HALF_UP)
assert round2(D('30') * D('2.2798')) == D('68.39')
assert round2(D('8.2') * D('2.2798')) == D('18.69')
assert round2(D('12') * D('4.148')) == D('49.78')
assert round2(D('12') * D('1.97815')) == D('23.74')
assert D('68.39') + D('4.15') + D('.01') == D('72.55')
assert D('18.69') + D('23.74') + D('4.15') + D('20.45') == D('67.03')
assert D('18.69') + D('12.86') + D('49.78') + D('20.45') == D('101.78')
receipt = json.loads((build / 'edit-receipt.json').read_text(encoding='utf-8-sig'))
assert sha(source).upper() == receipt['SourceHash']
assert sha(Path('01_exploration_dataset.ipynb')).upper() == receipt['NotebookHash']
(build / 'scope-math-validation.json').write_text(json.dumps({'passed': True, 'candidate_sha256': sha(candidate), 'unchanged_slides': [*range(1, 11), 13], 'checked_totals': ['72.55', '67.03', '101.78']}, indent=2), encoding='utf-8')
print('Scope, sums, source links, unchanged renders and notebook verified.')
