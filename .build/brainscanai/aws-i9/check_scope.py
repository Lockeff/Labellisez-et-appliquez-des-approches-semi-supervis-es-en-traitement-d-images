import hashlib
import json
import zipfile
from decimal import Decimal
from pathlib import Path
from xml.etree import ElementTree as ET

build=Path('.build/brainscanai/aws-i9')
source=Path('output/presentations/BrainScanAI_presentation_13_slides_actualisee.pptx')
candidate=build/'candidate.pptx'
ns={'a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
with zipfile.ZipFile(source) as z1,zipfile.ZipFile(candidate) as z2:
    for i in range(1,10):
        p=f'ppt/slides/slide{i}.xml'
        # Native PowerPoint may normalize XML without changing objects.
        r1=ET.fromstring(z1.read(p));r2=ET.fromstring(z2.read(p))
        assert ET.tostring(r1)==ET.tostring(r2),f'Slide {i} changed'
    for i in range(10,14):
        text=' '.join(t.text or '' for t in ET.fromstring(z2.read(f'ppt/slides/slide{i}.xml')).findall('.//a:t',ns))
        assert 'c6i.' not in text and '8 Gio' not in text
    for n in z1.namelist():
        if n.startswith('ppt/charts/') or n.startswith('ppt/embeddings/'):
            assert z1.read(n)==z2.read(n),f'Existing chart or workbook changed: {n}'
for i in range(1,10):
    p=f'slide-{i:02}.png'
    assert (build/'source-renders'/p).read_bytes()==(build/'candidate-renders'/p).read_bytes(),f'Slide {i} pixels changed'
receipt=json.loads((build/'edit-receipt.json').read_text(encoding='utf-8-sig'))
assert hashlib.sha256(Path('01_exploration_dataset.ipynb').read_bytes()).hexdigest().upper()==receipt['NotebookHash']
assert hashlib.sha256(source.read_bytes()).hexdigest().upper()==receipt['SourceHash']
print('First nine slides, existing chart/workbook, source file and notebook preserved.')
