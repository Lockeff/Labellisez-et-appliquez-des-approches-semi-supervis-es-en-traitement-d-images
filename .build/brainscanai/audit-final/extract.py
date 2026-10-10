import hashlib
import json
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
sys.stdout.reconfigure(encoding='utf-8')
build = Path('.build/brainscanai/audit-final')
sys.path.insert(0, str(Path('.build/brainscanai/audit-v1/pdf-tools')))
from pypdf import PdfReader
import fitz
paths = [Path('P7+DSML+FAE_editable.pdf'), Path('01_exploration_dataset.ipynb'), Path('output/presentations/BrainScanAI_presentation_V1_Final.pptx')]
(build / 'source-hashes.json').write_text(json.dumps({str(p):hashlib.sha256(p.read_bytes()).hexdigest() for p in paths},indent=2),encoding='utf-8')
reader = PdfReader(paths[0])
pdf = '\n\n'.join(f'PAGE {i+1}\n{page.extract_text()}' for i,page in enumerate(reader.pages))
(build / 'rubric.txt').write_text(pdf,encoding='utf-8')
doc = fitz.open(paths[0])
for i,page in enumerate(doc):
    page.get_pixmap(matrix=fitz.Matrix(1.6,1.6)).save(str(build / f'rubric-{i+1}.png'))
n = json.loads(paths[1].read_text(encoding='utf-8'))
chunks=[]
for i,c in enumerate(n['cells']):
    chunks.append(f"CELL {i+1} {c['cell_type']} id={c.get('id','')} execution={c.get('execution_count')}\n{''.join(c['source'])}")
    for o in c.get('outputs',[]):
        value=o.get('text',o.get('data',{}).get('text/plain',[]))
        if value: chunks.append('OUTPUT\n'+(''.join(value) if isinstance(value,list) else value))
        if o.get('output_type')=='error':chunks.append('ERROR\n'+o.get('ename','')+': '+o.get('evalue',''))
(build / 'notebook.txt').write_text('\n\n'.join(chunks),encoding='utf-8')
ns={'a':'http://schemas.openxmlformats.org/drawingml/2006/main','p':'http://schemas.openxmlformats.org/presentationml/2006/main'}
slides=[]
with zipfile.ZipFile(paths[2]) as z:
    files=sorted([f for f in z.namelist() if f.startswith('ppt/slides/slide') and f.endswith('.xml')],key=lambda f:int(f.split('slide')[-1].split('.')[0]))
    for f in files:
        i=int(f.split('slide')[-1].split('.')[0]);root=ET.fromstring(z.read(f))
        text='\n'.join(''.join(t.text or '' for t in p.findall('.//a:t',ns)) for p in root.findall('.//a:p',ns))
        notes=''
        np=f'ppt/notesSlides/notesSlide{i}.xml'
        if np in z.namelist():notes='\n'.join(t.text or '' for t in ET.fromstring(z.read(np)).findall('.//a:t',ns))
        slides.append({'slide':i,'text':text,'notes':notes})
(build / 'slides.json').write_text(json.dumps(slides,ensure_ascii=False,indent=2),encoding='utf-8')
print(pdf)
print('\nNotebook cells:',len(n['cells']),'PPT slides:',len(slides))
