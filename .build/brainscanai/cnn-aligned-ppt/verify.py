import hashlib
import json
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
build = Path('.build/brainscanai/cnn-aligned-ppt')
source = Path('output/presentations/BrainScanAI_presentation_V1_Final.pptx')
candidate = build / 'candidate.pptx'
ns = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
facts = json.loads((build / 'facts.json').read_text(encoding='utf-8'))
def tables(xml):
    return [[[''.join(t.text or '' for t in c.findall('.//a:t', ns)) for c in row.findall('a:tc', ns)] for row in tbl.findall('a:tr', ns)] for tbl in ET.fromstring(xml).findall('.//a:tbl', ns)]
def txt(xml):
    return ' '.join(t.text or '' for t in ET.fromstring(xml).findall('.//a:t', ns))
unchanged = [1,2,3,4,5,6,10,11,12,13]
with zipfile.ZipFile(source) as old, zipfile.ZipFile(candidate) as new:
    for i in unchanged:
        for member in [f'ppt/slides/slide{i}.xml', f'ppt/notesSlides/notesSlide{i}.xml']:
            assert ET.tostring(ET.fromstring(old.read(member))) == ET.tostring(ET.fromstring(new.read(member))), member
        png = f'slide-{i:02}.png'
        assert (build/'source-renders'/png).read_bytes() == (build/'candidate-renders'/png).read_bytes(), png
    for member in old.namelist():
        if member.startswith(('ppt/charts/', 'ppt/embeddings/')):
            assert old.read(member) == new.read(member), member
    tbls = tables(new.read('ppt/slides/slide7.xml'))
    grid = next(t for t in tbls if t[0][0] == 'Taux')
    for row, semi, sup in zip(grid[1:], facts['semi_grid'], facts['super_grid']):
        assert row[2] == f"{float(semi['F2 cancer validation']):.4f}".replace('.', ',')
        assert row[3] == f"{float(sup['F2 cancer validation']):.4f}".replace('.', ',')
    tbls = tables(new.read('ppt/slides/slide9.xml'))
    assert len(tbls) == 3
    metrics = next(t for t in tbls if t[0][0] == 'Métrique')
    assert all(row[1:] == ['0,900', '0,900'] for row in metrics[1:])
    for matrix in [t for t in tbls if t[0][0] == 'Vrai']:
        assert matrix[1][1:] == ['9', '1'] and matrix[2][1:] == ['1', '9']
        tn, fp, fn, tp = [int(matrix[r][c]) for r in [1,2] for c in [1,2]]
        assert (tn+tp)/(tn+fp+fn+tp) == 5*tp/(5*tp+4*fn+fp) == .9
    text8 = txt(new.read('ppt/slides/slide8.xml'))
    text9 = txt(new.read('ppt/slides/slide9.xml'))
    assert '0,973' in text8 and '0,977' not in text8
    assert '0 = normal et 1 = cancer' in text8
    for claim in ['mêmes performances', "n'apporte pas d'amélioration mesurée", '+6,9 points', '+4,3 points']:
        assert claim in text9
    assert 'meilleur F2' not in text9 and '2 faux négatifs' not in text9
hashes = json.loads((build / 'source-hashes.json').read_text(encoding='utf-8'))
assert all(hashlib.sha256(Path(p).read_bytes()).hexdigest() == h for p,h in hashes.items())
(build/'scope-math-validation.json').write_text(json.dumps({'passed': True, 'candidate_sha256': hashlib.sha256(candidate.read_bytes()).hexdigest(), 'changed_slides':[7,8,9], 'unchanged_slides':unchanged}, indent=2), encoding='utf-8')
print('Saved notebook scores, matrices, conclusions and unchanged slides verified.')
