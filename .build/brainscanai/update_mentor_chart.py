import io
import json
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
import numpy as np
import pandas as pd
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import StandardScaler
from sklearn.decomposition import PCA
from sklearn.cluster import KMeans
from sklearn.metrics import adjusted_rand_score

build = Path('.build/brainscanai/mentor-ppt')
features = pd.read_csv('resultats/etape2/features_resnet18.csv')
cols = [c for c in features.columns if c.startswith('feature_')]
strong = features[features.chemin.str.startswith('avec_labels/')].copy()
strong['label'] = strong.chemin.str.split('/').str[1]
train, test = train_test_split(strong, test_size=.2, stratify=strong.label, random_state=42)
train, val = train_test_split(train, test_size=.2, stratify=train.label, random_state=42)
weak, _ = train_test_split(features[features.chemin.str.startswith('sans_label/')], test_size=.2, random_state=42)
scaler = StandardScaler()
pca = PCA(n_components=2, random_state=42)
x = pca.fit_transform(scaler.fit_transform(weak[cols]))
kmeans = KMeans(n_clusters=2, random_state=42)
groups = kmeans.fit_predict(x)
ari_data = pd.concat([train, val])
ari = adjusted_rand_score(ari_data.label, kmeans.predict(pca.transform(scaler.transform(ari_data[cols]))))
assert np.bincount(groups).tolist() == [565, 483]
assert abs(ari - .3772887757615568) < 1e-12

with zipfile.ZipFile(build/'candidate-com.pptx') as z:
    contents = {n: z.read(n) for n in z.namelist()}
c = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
ns = {'c': c}
ET.register_namespace('c',c)
ET.register_namespace('a','http://schemas.openxmlformats.org/drawingml/2006/main')
ET.register_namespace('r','http://schemas.openxmlformats.org/officeDocument/2006/relationships')
root = ET.fromstring(contents['ppt/charts/chart1.xml'])
points = [x[groups == g] for g in range(2)]
for g, series in enumerate(root.findall('.//c:scatterChart/c:ser', ns)):
    for axis, tag in enumerate(('xVal', 'yVal')):
        ref = series.find(f'c:{tag}/c:numRef', ns)
        column = ('A','B','C','D')[g*2+axis]
        ref.find('c:f', ns).text = f'Sheet1!${column}$2:${column}${len(points[g])+1}'
        cache = ref.find('c:numCache', ns)
        for child in list(cache):
            cache.remove(child)
        ET.SubElement(cache, f'{{{c}}}formatCode').text = 'General'
        ET.SubElement(cache, f'{{{c}}}ptCount', val=str(len(points[g])))
        for i, v in enumerate(points[g][:, axis]):
            pt = ET.SubElement(cache, f'{{{c}}}pt', idx=str(i))
            ET.SubElement(pt, f'{{{c}}}v').text = format(v,'.17g')
contents['ppt/charts/chart1.xml'] = ET.tostring(root, xml_declaration=True, encoding='UTF-8')

relpath='ppt/charts/_rels/chart1.xml.rels'
rels=ET.fromstring(contents[relpath])
target=next(r.get('Target') for r in rels if r.get('Type').endswith('/package'))
import posixpath
wbpath=posixpath.normpath(posixpath.join('ppt/charts',target))
with zipfile.ZipFile(io.BytesIO(contents[wbpath])) as wb:
    wb_contents={n:wb.read(n) for n in wb.namelist()}
s='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
ET.register_namespace('',s)
sheet=ET.fromstring(wb_contents['xl/worksheets/sheet1.xml'])
data=sheet.find(f'{{{s}}}sheetData')
for row in list(data):
    if row.get('r') != '1': data.remove(row)
for i in range(max(map(len,points))):
    row=ET.SubElement(data,f'{{{s}}}row',r=str(i+2))
    for g in range(2):
        if i>=len(points[g]): continue
        for axis in range(2):
            col=('A','B','C','D')[g*2+axis]
            cell=ET.SubElement(row,f'{{{s}}}c',r=f'{col}{i+2}')
            ET.SubElement(cell,f'{{{s}}}v').text=format(points[g][i,axis],'.17g')
dimension=sheet.find(f'{{{s}}}dimension')
if dimension is not None: dimension.set('ref',f'A1:D{max(map(len,points))+1}')
wb_contents['xl/worksheets/sheet1.xml']=ET.tostring(sheet,xml_declaration=True,encoding='UTF-8')
stream=io.BytesIO()
with zipfile.ZipFile(stream,'w',zipfile.ZIP_DEFLATED) as wb:
    for n,b in wb_contents.items():wb.writestr(n,b)
contents[wbpath]=stream.getvalue()
with zipfile.ZipFile(build/'candidate.pptx','w',zipfile.ZIP_DEFLATED) as z:
    for n,b in contents.items():z.writestr(n,b)
(build/'chart-evidence.json').write_text(json.dumps({'counts':np.bincount(groups).tolist(),'ari':ari,'rows':len(weak)},indent=2))
print('Updated native chart and embedded workbook: 1048 points, ARI',ari)
