"""Add a native, editable scatter chart with the verified PCA coordinates."""
from pathlib import Path
from io import BytesIO
import hashlib
import json
import math
import zipfile
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parents[2]
BUILD = ROOT / '.build' / 'brainscanai'
DATA = json.loads((BUILD / 'sources.json').read_text(encoding='utf-8'))
SOURCE = BUILD / 'candidate-v3.pptx'
FINAL = ROOT / 'output/presentations/BrainScanAI_presentation_10_slides.pptx'
assert not FINAL.exists(), 'Use a new output filename for a revision.'

NS = {'c':'http://schemas.openxmlformats.org/drawingml/2006/chart',
      'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
      'p':'http://schemas.openxmlformats.org/presentationml/2006/main',
      'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      'pr':'http://schemas.openxmlformats.org/package/2006/relationships',
      'ct':'http://schemas.openxmlformats.org/package/2006/content-types',
      's':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
for prefix in ('a','c','p','r'):
    ET.register_namespace(prefix, NS[prefix])

def el(parent, tag, attrs=None, text=None):
    prefix, local = tag.split(':')
    node = ET.SubElement(parent, '{'+NS[prefix]+'}'+local, attrs or {})
    if text is not None: node.text = str(text)
    return node

def root(tag):
    prefix, local = tag.split(':')
    return ET.Element('{'+NS[prefix]+'}'+local)

def xml(node):
    if node.tag == '{'+NS['ct']+'}Types':
        ET.register_namespace('',NS['ct'])
    return ET.tostring(node, encoding='utf-8', xml_declaration=True)

def line_style(parent, color=None):
    props = el(parent, 'c:spPr')
    if color:
        el(el(props,'a:solidFill'),'a:srgbClr',{'val':color})
    ln = el(props, 'a:ln', {'w':'12700'})
    if color: el(el(ln,'a:solidFill'),'a:srgbClr',{'val':color})
    else: el(ln, 'a:noFill')
    return props

def rich(parent, tag='c:txPr', text=None, size=1300):
    tx=el(parent,tag)
    el(tx,'a:bodyPr')
    el(tx,'a:lstStyle')
    paragraph=el(tx,'a:p')
    pp=el(paragraph,'a:pPr')
    defaults=el(pp,'a:defRPr',{'sz':str(size)})
    el(el(defaults,'a:solidFill'),'a:srgbClr',{'val':'536473'})
    el(defaults,'a:latin',{'typeface':'Arial'})
    if text is not None:
        run=el(paragraph,'a:r')
        rp=el(run,'a:rPr',{'lang':'fr-FR','sz':str(size)})
        el(rp,'a:latin',{'typeface':'Arial'})
        el(run,'a:t',text=text)
    el(paragraph,'a:endParaRPr',{'lang':'fr-FR','sz':str(size)})
    return tx

# Embedded Excel workbook: one X/Y pair per group, complete source coordinates.
sheet=root('s:worksheet')
rows=el(sheet,'s:sheetData')
groups=DATA['kmeans']['groups']
headers=['Groupe 0 — PCA 1','Groupe 0 — PCA 2','Groupe 1 — PCA 1','Groupe 1 — PCA 2']
header=el(rows,'s:row',{'r':'1'})
for col,name in zip('ABCD',headers):
    cell=el(header,'s:c',{'r':col+'1','t':'inlineStr'})
    el(el(cell,'s:is'),'s:t',text=name)
for i in range(max(map(len,groups))):
    row=el(rows,'s:row',{'r':str(i+2)})
    for g,pair in enumerate(groups):
        if i < len(pair):
            for k,val in enumerate(pair[i]):
                cell=el(row,'s:c',{'r':'ABCD'[g*2+k]+str(i+2)})
                el(cell,'s:v',text=format(val,'.17g'))
book=root('s:workbook')
el(el(book,'s:sheets'),'s:sheet',{'name':'Sheet1','sheetId':'1','{'+NS['r']+'}id':'rId1'})
bookrels=root('pr:Relationships')
el(bookrels,'pr:Relationship',{'Id':'rId1','Type':NS['r']+'/worksheet','Target':'worksheets/sheet1.xml'})
pkgrel=root('pr:Relationships')
el(pkgrel,'pr:Relationship',{'Id':'rId1','Type':NS['r']+'/officeDocument','Target':'xl/workbook.xml'})
types=root('ct:Types')
el(types,'ct:Default',{'Extension':'rels','ContentType':'application/vnd.openxmlformats-package.relationships+xml'})
el(types,'ct:Default',{'Extension':'xml','ContentType':'application/xml'})
el(types,'ct:Override',{'PartName':'/xl/workbook.xml','ContentType':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'})
el(types,'ct:Override',{'PartName':'/xl/worksheets/sheet1.xml','ContentType':'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'})
workbook=BytesIO()
with zipfile.ZipFile(workbook,'w',zipfile.ZIP_DEFLATED) as z:
    for name,node in [('[Content_Types].xml',types),('_rels/.rels',pkgrel),('xl/workbook.xml',book),('xl/_rels/workbook.xml.rels',bookrels),('xl/worksheets/sheet1.xml',sheet)]: z.writestr(name,xml(node))

space=root('c:chartSpace')
el(space,'c:date1904',{'val':'0'})
el(space,'c:lang',{'val':'fr-FR'})
el(space,'c:roundedCorners',{'val':'0'})
chart=el(space,'c:chart')
el(chart,'c:autoTitleDeleted',{'val':'1'})
plot=el(chart,'c:plotArea')
el(plot,'c:layout')
scatter=el(plot,'c:scatterChart')
el(scatter,'c:scatterStyle',{'val':'marker'})
el(scatter,'c:varyColors',{'val':'0'})
for g,points in enumerate(groups):
    series=el(scatter,'c:ser')
    el(series,'c:idx',{'val':str(g)}); el(series,'c:order',{'val':str(g)})
    el(el(series,'c:tx'),'c:v',text='Groupe '+str(g))
    line_style(series)
    marker=el(series,'c:marker')
    el(marker,'c:symbol',{'val':'circle'}); el(marker,'c:size',{'val':'3'})
    color=['3978A9','D58A3D'][g]
    line_style(marker,color)
    for dim,tag in [(0,'c:xVal'),(1,'c:yVal')]:
        ref=el(el(series,tag),'c:numRef')
        col='ABCD'[g*2+dim]
        el(ref,'c:f',text=f'Sheet1!${col}$2:${col}${len(points)+1}')
        cache=el(ref,'c:numCache')
        el(cache,'c:formatCode',text='General')
        el(cache,'c:ptCount',{'val':str(len(points))})
        for i,point in enumerate(points): el(el(cache,'c:pt',{'idx':str(i)}),'c:v',text=format(point[dim],'.17g'))
    el(series,'c:smooth',{'val':'0'})
el(scatter,'c:axId',{'val':'101'}); el(scatter,'c:axId',{'val':'102'})
for dim,id,other,pos in [(0,101,102,'b'),(1,102,101,'l')]:
    vals=[point[dim] for group in groups for point in group]
    low=5*math.floor(min(vals)/5); high=5*math.ceil(max(vals)/5)
    axis=el(plot,'c:valAx')
    el(axis,'c:axId',{'val':str(id)})
    scaling=el(axis,'c:scaling')
    el(scaling,'c:orientation',{'val':'minMax'})
    el(scaling,'c:max',{'val':str(high)}); el(scaling,'c:min',{'val':str(low)})
    el(axis,'c:delete',{'val':'0'}); el(axis,'c:axPos',{'val':pos})
    title=el(axis,'c:title')
    rich(el(title,'c:tx'),'c:rich','PCA '+str(dim+1),1400)
    el(title,'c:layout'); el(title,'c:overlay',{'val':'0'})
    el(axis,'c:numFmt',{'formatCode':'0','sourceLinked':'0'})
    el(axis,'c:majorTickMark',{'val':'out'}); el(axis,'c:minorTickMark',{'val':'none'})
    el(axis,'c:tickLblPos',{'val':'nextTo'})
    props=el(axis,'c:spPr'); ln=el(props,'a:ln',{'w':'9525'})
    el(el(ln,'a:solidFill'),'a:srgbClr',{'val':'CAD7DF'})
    rich(axis,size=1200)
    el(axis,'c:crossAx',{'val':str(other)}); el(axis,'c:crosses',{'val':'min'})
    el(axis,'c:crossBetween',{'val':'midCat'})
legend=el(chart,'c:legend'); el(legend,'c:legendPos',{'val':'b'}); el(legend,'c:layout')
el(legend,'c:overlay',{'val':'0'}); rich(legend,size=1300)
el(chart,'c:plotVisOnly',{'val':'1'}); el(chart,'c:dispBlanksAs',{'val':'gap'})
el(chart,'c:showDLblsOverMax',{'val':'0'})
props=el(space,'c:spPr'); el(el(props,'a:solidFill'),'a:srgbClr',{'val':'FFFFFF'})
el(el(props,'a:ln'),'a:noFill')
rich(space,size=1300)
external=el(space,'c:externalData',{'{'+NS['r']+'}id':'rId1'})
el(external,'c:autoUpdate',{'val':'0'})
chartrels=root('pr:Relationships')
el(chartrels,'pr:Relationship',{'Id':'rId1','Type':NS['r']+'/package','Target':'../embeddings/Microsoft_Excel_Worksheet1.xlsx'})

with zipfile.ZipFile(SOURCE) as z:
    slide=ET.fromstring(z.read('ppt/slides/slide5.xml'))
    tree=slide.find('p:cSld/p:spTree',NS)
    target=next(sp for sp in tree.findall('p:sp',NS) if sp.find('p:nvSpPr/p:cNvPr',NS).get('name')=='KMeansChartPlaceholder')
    shapeid=target.find('p:nvSpPr/p:cNvPr',NS).get('id')
    coords=target.find('p:spPr/a:xfrm',NS)
    rels=ET.fromstring(z.read('ppt/slides/_rels/slide5.xml.rels'))
    rid='rId'+str(max(int(rel.get('Id')[3:]) for rel in rels)+1)
    frame=root('p:graphicFrame')
    nv=el(frame,'p:nvGraphicFramePr')
    el(nv,'p:cNvPr',{'id':shapeid,'name':'Clustering K-Means — graphique modifiable'})
    el(el(nv,'p:cNvGraphicFramePr'),'a:graphicFrameLocks',{'noGrp':'1'})
    el(nv,'p:nvPr')
    xf=el(frame,'p:xfrm')
    el(xf,'a:off',dict(coords.find('a:off',NS).attrib))
    el(xf,'a:ext',dict(coords.find('a:ext',NS).attrib))
    graphic=el(frame,'a:graphic')
    gd=el(graphic,'a:graphicData',{'uri':NS['c']})
    el(gd,'c:chart',{'{'+NS['r']+'}id':rid})
    index=list(tree).index(target); tree.remove(target); tree.insert(index,frame)
    el(rels,'pr:Relationship',{'Id':rid,'Type':NS['r']+'/chart','Target':'../charts/chart1.xml'})
    content=ET.fromstring(z.read('[Content_Types].xml'))
    if not any(n.get('Extension')=='xlsx' for n in content):
        el(content,'ct:Default',{'Extension':'xlsx','ContentType':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'})
    el(content,'ct:Override',{'PartName':'/ppt/charts/chart1.xml','ContentType':'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'})
    changed={'ppt/slides/slide5.xml':xml(slide),'ppt/slides/_rels/slide5.xml.rels':xml(rels),'[Content_Types].xml':xml(content)}
    FINAL.parent.mkdir(parents=True,exist_ok=True)
    with zipfile.ZipFile(FINAL,'w',zipfile.ZIP_DEFLATED) as out:
        for info in z.infolist(): out.writestr(info,changed.get(info.filename,z.read(info.filename)))
        out.writestr('ppt/charts/chart1.xml',xml(space))
        out.writestr('ppt/charts/_rels/chart1.xml.rels',xml(chartrels))
        out.writestr('ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx',workbook.getvalue())
assert hashlib.sha256(Path(DATA['source_notebook']).read_bytes()).hexdigest()==DATA['notebook_sha256']
(BUILD/'delivery.json').write_text(json.dumps({'path':str(FINAL),'slides':10,'native_chart_points':sum(map(len,groups)),'notebook_unchanged':True},ensure_ascii=False,indent=2),encoding='utf-8')
print(FINAL)
