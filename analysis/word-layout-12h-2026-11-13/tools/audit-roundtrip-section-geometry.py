"""Independent active OOXML section geometry audit of actual two-save files."""
from pathlib import Path
import zipfile,json,hashlib,os,datetime
from lxml import etree
B=Path(__file__).resolve().parents[1];R=B.parents[1]
NS={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'};W='{'+NS['w']+'}'
stage=os.environ.get('SECTION_RT_STAGE','government-6bc49fa')

def attributes(e):return {etree.QName(k).localname:v for k,v in e.attrib.items()} if e is not None else {}
def sections(p):
    with zipfile.ZipFile(p)as z:root=etree.fromstring(z.read('word/document.xml'))
    result=[]
    for e in root.xpath('//w:body//w:sectPr[not(ancestor::w:sectPrChange) and not(ancestor::w:pPrChange)]',namespaces=NS):
        size=attributes(e.find(W+'pgSz'));margins=attributes(e.find(W+'pgMar'));cols=attributes(e.find(W+'cols'));break_node=e.find(W+'type')
        normalized={'width':int(size.get('w',11906)),'height':int(size.get('h',16838)),'orientation':size.get('orient','portrait'),
            'breakType':break_node.get(W+'val','nextPage')if break_node is not None else 'nextPage',
            'margins':{k:int(margins.get(k,v))for k,v in dict(top=1440,right=1440,bottom=1440,left=1440,header=720,footer=720,gutter=0).items()},
            'columns':{'num':int(cols.get('num',1)),'space':int(cols.get('space',720)),'equalWidth':cols.get('equalWidth','1')not in ['0','false','off'],
                'separator':cols.get('sep','0')in ['1','true','on'],'explicit':[attributes(x)for x in e.findall(W+'cols/'+W+'col')]}}
        result.append({'normalized':normalized,'rawSize':size,'rawMargins':margins,'rawColumns':cols})
    return result

def differences(a,b,prefix=''):
    out=[]
    if isinstance(a,dict)and isinstance(b,dict):
        for k in set(a)|set(b):out.extend(differences(a.get(k),b.get(k),prefix+'.'+k))
    elif isinstance(a,list)and isinstance(b,list):
        if len(a)!=len(b):out.append({'field':prefix+'.length','before':len(a),'after':len(b)})
        for i,(x,y)in enumerate(zip(a,b)):out.extend(differences(x,y,prefix+f'[{i}]'))
    elif isinstance(a,int)and not isinstance(a,bool)and isinstance(b,int)and not isinstance(b,bool):
        # num/boolean are exact. Physical twip values tolerate only serializer rounding.
        tolerance=0 if prefix.endswith('.num') else 1
        if abs(a-b)>tolerance:out.append({'field':prefix,'before':a,'after':b,'toleranceTwips':tolerance})
    elif a!=b:out.append({'field':prefix,'before':a,'after':b})
    return out

rt=json.loads((B/('roundtrip-'+stage)/'results.json').read_text());rows=[]
for row in rt['rows']:
    if not row.get('files')or len(row['files'])!=2:continue
    paths=[Path(row['inputArtifact'])]+[B/f['file']for f in row['files']];states=[sections(p)for p in paths]
    original=[s['normalized']for s in states[0]];first=[s['normalized']for s in states[1]];second=[s['normalized']for s in states[2]]
    a=differences(original,first);b=differences(first,second)
    rows.append(dict(id=row['id'],inputSHA256=hashlib.sha256(paths[0].read_bytes()).hexdigest(),originalSectionCount=len(original),
        sourceToSave1Differences=a,save1ToSave2Differences=b,sourceGeometryPreserved=not a,twoSaveGeometryStable=not b,states=states))
result=dict(utc=datetime.datetime.now(datetime.timezone.utc).isoformat(),method='Active original package w:sectPr vs two actual saved DOCXs, independently parsed XML; source/default page dimensions, break type, margins/gutter/header/footer offsets and column properties. Physical numeric twips tolerate at most one serializer rounding twip; counts and flags exact. Explicit column raw strings preserved for inspection; no content or layout mutation.',sourceCommit=rt['sourceCommit'],stage=stage,rows=rows)
(B/('section-geometry-'+stage+'.json')).write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n')
print(json.dumps({'documents':len(rows),'sourcePreserved':sum(r['sourceGeometryPreserved']for r in rows),'twoSaveStable':sum(r['twoSaveGeometryStable']for r in rows),'differences':[{k:r[k]for k in ['id','sourceToSave1Differences','save1ToSave2Differences']}for r in rows if r['sourceToSave1Differences']or r['save1ToSave2Differences']]},ensure_ascii=False))
