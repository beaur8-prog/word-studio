"""Read-only source-package minimization for the next actual Word comparison.
The original target table XML is preserved in variant A. B/C/D are explicit
counterfactual diagnostics, never product changes or the fixed original cohort.
"""
from pathlib import Path
from copy import deepcopy
import json,zipfile,hashlib,posixpath
from lxml import etree as E
R=Path(__file__).resolve().parents[3];B=Path(__file__).resolve().parents[1];P=R/'analysis/hwp-cd-10h-2026-11-07';O=B/'minimal-crossed-merge';O.mkdir(exist_ok=True)
row=next(r for r in json.loads((P/'final-word-studio-comparison.json').read_text())['rows']if r['id']=='mss-funding-2026-2')
source=P/row['inputFile'];raw=source.read_bytes()
if hashlib.sha256(raw).hexdigest()!=row['inputSHA256']:raise ValueError('Original bytes changed')
w='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}';rel='{http://schemas.openxmlformats.org/package/2006/relationships}';ct='{http://schemas.openxmlformats.org/package/2006/content-types}'
with zipfile.ZipFile(source)as z:
 parts={n:z.read(n)for n in z.namelist()};original=E.fromstring(parts['word/document.xml'])
 table=next(t for t in original.iter(w+'tbl')if '로봇생산자동화기술'in ''.join(t.itertext()));settings=deepcopy(original.find(w+'body').find(w+'sectPr'))
 # There are no image drawings in the selected exact source table. Remove
 # body-image relationships and only assets with no other remaining reference.
 if list(table.iter('{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}anchor')):raise ValueError('Target must be a text-only table')
 relationships=E.fromstring(parts['word/_rels/document.xml.rels']);deleted_targets=set()
 for q in list(relationships):
  if q.get('Type','').endswith('/image'):
   deleted_targets.add(posixpath.normpath(posixpath.join('word',q.get('Target'))));relationships.remove(q)
 parts['word/_rels/document.xml.rels']=E.tostring(relationships,xml_declaration=True,encoding='UTF-8',standalone=True)
 other_targets=set()
 for name,data in parts.items():
  if name.endswith('.rels')and name!='word/_rels/document.xml.rels':
   base=posixpath.dirname(posixpath.dirname(name))
   for q in E.fromstring(data):
    if q.get('TargetMode')!='External':other_targets.add(posixpath.normpath(posixpath.join(base,q.get('Target',''))))
 for target in deleted_targets-other_targets:parts.pop(target,None)
 types=E.fromstring(parts['[Content_Types].xml'])
 for q in list(types):
  if q.tag==ct+'Override'and q.get('PartName','').lstrip('/')not in parts:types.remove(q)
 parts['[Content_Types].xml']=E.tostring(types,xml_declaration=True,encoding='UTF-8',standalone=True)
 target_text=''.join(table.itertext());target_xml=E.tostring(table,method='c14n')
 reports=[]
 for label in ['A-exact-source','B-no-row-minima','C-explicit-tiny-continuation-marks','D-no-vertical-merges']:
  root=E.Element(original.tag,nsmap=original.nsmap,attrib=dict(original.attrib));body=E.SubElement(root,w+'body');copy=deepcopy(table);body.append(copy);body.append(deepcopy(settings))
  if label=='B-no-row-minima':
   for tr in copy.iter(w+'tr'):
    props=tr.find(w+'trPr')
    if props is not None:
     for h in props.findall(w+'trHeight'):props.remove(h)
  if label=='C-explicit-tiny-continuation-marks':
   for tc in copy.iter(w+'tc'):
    merge=tc.find(w+'tcPr/'+w+'vMerge')
    if merge is None or merge.get(w+'val')=='restart':continue
    for p in tc.findall(w+'p'):
     pp=p.find(w+'pPr')
     if pp is None:pp=E.Element(w+'pPr');p.insert(0,pp)
     rp=pp.find(w+'rPr')
     if rp is None:rp=E.SubElement(pp,w+'rPr')
     for kind in ['sz','szCs']:
      size=rp.find(w+kind)
      if size is None:size=E.SubElement(rp,w+kind)
      size.set(w+'val','2')
  if label=='D-no-vertical-merges':
   for q in list(copy.iter(w+'vMerge')):q.getparent().remove(q)
  if ''.join(copy.itertext())!=target_text:raise ValueError('Diagnostic changed text')
  if label=='A-exact-source'and E.tostring(copy,method='c14n')!=target_xml:raise ValueError('Exact original table XML changed')
  parts['word/document.xml']=E.tostring(root,xml_declaration=True,encoding='UTF-8',standalone=True)
  file=O/('textbox-save-crossed-merge-'+label+'.docx')
  with zipfile.ZipFile(file,'w',zipfile.ZIP_DEFLATED)as out:
   for name,data in parts.items():out.writestr(name,data)
  reports.append({'variant':label,'file':str(file.relative_to(B)),'bytes':file.stat().st_size,'sha256':hashlib.sha256(file.read_bytes()).hexdigest(),'tableTextSHA256':hashlib.sha256(target_text.encode()).hexdigest(),'rowCount':len(copy.findall(w+'tr')),'originalTableXMLPreserved':label=='A-exact-source','MicrosoftWordMeasured':False})
report={'method':__doc__,'sourceInputSHA256':row['inputSHA256'],'sourceId':row['id'],'sourceTableIndex':22,'sourceRows':len(table.findall(w+'tr')),'sourceTableText':target_text,'sourceTableXMLSHA256':hashlib.sha256(target_xml).hexdigest(),'rows':reports,'notAnAccuracyResult':'Each smaller file is a different input from the fixed 29-document cohort. No new Microsoft measurement has been made.'}
(O/'manifest.json').write_text(json.dumps(report,ensure_ascii=False,indent=2));print(json.dumps({'rows':reports},ensure_ascii=False))
