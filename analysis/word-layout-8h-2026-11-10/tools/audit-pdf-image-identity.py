"""Read-only raster identity diagnostic, independent of Studio's imported model.

Compare public DOCX bitmap payloads and actual PDF XObject decoded pixels. A
match establishes presence, not correct page ownership, clipping or overlap.
A mismatch may be PDF color/recompression and is not automatically omission.
"""
from pathlib import Path
from collections import Counter
import hashlib,io,json,logging,os,posixpath,zipfile,xml.etree.ElementTree as E
from PIL import Image
from pypdf import PdfReader
logging.getLogger('pypdf').setLevel(logging.ERROR)
R=Path(__file__).resolve().parents[3];B=Path(__file__).resolve().parents[1];P=R/'analysis/hwp-cd-10h-2026-11-07'
selected=os.environ.get('IMAGE_IDS','mois-budget-0,mois-sop-3,nps-venture-hwp').split(',')
ns={'a':'http://schemas.openxmlformats.org/drawingml/2006/main','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
def fingerprint(im):
 if im.width*im.height>40000000:return None
 rgba=im.convert('RGBA');background=Image.new('RGBA',im.size,'white');background.alpha_composite(rgba);rgb=background.convert('RGB')
 return str(im.size)+':'+hashlib.sha256(rgb.tobytes()).hexdigest()
rows=[]
for row in json.loads((P/'final-word-studio-comparison.json').read_text())['rows']:
 if 'all'not in selected and row['id']not in selected:continue
 source=[];pdf=[];errors=[]
 with zipfile.ZipFile(P/row['inputFile'])as z:
  rel={x.attrib['Id']:x.attrib['Target']for x in E.fromstring(z.read('word/_rels/document.xml.rels'))}
  for index,el in enumerate(E.fromstring(z.read('word/document.xml')).findall('.//a:blip',ns)):
   target=rel[el.attrib['{'+ns['r']+'}embed']];name=posixpath.normpath(posixpath.join('word',target))if not target.startswith('/')else target.lstrip('/');data=z.read(name)
   try:
    with Image.open(io.BytesIO(data))as im:pixel=fingerprint(im)
    source.append({'index':index,'part':name,'sha256':hashlib.sha256(data).hexdigest(),'pixels':pixel})
   except Exception as exc:errors.append({'sourceIndex':index,'error':str(exc)})
 reader=PdfReader(B/'render-final-figure-product'/(row['id']+'.pdf'))
 for page,pg in enumerate(reader.pages):
  for item in pg.images:
   try:pdf.append({'page':page+1,'name':item.name,'pixels':fingerprint(item.image)})
   except Exception as exc:errors.append({'page':page+1,'name':item.name,'error':str(exc)})
  pg.clear_cache()if hasattr(pg,'clear_cache')else None
 counts=Counter(x['pixels']for x in pdf if x['pixels']);unmatched=[x for x in source if not x['pixels']or not counts[x['pixels']]]
 result={'id':row['id'],'sourceBodyImagePlacements':len(source),'PDFImageResources':len(pdf),'sourcePlacementsWhoseDecodedPixelsExistInPDF':len(source)-len(unmatched),'unmatchedSourcePlacements':unmatched,'errors':errors,'source':source,'pdf':pdf};rows.append(result)
 (B/'pdf-image-identity-final.json').write_text(json.dumps({'method':__doc__,'rows':rows},indent=2));print({k:v for k,v in result.items()if k not in ['source','pdf']},flush=True)
