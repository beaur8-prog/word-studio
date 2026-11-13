"""Map actual Word PDF image operators to preserved source raster pixels, then compare page ownership.
RGBA images are composited on the same white paper before comparison. Exact pixels preferred; close decoded RGB pixels explicitly labeled (mean channel error below 1/255 on a 64-square thumbnail). Ambiguous identities remain unscored.
"""
from pathlib import Path
import base64,gzip,json,io,hashlib,logging,math,collections
from PIL import Image,ImageChops,ImageStat
from pypdf import PdfReader
import pdfplumber
logging.getLogger('pdfminer').setLevel(logging.ERROR)
R=Path(__file__).resolve().parents[3];B=Path(__file__).resolve().parents[1];P=R/'analysis/hwp-cd-10h-2026-11-07'
model=json.loads(gzip.decompress((B/'native-sop-77d0d06/pass-0-before.json.gz').read_bytes()));source=[]
def matte(im):
 a=im.convert('RGBA');paper=Image.new('RGBA',a.size,(255,255,255,255));paper.alpha_composite(a);return paper.convert('RGB')
def digest(im):return hashlib.sha256(im.convert('RGB').tobytes()).hexdigest()
def walk(n):
 if n.get('type')in['image','inlineImage']:
  a=n['attrs'];raw=base64.b64decode(a['src'].split(',')[1]);im=matte(Image.open(io.BytesIO(raw)));source.append({'index':len(source),'size':im.size,'pixels':digest(im),'thumb':im.resize((64,64)),'width':a['width'],'height':a['height'],'bytesSHA256':hashlib.sha256(raw).hexdigest()})
 for c in n.get('content',[]):walk(c)
walk(model)
file=P/'word-verified-returned/WS-CD-20261107-VERIFIED2-mois-sop-3.pdf';reader=PdfReader(file);word=[];unmatched=[]
with pdfplumber.open(file)as pdf:
 for pi,page in enumerate(pdf.pages):
  if not page.images:page.close();continue
  decoded={Path(item.name).stem:matte(item.image)for item in reader.pages[pi].images}
  for operator in page.images:
   name=operator['name'];im=decoded.get(name)
   if im is None:unmatched.append({'page':pi+1,'name':name,'reason':'No decoded image for painted operator'});continue
   candidates=[s for s in source if s['size']==im.size];exact=[s for s in candidates if s['pixels']==digest(im)]
   mode='exact-decoded-RGB';score=0
   if not exact:
    ranked=sorted([(sum(ImageStat.Stat(ImageChops.difference(s['thumb'],im.resize((64,64)))).mean)/3,s['index'],s)for s in candidates],key=lambda v:v[0]);exact=[s for v,_,s in ranked if v<1];mode='near-decoded-RGB';score=ranked[0][0]if ranked else None
   if len(exact)!=1:unmatched.append({'page':pi+1,'name':name,'size':im.size,'reason':'Ambiguous or missing pixel identity','candidateIndices':[s['index']for s in exact],'bestMeanChannelError':score});continue
   s=exact[0];word.append({'sourceIndex':s['index'],'page':pi+1,'name':name,'match':mode,'meanChannelError':score,'x':operator['x0']*4/3,'y':operator['top']*4/3,'width':operator['width']*4/3,'height':operator['height']*4/3})
  page.close()
sets={}
for label,file in [('product-77d0d06',B/'first-drift-sop-cell-section-fair-v1.json.gz'),('empty-prefix-candidate',B/'first-drift-sop-empty-prefix-move-v1.json.gz'),('own-marker-v3',B/'first-drift-sop-empty-own-fit-v3.json.gz'),('blank-lines-v4',B/'first-drift-sop-empty-line-wrap-v4.json.gz')]:
 d=json.loads(gzip.decompress(file.read_bytes()));drawings=d.get('selectedDrawings');rows=[]
 if drawings is None:continue
 for w in word:
  a=drawings[w['sourceIndex']];page=math.floor((a['y']+.5)/(1122.52+24))+1;actual={'page':page,'x':a['x'],'y':a['y']-(page-1)*(1122.52+24),'width':a['width'],'height':a['height']};rows.append({'sourceIndex':w['sourceIndex'],'word':w,'editor':actual,'pageDifference':page-w['page'],'hidden':a['width']==0 or a['height']==0})
 sets[label]={'pages':d['editorPages'],'rows':rows,'mismatched':sum(r['pageDifference']!=0 for r in rows),'absolutePageError':sum(abs(r['pageDifference'])for r in rows),'hiddenMatchedPictures':sum(r['hidden']for r in rows)}
out={'method':__doc__,'wordPDFSHA256':hashlib.sha256((P/'word-verified-returned/WS-CD-20261107-VERIFIED2-mois-sop-3.pdf').read_bytes()).hexdigest(),'sourcePictures':len(source),'paintedMatched':len(word),'unmatchedOrAmbiguous':unmatched,'comparisons':sets}
(B/'sop-word-picture-pages-rgba25.json').write_text(json.dumps(out,ensure_ascii=False,indent=2));print(json.dumps({k:v for k,v in out.items()if k!='comparisons'},ensure_ascii=False));print(json.dumps({k:{n:v[n]for n in ['pages','mismatched','absolutePageError','hiddenMatchedPictures']}for k,v in sets.items()}));print(json.dumps({k:[(r['sourceIndex'],r['word']['page'],r['editor']['page'])for r in v['rows']]for k,v in sets.items()}))
