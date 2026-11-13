"""Same-input historical Word versus current Studio actual PDF picture operators.
Match decoded RGB on white paper. Score only a unique pixel identity in both files;
near RGB (thumbnail channel error below 1/255) is explicitly labeled, unresolved or
repeated image identities stay unscored. This diagnoses page/size ownership, not OCR.
"""
from pathlib import Path
import os,json,hashlib,logging,collections
from PIL import Image,ImageChops,ImageStat
from pypdf import PdfReader
import pdfplumber
logging.getLogger('pdfminer').setLevel(logging.ERROR)
R=Path(__file__).resolve().parents[3];B=Path(__file__).resolve().parents[1];P=R/'analysis/hwp-cd-10h-2026-11-07'
id=os.environ.get('PICTURE_ID','mss-funding-2026-2');stage=os.environ.get('PICTURE_STAGE','final-1012ae8-government');row=next(r for r in json.loads((P/'final-word-studio-comparison.json').read_text())['rows']if r['id']==id)
def matte(im):
 a=im.convert('RGBA');paper=Image.new('RGBA',a.size,(255,255,255,255));paper.alpha_composite(a);return paper.convert('RGB')
def read(file):
 reader=PdfReader(file);out=[]
 with pdfplumber.open(file)as pdf:
  count=len(pdf.pages)
  for pi,page in enumerate(pdf.pages):
   decoded={Path(i.name).stem:matte(i.image)for i in reader.pages[pi].images}if page.images else {}
   for op in page.images:
    im=decoded.get(op['name']);bounds=[float(op[k])for k in ['x0','top','width','height']]
    if im is None:raise ValueError('Undecodable actual PDF image')
    x,y,w,h=bounds;visible=max(0,min(x+w,page.width)-max(x,0))*max(0,min(y+h,page.height)-max(y,0))
    out.append({'index':len(out),'page':pi+1,'name':op['name'],'pixels':hashlib.sha256(im.tobytes()).hexdigest(),'size':im.size,'thumb':im.resize((64,64)),'bounds':bounds,'visibleFraction':visible/(w*h)if w*h>0 else 0})
   page.close()
 return count,out
word=P/row['wordPDF'];studio=B/('render-'+stage)/(id+'.pdf')
if hashlib.sha256(word.read_bytes()).hexdigest()!=row['wordPDFSHA256']:raise ValueError('Historical Word reference changed')
wpages,words=read(word);spages,studios=read(studio);matched=[];unresolved=[]
for a in words:
 candidates=[b for b in studios if a['pixels']==b['pixels']and a['size']==b['size']];mode='exact-decoded-RGB';error=0
 if not candidates:
  ranked=sorted([(sum(ImageStat.Stat(ImageChops.difference(a['thumb'],b['thumb'])).mean)/3,b['index'],b)for b in studios if b['size']==a['size']],key=lambda q:q[0]);candidates=[b for e,_,b in ranked if e<1];mode='near-decoded-RGB';error=ranked[0][0]if ranked else None
 uniqueWord=sum(q['size']==a['size']and q['pixels']==a['pixels']for q in words)==1
 if len(candidates)!=1 or not uniqueWord:
  unresolved.append({'wordIndex':a['index'],'wordPage':a['page'],'sourcePixels':a['size'],'reason':'missing or repeated pixel identity','candidateIndices':[b['index']for b in candidates],'uniqueWord':uniqueWord,'bestMeanChannelError':error});continue
 b=candidates[0];matched.append({'wordIndex':a['index'],'studioIndex':b['index'],'pixelsSHA256':a['pixels'],'match':mode,'meanChannelError':error,'sourcePixels':a['size'],'word':{k:a[k]for k in ['page','bounds','visibleFraction']},'studio':{k:b[k]for k in ['page','bounds','visibleFraction']},'pageDifference':b['page']-a['page'],'widthDifferencePt':b['bounds'][2]-a['bounds'][2],'heightDifferencePt':b['bounds'][3]-a['bounds'][3]})
result={'method':__doc__,'id':id,'inputSHA256':row['inputSHA256'],'wordPDFSHA256':row['wordPDFSHA256'],'studioPDFSHA256':hashlib.sha256(studio.read_bytes()).hexdigest(),'wordPages':wpages,'studioPages':spages,'wordImageOperators':len(words),'studioImageOperators':len(studios),'uniquePixelMatches':len(matched),'pageMismatches':sum(q['pageDifference']!=0 for q in matched),'absolutePageError':sum(abs(q['pageDifference'])for q in matched),'matches':matched,'unresolved':unresolved}
(B/('word-studio-pictures-'+id+'-'+stage+'.json')).write_text(json.dumps(result,ensure_ascii=False,indent=2));print(json.dumps({k:v for k,v in result.items()if k not in ['method','matches','unresolved']},ensure_ascii=False));print(json.dumps({'unresolved':unresolved,'matches':matched},ensure_ascii=False))
