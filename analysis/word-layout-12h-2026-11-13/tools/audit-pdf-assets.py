from pathlib import Path
import json,logging,hashlib,os
import pdfplumber
logging.getLogger('pdfminer').setLevel(logging.ERROR)
B=Path(__file__).resolve().parents[1]
stage=os.environ.get('PDF_STAGE','final-product02cb')
manifest=json.loads((B/('render-'+stage)/'results.json').read_text());rows=[]
for row in manifest['rows']:
 file=B/('render-'+stage)/(row['id']+'.pdf');images=[];outliers=[];chars=0;from collections import Counter;glyphs=Counter()
 with pdfplumber.open(file)as pdf:
  pages=len(pdf.pages)
  for i,page in enumerate(pdf.pages):
   chars+=sum(bool(c['text'].strip())for c in page.chars);glyphs.update(c['text']for c in page.chars if c['text'].strip())
   for n,img in enumerate(page.images):
    box=[float(img[k])for k in ['x0','top','width','height']];left,top,width,height=box
    visible=max(0,min(left+width,page.width)-max(left,0))*max(0,min(top+height,page.height)-max(top,0))
    item={'page':i+1,'bounds':box,'visibleOnPaperArea':visible,'sourcePixels':[img.get('srcsize')],'bits':img.get('bits')}
    images.append(item)
    if visible<=0 or width<=0 or height<=0:outliers.append(item)
   page.close()
 out={'id':row['id'],'inputSHA256':row['inputSHA256'],'pdfSHA256':hashlib.sha256(file.read_bytes()).hexdigest(),'pages':pages,'sourceDOMImagePlacements':row['imagePlacements'],'pdfImageOperators':len(images),'printablePDFGlyphs':chars,'whollyOutsidePaperImages':outliers,'imagePositions':images,'glyphCounts':dict(glyphs)}
 rows.append(out);(B/('pdf-assets-'+stage+'.json')).write_text(json.dumps({'method':'Independent actual PDF operators and positive paper intersections. This proves image operators are present on paper; vector SVG/text pictures, masks, overlap, correct page ownership and visible ink require separate visual/pixel comparison. No substitution for raw source/image-byte preservation.','rows':rows},ensure_ascii=False,indent=2));print(json.dumps({k:v for k,v in out.items()if k not in ['imagePositions','glyphCounts']}),flush=True)
