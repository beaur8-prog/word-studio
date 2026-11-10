"""Investigate failed exact-pixel JPEG matches without classifying them as loss."""
from pathlib import Path
import json,io,hashlib,zipfile,logging
from PIL import Image,ImageChops,ImageStat
from pypdf import PdfReader
logging.getLogger('pypdf').setLevel(logging.ERROR)
R=Path(__file__).resolve().parents[3];B=Path(__file__).resolve().parents[1];P=R/'analysis/hwp-cd-10h-2026-11-07'
inputs={x['id']:P/x['inputFile']for x in json.loads((P/'final-word-studio-comparison.json').read_text())['rows']};rows=[]
for row in json.loads((B/'pdf-image-identity-final.json').read_text())['rows']:
 if row['id']=='mois-sop-3':continue
 reader=PdfReader(B/'render-final-figure-product'/(row['id']+'.pdf'));candidates=[]
 for page,pg in enumerate(reader.pages):
  for item in pg.images:
   im=item.image
   candidates.append((page+1,item.name,im.convert('RGB'),hashlib.sha256(item.data).hexdigest()))
 with zipfile.ZipFile(inputs[row['id']])as z:
  for src in row['unmatchedSourcePlacements']:
   data=z.read(src['part']);im=Image.open(io.BytesIO(data)).convert('RGB');scores=[]
   for page,name,pdf,encodedSHA in candidates:
    if pdf.size!=im.size:continue
    a=im.copy();a.thumbnail((256,256));c=pdf.copy();c.thumbnail((256,256));stat=ImageStat.Stat(ImageChops.difference(a,c));scores.append((sum(stat.mean)/3,page,name,pdf,encodedSHA))
   if not scores:continue
   _,page,name,pdf,encodedSHA=min(scores,key=lambda x:x[0]);stat=ImageStat.Stat(ImageChops.difference(im,pdf));result={'id':row['id'],'sourceIndex':src['index'],'part':src['part'],'pixels':im.size,'candidatePage':page,'candidateImage':name,'encodedJPEGBytesIdentical':encodedSHA==hashlib.sha256(data).hexdigest(),'meanAbsoluteChannelDifference':sum(stat.mean)/3,'rootMeanSquareChannelDifference':sum(stat.rms)/3,'maximumChannelDifference':max(x[1]for x in stat.extrema)};rows.append(result);print(result,flush=True)
(B/'jpeg-pdf-color-diagnostic.json').write_text(json.dumps({'method':__doc__,'comparison':'Same raster pixel dimensions. Choose nearest downsampled candidate then compare all RGB pixels. Presence evidence only; not source layout, not a losslessness claim.','rows':rows},indent=2))
