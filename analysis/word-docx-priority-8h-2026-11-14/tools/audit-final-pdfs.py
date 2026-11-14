from pathlib import Path
import json,hashlib,logging,subprocess,concurrent.futures
from collections import Counter
import pdfplumber
from PIL import Image,ImageDraw
logging.getLogger('pdfminer').setLevel(logging.ERROR)
R=Path(__file__).resolve().parents[3];B=Path(__file__).resolve().parents[1]
baseline=json.loads((B/'baseline.json').read_text()); rows=[]
def inventory(file):
 chars=Counter();images=[];outside=[]
 with pdfplumber.open(file)as pdf:
  pages=len(pdf.pages)
  for page in pdf.pages:
   chars.update(c['text']for c in page.chars if c['text'].strip())
   for im in page.images:
    x,y,w,h=[float(im[k])for k in ['x0','top','width','height']]
    area=max(0,min(x+w,page.width)-max(x,0))*max(0,min(y+h,page.height)-max(y,0))
    images.append({'page':page.page_number,'sha256':hashlib.sha256(im['stream'].get_data()).hexdigest(),'bounds':[x,y,w,h]})
    if area<=0:outside.append(images[-1])
   page.close()
 return {'pages':pages,'chars':dict(chars),'images':images,'outside':outside}
for row in baseline['rows']:
 before=inventory(B/'baseline-native'/(row['id']+'.pdf'));after=inventory(B/'native-final-5c3bd56-original'/(row['id']+'.pdf'))
 item={'id':row['id'],'pages':[before['pages'],after['pages']],'sameGlyphKindMultiplicities':before['chars']==after['chars'],'sameImageBytesOrder':[x['sha256']for x in before['images']]==[x['sha256']for x in after['images']],'beforeImages':len(before['images']),'afterImages':len(after['images']),'beforeOutside':len(before['outside']),'afterOutside':len(after['outside'])};rows.append(item)
(B/'final-pdf-asset-audit.json').write_text(json.dumps({'method':'Independent actual PDF glyph-kind multiplicity and decoded raster operator bytes/order baseline to final. Operator presence is not a proof against all overlaps, vector losses, or correct page ownership; visual QA separate. Same source input SHA fixed.', 'rows':rows},indent=2));print(json.dumps(rows))
visual=B/'visual';visual.mkdir(exist_ok=True)
inputs=[('Word reference',R/baseline['rows'][6]['wordPDF']),('Studio before',B/'baseline-native/07-characters.pdf'),('Studio final',B/'native-final-5c3bd56-original/07-characters.pdf')]
def render(pair):
 i,(title,file)=pair;prefix=visual/('07-page2-'+str(i));subprocess.run(['pdftoppm','-f','2','-singlefile','-scale-to','1000','-png',str(file),str(prefix)],check=True,stdout=subprocess.DEVNULL,stderr=subprocess.PIPE);return title,Image.open(str(prefix)+'.png').convert('RGB')
with concurrent.futures.ThreadPoolExecutor(max_workers=3)as pool:images=list(pool.map(render,enumerate(inputs)))
w=sum(im.width for _,im in images);h=max(im.height for _,im in images)+30;panel=Image.new('RGB',(w,h),'#dddddd');d=ImageDraw.Draw(panel);x=0
for title,im in images:d.text((x+10,9),title,fill='black');panel.paste(im,(x,30));x+=im.width
panel.save(visual/'07-page2-final-comparison.png');print(visual/'07-page2-final-comparison.png')
