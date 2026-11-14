from pathlib import Path
import json,subprocess,tempfile,shutil,concurrent.futures
from PIL import Image,ImageChops
B=Path(__file__).resolve().parents[1];tmp=Path(tempfile.mkdtemp(prefix='ws-word8-raster-',dir='/private/tmp'))
rows=[]
try:
 baseline=json.loads((B/'baseline.json').read_text())
 for row in baseline['rows']:
  def render(stage):
   folder=tmp/stage;folder.mkdir(exist_ok=True)
   subprocess.run(['pdftoppm','-r','96','-png',str(B/stage/(row['id']+'.pdf')),str(folder/row['id'])],check=True,stdout=subprocess.DEVNULL,stderr=subprocess.PIPE)
   return sorted(folder.glob(row['id']+'-*.png'))
  with concurrent.futures.ThreadPoolExecutor(max_workers=2)as pool:a,b=list(pool.map(render,['baseline-native','native-final-5c3bd56-original']))
  changes=[]
  for x,y in zip(a,b):
   before=Image.open(x).convert('RGB');after=Image.open(y).convert('RGB');diff=ImageChops.difference(before,after);box=diff.getbbox();hist=diff.convert('L').point(lambda n:255 if n else 0).histogram();changes.append({'page':len(changes)+1,'identical':box is None,'differenceBoundsPx':box,'differentPixels':hist[255]})
   before.close();after.close();x.unlink();y.unlink()
  item={'id':row['id'],'pages':len(a),'samePageCount':len(a)==len(b),'pagesWithIdenticalRaster':sum(x['identical']for x in changes),'changes':changes};rows.append(item);print({k:v for k,v in item.items()if k!='changes'},flush=True)
 (B/'final-raster-preservation.json').write_text(json.dumps({'method':'Actual Studio PDF raster at96DPI, baseline versus final same input. Independent of DOM/glyph matching, includes raster/vector/text ink. Pixel identity in unchanged documents proves unchanged rendering at this resolution; changed07/08 need separate visual review. Not a Word accuracy score.', 'rows':rows},indent=2))
finally:shutil.rmtree(tmp)
