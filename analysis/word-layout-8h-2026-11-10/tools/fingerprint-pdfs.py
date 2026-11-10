from pathlib import Path
import hashlib,json,re,logging,os
import pdfplumber
logging.getLogger('pdfminer').setLevel(logging.ERROR)
B=Path(__file__).resolve().parents[1];stage=os.environ.get('PDF_STAGE','final-product02cb');result=[]
def norm(v):
 if isinstance(v,float):return round(v,3)
 if isinstance(v,(tuple,list)):return[norm(x)for x in v]
 if isinstance(v,dict):return{k:norm(x)for k,x in v.items()}
 if v is None or isinstance(v,(int,str,bool)):return v
 return str(v)
def sha(v):return hashlib.sha256(json.dumps(norm(v),ensure_ascii=False,separators=(',',':')).encode()).hexdigest()
for item in json.loads((B/('render-'+stage)/'results.json').read_text())['rows']:
 file=B/('render-'+stage)/(item['id']+'.pdf');pages=[]
 with pdfplumber.open(file)as pdf:
  for p in pdf.pages:
   glyphs=[[c['text'],c['x0'],c['top'],c['x1'],c['bottom'],c['size'],re.sub(r'^[A-Z]{6}\+','',c['fontname']),c.get('non_stroking_color')]for c in p.chars]
   images=[[i['x0'],i['top'],i['width'],i['height'],i.get('srcsize'),i.get('bits')]for i in p.images]
   vectors=[{k:v.get(k)for k in ['object_type','x0','top','x1','bottom','pts','path','linewidth','fill','stroke','stroking_color','non_stroking_color']}for v in p.lines+p.rects+p.curves]
   pages.append({'size':[p.width,p.height],'rotation':p.rotation,'glyphs':len(glyphs),'glyphSHA256':sha(glyphs),'images':len(images),'imageGeometrySHA256':sha(images),'vectors':len(vectors),'vectorSHA256':sha(vectors)});p.close()
 result.append({'id':item['id'],'inputSHA256':item['inputSHA256'],'pdfSHA256':hashlib.sha256(file.read_bytes()).hexdigest(),'pages':pages});(B/('pdf-fingerprint-'+stage+'.json')).write_text(json.dumps({'method':'Studio PDF before/after optimization only, same input bytes. Actual page geometry plus glyph order/text/font/color/coordinates, raster image geometry and vector paths, rounded to .001pt. Not Microsoft fidelity score.','rows':result},ensure_ascii=False,indent=2));print(item['id'],len(pages),flush=True)
