from pathlib import Path
import json,gzip
B=Path(__file__).resolve().parents[1];manifest=json.loads((B/'baseline.json').read_text());rows=[]
for row in manifest['rows']:
 original=json.loads(gzip.decompress((B/'native-final-5c3bd56-original'/(row['id']+'.json.gz')).read_bytes()));saves=[]
 for save in [1,2]:
  current=json.loads(gzip.decompress((B/('native-final-5c3bd56-save'+str(save))/(row['id']+'.json.gz')).read_bytes()));modes={}
  for mode in ['editor','print']:
   a=original[mode];b=current[mode];sameText=''.join(c['text']for c in a['glyphs'])==''.join(c['text']for c in b['glyphs']);differentPages=0;maxX=maxY=0;changedGlyphs=0
   if sameText:
    for x,y in zip(a['glyphs'],b['glyphs']):
     differentPages+=x['page']!=y['page'];dx=abs(x['x']-y['x']);dy=abs(x['y']-y['y']);maxX=max(maxX,dx);maxY=max(maxY,dy);changedGlyphs+=dx>.1 or dy>.1
   modes[mode]={'pagesBefore':a['pages'],'pagesAfter':b['pages'],'sameGlyphTextOrder':sameText,'glyphPageChanges':differentPages,'maxXCSSpx':maxX,'maxYCSSpx':maxY,'glyphsMovedOverPointOneCSSpx':changedGlyphs}
  saves.append({'save':save,'modes':modes})
 rows.append({'id':row['id'],'saves':saves});print(row['id'],[(x['save'],x['modes']['print'])for x in saves],flush=True)
(B/'native-roundtrip-coordinate-audit.json').write_text(json.dumps({'method':'Same compiled Studio source5c3bd56, original Word10 versus two actual production DOCX export/reopen inputs from b5744c1, independent native glyph coordinates. This is Studio save stability, not a Microsoft accuracy claim for changed input bytes.', 'rows':rows},indent=2))
