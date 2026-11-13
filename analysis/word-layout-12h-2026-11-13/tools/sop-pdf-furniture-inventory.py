from pathlib import Path
from collections import Counter
import json,logging,os
import pdfplumber
logging.getLogger('pdfminer').setLevel(logging.ERROR)
B=Path(__file__).resolve().parents[1];rows=[]
stages=['page-segment-checkpoint',os.environ.get('SOP_PDF_FINAL_STAGE','final-b6edda3-government')]
for stage in stages:
 all_glyphs=Counter();furniture=Counter();body=Counter()
 with pdfplumber.open(B/('render-'+stage)/'mois-sop-3.pdf') as pdf:
  for p in pdf.pages:
   for g in p.chars:
    if not g['text'].strip():continue
    all_glyphs[g['text']]+=1
    (furniture if g['top']<72 or g['top']>790 else body)[g['text']]+=1
   p.close()
  rows.append({'stage':stage,'pages':len(pdf.pages),'all':dict(all_glyphs),'furniture':dict(furniture),'body':dict(body)})
a,b=rows;body_same=Counter(a['body'])==Counter(b['body']);lost=Counter(a['all'])-Counter(b['all']);gained=Counter(b['all'])-Counter(a['all']);lost_f=Counter(a['furniture'])-Counter(b['furniture']);gained_f=Counter(b['furniture'])-Counter(a['furniture'])
result={'method':'Same Chromium PDF source bytes. Source-specific physical header/footer bands top<72pt or top>790pt, below/above the authored SOP body margins. Classifies page furniture multiplicity separately; not an OCR or raw-source preservation substitute.','sameBodyGlyphInventory':body_same,'totalDeltaExactlyFurnitureDelta':lost==lost_f and gained==gained_f,'lost':dict(lost),'gained':dict(gained),'rows':rows};(B/('sop-pdf-furniture-inventory-'+stages[1]+'.json')).write_text(json.dumps(result,ensure_ascii=False,indent=2));print({k:v for k,v in result.items()if k!='rows'})
