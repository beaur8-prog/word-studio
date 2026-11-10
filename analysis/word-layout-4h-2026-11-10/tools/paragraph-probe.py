from pathlib import Path
import json, gzip, collections, hashlib, logging, os
from zipfile import ZipFile
from xml.etree import ElementTree as E
logging.getLogger('pdfminer').setLevel(logging.ERROR)
R=Path(__file__).resolve().parents[3]; B=Path(__file__).resolve().parents[1]; P=R/'analysis/hwp-cd-10h-2026-11-07'
ns={}; exec((R/'analysis/docx-word-layout-2026-10-31/tools/docx-paired-baseline.py').read_text().split('rows=[];cohorts=[]')[0],ns)
extract,locate,norm=[ns[k] for k in ['extract','locate','norm']]
manifest=json.loads((P/'final-word-studio-comparison.json').read_text())
W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
ids=os.environ.get('LAYOUT_IDS','mods-household-1,mss-funding-2026-0').split(',')
stage=os.environ.get('LAYOUT_STAGE','frozen-baseline')
for item in manifest['rows']:
 if item['id'] not in ids: continue
 file=P/item['inputFile']; assert hashlib.sha256(file.read_bytes()).hexdigest()==item['inputSHA256']
 with ZipFile(file) as z: xml=E.fromstring(z.read('word/document.xml'))
 parents={c:p for p in xml.iter() for c in p}; source=[]
 for key,p in enumerate(xml.iter(W+'p')):
  if any(n is not p for n in p.iter(W+'p')): continue
  text=''.join((n.text or '') if n.tag==W+'t' else '\t' if n.tag==W+'tab' else '\n' if n.tag==W+'br' else '' for n in p.iter())
  anc=p; cell=False
  while anc in parents:
   anc=parents[anc]
   if anc.tag==W+'tc':cell=True
  spacing=p.find(W+'pPr/'+W+'spacing');pr=p.find(W+'pPr')
  source.append({'key':key,'text':text,'cell':cell,'spacing':{k.split('}')[-1]:v for k,v in spacing.attrib.items()} if spacing is not None else {},'keep':pr is not None and pr.find(W+'keepNext') is not None})
 counts=collections.Counter(norm(r['text'])for r in source)
 eligible=[r for r in source if len(norm(r['text']))>=12 and counts[norm(r['text'])]==1]
 # Freeze the source-eligible set before reading either PDF; retain all failures.
 cohort=[r['key'] for r in eligible]; cohortSHA=hashlib.sha256(json.dumps(cohort).encode()).hexdigest()
 sets={'word':extract(P/item['wordPDF']),'studio':extract(B/f'render-{stage}'/(item['id']+'.pdf'))}
 rows=[]
 for r in eligible:
  row=dict(r)
  for name,(stream,pos,pages) in sets.items():row[name]=locate(norm(r['text']),stream,pos)
  rows.append(row)
 match=[r for r in rows if r['word']['status']==r['studio']['status']=='matched']
 summary={'id':item['id'],'inputSHA256':item['inputSHA256'],'sourceCohortSHA256':cohortSHA,'sourceEligible':len(eligible),'matchedBoth':len(match),'unmatchedRetained':len(rows)-len(match),'wordPages':sets['word'][2],'studioPages':sets['studio'][2],'sameLineCount':sum(len(r['word']['lineEnds'])==len(r['studio']['lineEnds'])for r in match),'wordMoreLines':sum(len(r['word']['lineEnds'])>len(r['studio']['lineEnds'])for r in match),'wordFewerLines':sum(len(r['word']['lineEnds'])<len(r['studio']['lineEnds'])for r in match),'firstGlyphSizeMismatchesOver01pt':sum(abs(r['word']['firstSize']-r['studio']['firstSize'])>.1 for r in match),'startPageMismatches':sum(r['word']['origin']['page']!=r['studio']['origin']['page']for r in match),'absoluteStartPageError':sum(abs(r['word']['origin']['page']-r['studio']['origin']['page'])for r in match)}
 report={'method':'Development-only diagnostic. Prospective unique raw-OOXML paragraphs>=12 normalized characters, frozen before PDF lookup. Missing/ambiguous retained; not the historical 160/3007 or 398/140 denominators. Same frozen actual Word PDF/input hash.','summary':summary,'rows':rows}
 (B/f"{item['id']}-paragraphs-{stage}.json.gz").write_bytes(gzip.compress(json.dumps(report,ensure_ascii=False).encode()))
 (B/f"{item['id']}-paragraphs-{stage}-summary.json").write_text(json.dumps({'summary':summary,'firstPageDifferences':[r for r in match if r['word']['pages']!=r['studio']['pages']][:12],'firstLineDifferences':[r for r in match if r['word']['lineEnds']!=r['studio']['lineEnds']][:8]},ensure_ascii=False,indent=2))
 print(json.dumps(summary,ensure_ascii=False),flush=True)
