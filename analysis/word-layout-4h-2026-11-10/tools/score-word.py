from pathlib import Path
import gzip,json,logging,os,importlib.util
logging.getLogger('pdfminer').setLevel(logging.ERROR)
R=Path(__file__).resolve().parents[3];B=Path(__file__).resolve().parents[1]
ns={};exec((R/'analysis/docx-word-layout-2026-10-31/tools/docx-paired-baseline.py').read_text().split('rows=[];cohorts=[]')[0],ns)
extract,locate,norm,metric=[ns[k]for k in ['extract','locate','norm','metric']]
frozen=json.loads(gzip.decompress((R/'analysis/word-pagination-3h-2026-11-05/paragraph-comparison-final.json.gz').read_bytes()))
stage=os.environ.get('LAYOUT_WORD_STAGE','native-candidate');reports=[]
for doc in frozen['documents']:
 before=extract(B/'native-baseline'/(doc['id']+'.pdf'));after=extract(B/stage/(doc['id']+'.pdf'));rows=[]
 for p in doc['paragraphs']:
  row={**p}
  for name,(stream,pos,pages) in [('before',before),('after',after)]:row[name]={'status':'ambiguousSource'}if p['word']['status']=='ambiguousSource'else locate(norm(p['text']),stream,pos)
  rows.append(row)
 scores={}
 for cohort in ['common402','fixed150']:
  keys=doc['cohorts'][cohort]['lockedKeys'];scores[cohort]={name:metric([{**r,'studio':r[name]}for r in rows],keys)for name in ['before','after']}
 reports.append({'id':doc['id'],'pages':{'before':before[2],'after':after[2],'word':doc['WordPages']},'cohorts':scores,'paragraphs':rows})
 print(doc['id'],after[2],scores['common402']['after']['startPageMismatches'],flush=True)
summary={}
for cohort in ['common402','fixed150']:
 summary[cohort]={}
 for stageName in ['before','after']:
  scores=[d['cohorts'][cohort][stageName]for d in reports]
  summary[cohort][stageName]={k:sum(s[k]for s in scores)for k in ['requested','matchedBoth','lostOrAmbiguous','startPageMismatches','absolutePageError','exactLineParagraphs']}
result={'method':'Original Word10, identical historical Microsoft input hashes and same fixed398/140 keys. Fresh isolated Electron DPR1 baseline versus candidate. Not a new desktop Word measurement.','summary':summary,'documents':reports}
(B/('word-'+stage+'-scores.json.gz')).write_bytes(gzip.compress(json.dumps(result,ensure_ascii=False).encode()))
(B/('word-'+stage+'-summary.json')).write_text(json.dumps(summary,ensure_ascii=False,indent=2))
print(json.dumps(summary),flush=True)
