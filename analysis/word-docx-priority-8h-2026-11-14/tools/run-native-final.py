from pathlib import Path
import json,subprocess,os,time
R=Path(__file__).resolve().parents[3];B=Path(__file__).resolve().parents[1]
build=json.loads((B/'native-build-final-5c3bd56.json').read_text());commit=subprocess.check_output(['git','rev-parse',build['ref']],cwd=R,text=True).strip()
source=json.loads((R/'analysis/word-pagination-3h-2026-11-05/baseline.json').read_text())
rt=json.loads((B/'roundtrip-default-size-b5744c1/results.json').read_text());saved={r['id']:r for r in rt['rows']}
assert not rt['errors'] and all(r['preservationPassed'] and r['gridStableAfterFirstSave'] for r in saved.values())
reports=[]
for save in [0,1,2]:
 label='native-final-5c3bd56-'+('original' if save==0 else 'save'+str(save));env={**os.environ,'LIST_STAGE':label,'LIST_IDS':'*','NATIVE_EDIT_CHECK':'1','NATIVE_BUILD_COMMIT':commit,'LAYOUT_NATIVE_MAIN':build['app']+'/electron/main.cjs'}
 if save:
  rows=[]
  for row in source['rows']:
   a=next(f for f in saved[row['id']]['files'] if f['save']==save)
   rows.append({**row,'path':str(B/a['file']),'sha256':a['sha256']})
  manifest=B/('native-final-5c3bd56-save'+str(save)+'-manifest.json');manifest.write_text(json.dumps({'rows':rows,'exportProducer':rt['sourceCommit'],'renderingCommit':commit},indent=2));env['LIST_MANIFEST']=str(manifest)
 start=time.time()
 with (B/(label+'.log')).open('w') as log:r=subprocess.run(['node',str(B/'tools/native-word.cjs')],cwd=R,env=env,stdout=log,stderr=subprocess.STDOUT,timeout=420)
 reports.append({'label':label,'sourceCommit':commit,'exportProducer':rt['sourceCommit'] if save else None,'exitCode':r.returncode,'seconds':round(time.time()-start,2)});(B/'native-final-5c3bd56-processes.json').write_text(json.dumps(reports,indent=2));print(reports[-1],flush=True)
 if r.returncode:break
