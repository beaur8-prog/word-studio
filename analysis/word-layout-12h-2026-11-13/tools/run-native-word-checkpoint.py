from pathlib import Path
import json,subprocess,os,time
R=Path(__file__).resolve().parents[3];B=Path(__file__).resolve().parents[1];build=json.loads((B/'native-checkpoint-build.json').read_text());reports=[];tag=os.environ.get('NATIVE_WORD_TAG',build['sourceCommit'][:7])
manifests=[None,B/'native-word-77d0d06-save1-manifest.json',B/'native-word-77d0d06-save2-manifest.json']
if os.environ.get('NATIVE_WORD_ROUNDTRIP'):
 source=json.loads((R/'analysis/word-pagination-3h-2026-11-05/baseline.json').read_text());roundtrip=json.loads((B/('roundtrip-'+os.environ['NATIVE_WORD_ROUNDTRIP'])/'results.json').read_text());saved={r['id']:r for r in roundtrip['rows']}
 if roundtrip['errors'] or any(not r['preservationPassed'] or not r['gridStableAfterFirstSave']for r in saved.values()):raise ValueError('DOCX roundtrip inventory failed')
 for save in [1,2]:
  rows=[]
  for row in source['rows']:
   artifact=next(f for f in saved[row['id']]['files']if f['save']==save)
   rows.append({**row,'path':str(B/artifact['file']),'sha256':artifact['sha256']})
  manifests[save]=B/f'native-word-{tag}-save{save}-manifest.json';manifests[save].write_text(json.dumps({'rows':rows,'sourceCommit':roundtrip.get('sourceCommit')},indent=2))
for label,manifest in zip([f'native-word-{tag}-original',f'native-word-{tag}-save1',f'native-word-{tag}-save2'],manifests):
 env={**os.environ,'LIST_STAGE':label,'LIST_IDS':'*','NATIVE_EDIT_CHECK':'1','NATIVE_BUILD_COMMIT':build['sourceCommit'],'LAYOUT_NATIVE_MAIN':build['app']+'/electron/main.cjs'}
 if manifest:env['LIST_MANIFEST']=str(manifest)
 start=time.time()
 with (B/(label+'.log')).open('w')as log:
  result=subprocess.run(['node',str(B/'tools/native-word.cjs')],cwd=R,env=env,stdout=log,stderr=subprocess.STDOUT,timeout=280)
 reports.append({'label':label,'sourceCommit':build['sourceCommit'],'exitCode':result.returncode,'seconds':round(time.time()-start,2)});(B/f'native-word-{tag}-processes.json').write_text(json.dumps(reports,indent=2));print(reports[-1],flush=True)
 if result.returncode:break
