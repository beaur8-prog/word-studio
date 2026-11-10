from pathlib import Path
import json,os,subprocess,time
R=Path(__file__).resolve().parents[3];B=Path(__file__).resolve().parents[1]
build=json.loads((B/'native-figure-final-build.json').read_text());records=[]
for stage,manifest,ids in [
 ('native-figure-word-original',None,'*'),
 ('native-figure-word-save1',B/'native-word-normal-save1-manifest.json','*'),
 ('native-figure-word-save2',B/'native-word-normal-save2-manifest.json','*'),
 ('native-figure-nps',B/'native-government-manifest.json','nps-venture-hwp'),
]:
 env={**os.environ,'LIST_STAGE':stage,'LIST_IDS':ids,'NATIVE_EDIT_CHECK':'1','NATIVE_BUILD_COMMIT':build['sourceCommit'],'LAYOUT_NATIVE_MAIN':str(B/'native-app/electron/main.cjs')}
 if manifest:env['LIST_MANIFEST']=str(manifest)
 started=time.time()
 with (B/(stage+'.log')).open('w')as log:
  try:result=subprocess.run(['node',str(B/'tools/native-word.cjs')],cwd=R,env=env,stdout=log,stderr=subprocess.STDOUT,timeout=280);code=result.returncode
  except subprocess.TimeoutExpired:code='outer-timeout'
 records.append({'stage':stage,'exitCode':code,'seconds':round(time.time()-started,2),'sourceCommit':build['sourceCommit']})
 (B/'native-figure-final-processes.json').write_text(json.dumps(records,indent=2));print(records[-1],flush=True)
 if code!=0:break
