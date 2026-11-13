from pathlib import Path
import os,subprocess,json
R=Path(__file__).resolve().parents[3];B=Path(__file__).resolve().parents[1];build=json.loads((B/'native-checkpoint-build.json').read_text());ref=build['sourceCommit'];tag=ref[:7]
env={**os.environ,'GOV_RT_SET':'word','GOV_RT_LABEL':'word-'+tag,'GOV_RT_REF':ref,'GOV_RT_KEEP_SAVE1':'1','GOV_RT_PORT':'5747'}
with(B/('roundtrip-word-'+tag+'.log')).open('w')as log:subprocess.run(['node',str(B/'tools/corpus-roundtrip.cjs')],cwd=R,env=env,stdout=log,stderr=subprocess.STDOUT,check=True)
with(B/('native-word-'+tag+'-runner.log')).open('w')as log:subprocess.run(['python3',str(B/'tools/run-native-word-checkpoint.py')],cwd=R,env={**os.environ,'NATIVE_WORD_TAG':tag,'NATIVE_WORD_ROUNDTRIP':'word-'+tag},stdout=log,stderr=subprocess.STDOUT,check=True)
print('Completed fresh Word 10 original/two exports and 30 actual native opens on '+ref)
