from pathlib import Path
import subprocess,json,tempfile,hashlib,shutil,datetime
R=Path(__file__).resolve().parents[3];B=Path(__file__).resolve().parents[1]
commit=subprocess.check_output(['git','rev-parse','HEAD'],cwd=R,text=True).strip()
if subprocess.check_output(['git','diff','--name-only','--','src'],cwd=R,text=True).strip():raise ValueError('Product source must be committed for native QA')
build=Path(tempfile.mkdtemp(prefix='ws-word-layout-12h-native-',dir='/private/tmp'))
with (B/'native-checkpoint-build.log').open('w')as log:
 subprocess.run([str(R/'node_modules/.bin/vite'),'build','--outDir',str(build)],cwd=R,stdout=log,stderr=subprocess.STDOUT,check=True)
app=B/('native-app-'+commit[:7]);app.mkdir(exist_ok=True);shutil.copytree(R/'electron',app/'electron',dirs_exist_ok=True);shutil.copy2(R/'package.json',app/'package.json')
for name,target in [('node_modules',R/'node_modules'),('dist',build)]:
 link=app/name
 if link.is_symlink():link.unlink()
 link.symlink_to(target,target_is_directory=True)
record={'utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'sourceCommit':commit,'dist':str(build),'app':str(app),'files':{str(f.relative_to(build)):hashlib.sha256(f.read_bytes()).hexdigest()for f in sorted(build.rglob('*'))if f.is_file()}}
(B/'native-checkpoint-build.json').write_text(json.dumps(record,indent=2));print({k:v for k,v in record.items()if k!='files'})
