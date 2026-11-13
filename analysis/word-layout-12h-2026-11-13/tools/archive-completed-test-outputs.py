"""Preserve earlier completed QA outputs losslessly while reducing duplication."""
from pathlib import Path
import tarfile,json,hashlib,datetime,subprocess,shutil
B=Path(__file__).resolve().parents[1]
names=['pc-b6edda3','pc-1012ae8','pc-fe9-frozen','metrics-furniture-focused'];files=[]
for name in names:
    root=B/name
    if not root.exists():continue
    opened=subprocess.run(['lsof','-n','+D',str(root)],capture_output=True,text=True)
    assert opened.returncode==1 and not opened.stdout.strip(),name
    for p in root.rglob('*'):
        assert not p.is_symlink(),p
        if p.is_file():files.append(dict(file=str(p.relative_to(B)),bytes=p.stat().st_size,sha256=hashlib.sha256(p.read_bytes()).hexdigest()))
assert files
archive=B/'completed-prior-pc-outputs.tar.xz'
with tarfile.open(archive,'w:xz',preset=9)as tf:
    for name in names:
        if (B/name).exists():tf.add(B/name,arcname=name)
expected={r['file']:r for r in files};seen=set()
with tarfile.open(archive,'r:xz')as tf:
    for entry in tf:
        if entry.isdir():continue
        assert entry.isfile() and entry.name in expected
        raw=tf.extractfile(entry).read();r=expected[entry.name]
        assert len(raw)==r['bytes'] and hashlib.sha256(raw).hexdigest()==r['sha256'];seen.add(entry.name)
assert seen==set(expected)
result=dict(utc=datetime.datetime.now(datetime.timezone.utc).isoformat(),method='Earlier completed owned QA outputs only, including earlier failures/screenshots/DOCX evidence; all raw file byte SHAs verified after archive reopen. Final PC6bc, active government checks, raw source/user recovery and current native outputs retained directly.',archive=archive.name,archiveBytes=archive.stat().st_size,archiveSHA256=hashlib.sha256(archive.read_bytes()).hexdigest(),originalBytes=sum(r['bytes']for r in files),files=files)
result['reclaimedBytes']=result['originalBytes']-result['archiveBytes']
assert result['reclaimedBytes']>0
(B/'cleanup-completed-prior-pc-outputs.json').write_text(json.dumps(result,indent=2)+'\n')
for name in names:
    if(B/name).exists():shutil.rmtree(B/name)
print({k:v for k,v in result.items()if k not in ['files','method']},flush=True)
