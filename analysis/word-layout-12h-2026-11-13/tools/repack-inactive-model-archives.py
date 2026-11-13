"""Losslessly repack owned completed QA archives with a cross-snapshot dictionary."""
from pathlib import Path
import json,lzma,tarfile,hashlib,shutil,datetime,time
B=Path(__file__).resolve().parents[1];rows=[]
for name in ['native-sop-menu-all-eee8aa0','native-sop-77d0d06','native-sop-b6edda3-cover','native-sop-b6edda3-body22']:
    manifest=B/'native-model-archives'/(name+'.manifest.json');j=json.loads(manifest.read_text());src=B/j['archive']
    assert src.parent==B/'native-model-archives' and src.suffix=='.xz'
    old=src.read_bytes();assert hashlib.sha256(old).hexdigest()==j['archiveSHA256']
    tmp=src.with_suffix('.repack.xz');start=time.monotonic()
    with lzma.open(src,'rb')as f,lzma.open(tmp,'wb',preset=9)as out:shutil.copyfileobj(f,out,1024*1024)
    expected={r['entry']:r for r in j['files']};seen=set()
    with tarfile.open(tmp,'r:xz')as archive:
        for entry in archive:
            assert entry.isfile() and entry.name in expected
            data=archive.extractfile(entry).read();r=expected[entry.name]
            assert len(data)==r['uncompressedBytes'] and hashlib.sha256(data).hexdigest()==r['uncompressedSHA256']
            seen.add(entry.name)
    assert seen==set(expected)
    n=tmp.stat().st_size;h=hashlib.sha256(tmp.read_bytes()).hexdigest();adopt=n<len(old)
    row=dict(stage=name,beforeBytes=len(old),afterBytes=n,adopted=adopt,verifiedEntries=len(seen),seconds=round(time.monotonic()-start,2))
    if adopt:
        j.setdefault('archiveHistory',[]).append(dict(sha256=j['archiveSHA256'],bytes=j['bytes'],reason='Repacked exact same uncompressed entries with 64MiB LZMA dictionary; all raw SHAs verified.'))
        tmp.replace(src);j['archiveSHA256']=h;j['bytes']=n;j['compression']='XZ/LZMA2 preset9, 64MiB dictionary';manifest.write_text(json.dumps(j,indent=2)+'\n')
    else:tmp.unlink()
    rows.append(row);result=dict(utc=datetime.datetime.now(datetime.timezone.utc).isoformat(),method='Completed owned QA snapshots only; same original decompressed bytes and entry SHAs, no source/user data removal.',rows=rows,reclaimedBytes=sum(r['beforeBytes']-r['afterBytes']for r in rows if r['adopted']))
    (B/'cleanup-lossless-model-archive-repack.json').write_text(json.dumps(result,indent=2)+'\n');print(json.dumps(row),flush=True)
