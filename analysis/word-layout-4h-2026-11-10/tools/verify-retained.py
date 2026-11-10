from pathlib import Path
import hashlib,json,gzip
B=Path(__file__).resolve().parents[1];R=B.parents[1];errors=[];checks=[]
def check_file(path,expected=None):
 if not path.is_file():errors.append({'path':str(path),'error':'missing'});return
 sha=hashlib.sha256(path.read_bytes()).hexdigest()
 if expected and sha!=expected:errors.append({'path':str(path),'error':'hash changed'})
 checks.append({'path':str(path.relative_to(R))if path.is_relative_to(R)else str(path),'sha256':sha})
d=json.loads((B/'government-final-comparison.json').read_text());assert len(d['rows'])==29
for row in d['rows']:
 check_file(B/row['sourceDOCX'],row['inputSHA256']);check_file(B/row['studioPDF'],row['pdfSHA256']);check_file(B/row['studioState']);check_file(B/row['microsoftPDF'])
 check_file(B/'render-frozen-baseline'/(row['id']+'.pdf'))
exports=json.loads((B/'word-export-zip-delta.json').read_text())
for row in exports['rows']:check_file(B/row['beforeFile'],row['beforeSHA256'])
sdk=json.loads((B/'sdk-input-manifest-final.json').read_text());assert sdk['expected']==71
for row in sdk['files']:check_file(B/row['source'],row['sha256'])
word=json.loads((R/'analysis/word-pagination-3h-2026-11-05/baseline.json').read_text())
for row in word['rows']:
 check_file(Path(row['path']),row['sha256']);check_file(B/'native-baseline'/(row['id']+'.pdf'));check_file(B/'native-final2'/(row['id']+'.pdf'))
report={'method':'Verify retained source bytes, all baseline/final/reference PDFs and exact originals of all71 SDK inputs after cleanup. Does not rerun layout.','checkedFiles':len(checks),'errors':errors,'files':checks}
(B/'retention-verification.json').write_text(json.dumps(report,indent=2));print(json.dumps({'checkedFiles':len(checks),'errors':errors}));raise SystemExit(bool(errors))
