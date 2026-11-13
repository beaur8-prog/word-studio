from pathlib import Path
import json,hashlib,logging,datetime
import pdfplumber
logging.getLogger('pdfminer').setLevel(logging.ERROR)
B=Path(__file__).resolve().parents[1]
p=B/'docx-pdf-picture-content-6bc49fa-hwp-saved2.json';j=json.loads(p.read_text())
for row in j['rows']:
    pending=row['unmatchedRasterAssets'];proofs=[]
    if pending:
        byte_hashes={a['bytesSHA256']:a for a in pending}
        with pdfplumber.open(B/'render-final-6bc49fa-hwp-saved2'/(row['id']+'.pdf'))as pdf:
            for number,page in enumerate(pdf.pages,1):
                for image in page.images:
                    stream=image['stream'];h=hashlib.sha256(stream.get_data()).hexdigest()
                    if h in byte_hashes:
                        proofs.append(dict(sourceMedia=byte_hashes[h]['name'],encodedBytesSHA256=h,objectId=stream.objid,page=number,filter=str(stream.get('Filter'))))
                page.close()
    row['exactEncodedStreamMatchesForNonExactPixels']=proofs
    row['exactEncodedStreamMatchedAssets']=len({p['sourceMedia']for p in proofs})
    row['allRasterAssetsAccountedFor']=row['exactPixelMatchedAssets']+row['exactEncodedStreamMatchedAssets']==row['uniqueRasterAssets']
    assert row['allRasterAssetsAccountedFor'] and not row['nonRasterOrUndecodedSourceAssets'] and not row['decodeErrors']
j['JPEGExplanation']='Five JPEG media are not pixel-hash identical after two different decoding paths; each original encoded JPEG byte stream is exactly present in an actually painted PDF DCTDecode XObject. Those five are proven by encoded-byte SHA, without a near-pixel heuristic. Pixel-unmatched entries retained above to distinguish the methods.'
j['totalUniqueRasterAssets']=sum(r['uniqueRasterAssets']for r in j['rows']);j['totalExactPixelAssets']=sum(r['exactPixelMatchedAssets']for r in j['rows']);j['totalExactEncodedJPEGAssets']=sum(r['exactEncodedStreamMatchedAssets']for r in j['rows'])
j['finalizedUTC']=datetime.datetime.now(datetime.timezone.utc).isoformat();assert j['totalUniqueRasterAssets']==412 and j['totalExactPixelAssets']==407 and j['totalExactEncodedJPEGAssets']==5
p.write_text(json.dumps(j,ensure_ascii=False,indent=2)+'\n');print({k:v for k,v in j.items()if k.startswith('total')})
