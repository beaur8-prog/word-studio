"""Match decoded source DOCX rasters to actually painted PDF XObjects.
Exact RGB pixel matches only. Paper/cell clipping, order and position are separate checks.
"""
from pathlib import Path
from collections import defaultdict
from io import BytesIO
import hashlib,json,logging,zipfile,datetime
from PIL import Image
import pdfplumber
from pypdf import PdfReader
logging.getLogger('pdfminer').setLevel(logging.ERROR)
B=Path(__file__).resolve().parents[1]

def pixels(im):
    rgba=im.convert('RGBA');white=Image.new('RGBA',rgba.size,'white');white.alpha_composite(rgba)
    rgb=white.convert('RGB')
    return (rgb.width,rgb.height,hashlib.sha256(rgb.tobytes()).hexdigest())

rows=[]
for source in json.loads((B/'fresh-hwp-6bc49fa/results.json').read_text())['rows']:
    ident=source['id'];docx=B/source['saves'][1]['file'];pdf=B/'render-final-6bc49fa-hwp-saved2'/(ident+'.pdf')
    assets=[];unhandled=[]
    with zipfile.ZipFile(docx)as z:
        for name in z.namelist():
            if not name.startswith('word/media/') or name.endswith('/'):continue
            raw=z.read(name)
            try:
                with Image.open(BytesIO(raw))as im:key=pixels(im)
                assets.append(dict(name=name,bytesSHA256=hashlib.sha256(raw).hexdigest(),width=key[0],height=key[1],pixelSHA256=key[2]))
            except Exception as e:unhandled.append(dict(name=name,format=Path(name).suffix,reason=type(e).__name__))
    painted=defaultdict(list)
    with pdfplumber.open(pdf)as p:
        for page_no,page in enumerate(p.pages,1):
            for image in page.images:
                stream=image.get('stream');object_id=getattr(stream,'objid',None)
                if object_id is not None:painted[object_id].append(page_no)
            page.close()
    decoded={};errors=[];reader=PdfReader(pdf)
    for page in reader.pages:
        try:images=page.images
        except Exception as e:errors.append(dict(reason=type(e).__name__,phase='list'));continue
        for image in images:
            ref=image.indirect_reference
            object_id=ref.idnum if ref is not None else None
            if object_id not in painted or object_id in decoded:continue
            try:decoded[object_id]=pixels(image.image)
            except Exception as e:errors.append(dict(objectId=object_id,reason=type(e).__name__,phase='decode'))
    reverse=defaultdict(list)
    for object_id,key in decoded.items():reverse[key].append(dict(objectId=object_id,pages=sorted(set(painted[object_id]))))
    matched=[];missing=[]
    for asset in assets:
        matches=reverse.get((asset['width'],asset['height'],asset['pixelSHA256']),[])
        item=dict(**asset,paintedPDFObjects=matches)
        (matched if matches else missing).append(item)
    row=dict(id=ident,docxSHA256=hashlib.sha256(docx.read_bytes()).hexdigest(),pdfSHA256=hashlib.sha256(pdf.read_bytes()).hexdigest(),sourceModelImageReferences=len(source['states'][0]['images']),
        uniqueRasterAssets=len(assets),exactPixelMatchedAssets=len(matched),unmatchedRasterAssets=missing,nonRasterOrUndecodedSourceAssets=unhandled,
        paintedPDFObjectIds=len(painted),decodedPaintedPDFObjects=len(decoded),decodeErrors=errors,exactMatches=matched)
    rows.append(row)
    result=dict(utc=datetime.datetime.now(datetime.timezone.utc).isoformat(),method='Exact decoded RGB pixels, transparent input composited on white, source DOCX media against PDF XObjects independently confirmed in actual PDF paint operators by object ID. Counts unique media, not source reference multiplicity. Unmatched assets require review; no near-match heuristics, no claim about full visibility, cell clip, order, page accuracy, or raw HWP completeness.',rows=rows)
    (B/'docx-pdf-picture-content-6bc49fa-hwp-saved2.json').write_text(json.dumps(result,ensure_ascii=False,indent=2))
    print(json.dumps({k:v for k,v in row.items()if k not in ['exactMatches','unmatchedRasterAssets']},ensure_ascii=False),flush=True)
