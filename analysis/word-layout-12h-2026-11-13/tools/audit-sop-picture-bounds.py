"""Read-only original/saved OOXML extents and actual native picture bounds."""
from pathlib import Path,PurePosixPath
from lxml import etree
import json,gzip,zipfile,hashlib,base64,datetime
B=Path(__file__).resolve().parents[1];R=B.parents[1]
NS={'wp':'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing','a':'http://schemas.openxmlformats.org/drawingml/2006/main','pic':'http://schemas.openxmlformats.org/drawingml/2006/picture','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}

def drawings(file):
    result={}
    with zipfile.ZipFile(file)as z:
        rels={r.get('Id'):r.get('Target')for r in etree.fromstring(z.read('word/_rels/document.xml.rels'))}
        root=etree.fromstring(z.read('word/document.xml'))
        for e in root.xpath('//wp:anchor|//wp:inline',namespaces=NS):
            extent=e.find('wp:extent',NS);paint=e.find('.//pic:spPr/a:xfrm/a:ext',NS)
            for rid in e.xpath('.//a:blip/@r:embed',namespaces=NS):
                target=rels.get(rid,'');part=str(PurePosixPath('word')/target)if not target.startswith('/')else target.lstrip('/')
                h=hashlib.sha256(z.read(part)).hexdigest();item={'wpExtent':dict(extent.attrib)if extent is not None else None,'pictureExtent':dict(paint.attrib)if paint is not None else None}
                if item not in result.setdefault(h,[]):result[h].append(item)
    return result

native=json.loads(gzip.decompress((B/'native-sop-picture-bounds-final-6bc49fa.json.gz').read_bytes()))
doc=json.loads(gzip.decompress((B/'native-sop-6bc49fa-cover-retry/pass-0-final.wordstudio.gz').read_bytes()));images=[]
def walk(n):
    if n['type']in ['image','inlineImage']:images.append(n['attrs'])
    for c in n.get('content',[]):walk(c)
walk(doc['content']);assert len(images)==len(native['floatingPictures'])==40
original=R/'analysis/hwp-cd-10h-2026-11-07/word-verified-upload/WS-CD-20261107-VERIFIED2-mois-sop-3.docx'
saved=B/'roundtrip-government-6bc49fa/textbox-save-mois-sop-3-save2.docx'
a=drawings(original);z=drawings(saved)
word=json.loads((B/'sop-word-picture-pages-rgba25.json').read_text());reference={r['sourceIndex']:r['word']for r in word['comparisons']['product-77d0d06']['rows']};rows=[]
for index,source in enumerate(images):
    if index not in reference:continue
    h=hashlib.sha256(base64.b64decode(source['src'].split(',')[1])).hexdigest();view=native['floatingPictures'][index];w=reference[index];frame=view['frame'];rect=view['box']
    paper=next((p for p in native['papers']if p['y']-.5<=rect['y']<p['y']+p['height']+.5),None)
    expected_right=rect['x']+source['width'];expected_bottom=rect['y']+source['height']
    fits=bool(paper and rect['x']>=paper['x']-.5 and expected_right<=paper['x']+paper['width']+.5 and expected_bottom<=paper['y']+paper['height']+.5)
    rows.append({'sourceIndex':index,'sourceRasterSHA256':h,'modelWidth':source['width'],'modelHeight':source['height'],'nativeWidth':rect['width'],'nativeHeight':rect['height'],'nativeComputedMaxWidth':view['computed']['maxWidth'],'paragraphWidth':view['paragraph']['box']['width'],'wordWidthCSSpx':w['width'],'wordHeightCSSpx':w['height'],'wordPage':w['page'],'wordPixelIdentityMethod':w['match'],'originalSourceExtents':a.get(h,[]),'secondSaveExtents':z.get(h,[]),'sourceExtentSetsPreserved':a.get(h,[])==z.get(h,[]),'actualNativeWidthSmallerThanModel':rect['width']<source['width']-.5,'authoredSizeWouldFitActualPhysicalPaper':fits})
result={'utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'sourceCommit':native['sourceCommit'],'inputSHA256':native['inputSHA256'],'historicalWordPDFSHA256':word['wordPDFSHA256'],'method':'Original source package and actual second-save DOCX extents; original body image pixels identify the matching historical Word drawing. Actual frozen native DOM bounds read without style overrides. Paper-fit uses the current native top-left plus authored extent; not a general crop/rotate/group solver. No new Microsoft measurement or primary 29-document score adjustment.','matchedPaintedSourcePictures':len(rows),'renderedWidthSmallerThanModel':sum(r['actualNativeWidthSmallerThanModel']for r in rows),'secondSaveExtentSetDifferences':sum(not r['sourceExtentSetsPreserved']for r in rows),'CSSRule':'src/styles.css .tiptap img { max-width:100%; height:auto; }','rows':rows}
(B/'remaining-sop-picture-bounds.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n');print({k:v for k,v in result.items()if k!='rows'});print('shrunk',[{k:r[k]for k in ['sourceIndex','modelWidth','nativeWidth','wordWidthCSSpx','authoredSizeWouldFitActualPhysicalPaper','sourceExtentSetsPreserved']}for r in rows if r['actualNativeWidthSmallerThanModel']])
