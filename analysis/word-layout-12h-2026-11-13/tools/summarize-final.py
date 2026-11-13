"""Recompute final validation totals from retained evidence; no product mutations."""
from pathlib import Path
from collections import Counter
import json, gzip, re, datetime, hashlib

B = Path(__file__).resolve().parents[1]
FINAL = '6bc49fae11068d43b258f5568ceccd3c3cb8da30'

def read(name):
    p = B / name
    if not p.exists() and p.with_suffix(p.suffix + '.gz').exists():
        p = p.with_suffix(p.suffix + '.gz')
    return json.loads(gzip.decompress(p.read_bytes()) if p.suffix == '.gz' else p.read_bytes())

def write(name, value):
    (B / name).write_text(json.dumps(value, ensure_ascii=False, indent=2) + '\n')

base = read('render-page-segment-checkpoint/results.json')['rows']
final = read('render-final-6bc49fa-government/results.json')['rows']
base_by_id = {r['id']: r for r in base}
assert len(base_by_id) == len(final) == 29
rows = []
for f in final:
    a = base_by_id[f['id']]
    same_input = a['inputSHA256'] == f['inputSHA256']
    same_body = a['bodySHA256'] == f['bodySHA256']
    same_images = a['imagePlacements'] == f['imagePlacements']
    assert same_input and same_body and same_images
    assert a['wordPages'] == f['wordPages']
    rows.append(dict(id=f['id'], split=f['split'], inputSHA256=f['inputSHA256'],
        wordPages=f['wordPages'], beforePages=a['pages'], finalPages=f['pages'],
        beforeAbsolutePageError=abs(a['pages']-f['wordPages']),
        finalAbsolutePageError=abs(f['pages']-f['wordPages']),
        sameInput=same_input, sameBody=same_body, sameImagePlacements=same_images,
        tableFragments=[a['tableFragments'], f['tableFragments']]))
counts = {}
for split in ['all', 'development', 'validation', 'heldOut']:
    subset = rows if split == 'all' else [r for r in rows if r['split'] == split]
    counts[split] = dict(count=len(subset),
        beforeAbsolutePageError=sum(r['beforeAbsolutePageError'] for r in subset),
        finalAbsolutePageError=sum(r['finalAbsolutePageError'] for r in subset),
        beforeExact=sum(r['beforePages']==r['wordPages'] for r in subset),
        finalExact=sum(r['finalPages']==r['wordPages'] for r in subset))
assert counts['all']['beforeAbsolutePageError'] == counts['all']['finalAbsolutePageError'] == 39
assert counts['all']['beforeExact'] == 15 and counts['all']['finalExact'] == 16
write('final-government-comparison.json', dict(
    method='Same frozen original DOCX bytes and historical actual web Word print PDFs. Baseline is product 0bebe65/a58ad27; not the earlier 42-error run. No new Microsoft measurements. Prior heldOut documents are now regression controls, not unseen validation.',
    sourceCommit=FINAL, counts=counts, targetAbsolutePageError=30,
    targetMet=counts['all']['finalAbsolutePageError']<=30, rows=rows))

word = read('word-native-word-6bc49fa-original-scores.json.gz')
rt = read('roundtrip-word-6bc49fa/results.json')
assert rt['sourceCommit'] == FINAL and len(rt['rows']) == 10
assert all(r['preservationPassed'] and r['gridStableAfterFirstSave'] for r in rt['rows'])
processes = read('native-word-6bc49fa-processes.json')
assert len(processes) == 3 and all(r['exitCode']==0 and r['sourceCommit']==FINAL for r in processes)

fresh = read('fresh-hwp-6bc49fa/results.json')['rows']
assert len(fresh)==10
fresh_rows=[]
for r in fresh:
    assert r['ref']=='6bc49fa'
    states=r['states']; assert len(states)==3
    preserved=all(all(s[k]==states[0][k] for k in ['characters','bodySHA256','cellsSHA256','images']) for s in states[1:])
    pages=[r['printed']['pages']]+[s['pages'] for s in r['savedPrinted']]
    assert preserved and len(set(pages))==1
    fresh_rows.append(dict(id=r['id'], inputSHA256=r['inputSHA256'],
        originalSave1Save2Pages=pages, modelCharactersUTF16=states[0]['characters'],
        modelHardBreaks=states[0]['counts'].get('hardBreak',0),
        modelImages=len(states[0]['images']), bodyCellsImageBytesAndOrderPreserved=preserved))
assert sum(r['modelCharactersUTF16'] for r in fresh_rows)==288591
assert sum(r['modelHardBreaks'] for r in fresh_rows)==140
assert sum(r['modelImages'] for r in fresh_rows)==638

legacy=read('legacy-6bc49fa.json.gz')['rows']
legacy_totals={}
for key in ['fixed160','fixed3007']:
    legacy_totals[key]={k:sum(r[key][k] for r in legacy) for k in ['total','found','missing','mismatched','absolutePageError']}
assert legacy_totals['fixed160']['total']==160 and legacy_totals['fixed160']['mismatched']==78
assert legacy_totals['fixed3007']['total']==3007 and legacy_totals['fixed3007']['mismatched']==1819

pdf_before=read('pdf-assets-page-segment-checkpoint.json')['rows']
pdf_final=read('pdf-assets-final-6bc49fa-government.json')['rows']
pdf_base_by={r['id']:r for r in pdf_before}
pdf_rows=[]
for r in pdf_final:
    a=pdf_base_by[r['id']]
    assert a['inputSHA256']==r['inputSHA256']
    pdf_rows.append(dict(id=r['id'],sameGlyphMultiplicity=a['glyphCounts']==r['glyphCounts'],
        imageOperators=r['pdfImageOperators'], whollyOutsidePaper=len(r['whollyOutsidePaperImages'])))
assert sum(r['sameGlyphMultiplicity'] for r in pdf_rows)==28
assert sum(r['imageOperators'] for r in pdf_rows)==343
assert sum(r['whollyOutsidePaper'] for r in pdf_rows)==0
furniture=read('sop-pdf-furniture-inventory-final-6bc49fa-government.json')
assert furniture['sameBodyGlyphInventory'] and furniture['totalDeltaExactlyFurnitureDelta']
hwp_pdf=read('pdf-assets-final-6bc49fa-hwp-saved2.json')['rows']
assert len(hwp_pdf)==10 and sum(len(r['whollyOutsidePaperImages']) for r in hwp_pdf)==0

local={}
for key,log,expected in [('PC','pc-6bc49fa.log',893),('DPR2','dpr2-6bc49fa.log',82),('AndroidSimulation','android-6bc49fa.log',38)]:
    matches=re.findall(r'\b(\d+) passed \(([^)]+)\)',(B/log).read_text())
    assert matches and int(matches[-1][0])==expected
    local[key]=dict(passed=expected,duration=matches[-1][1],log=log)
node=(B/'native-tests-6bc49fa.log').read_text();assert re.search(r'(?:#|ℹ) pass 173\b',node) and re.search(r'(?:#|ℹ) fail 0\b',node)
local['Node']=dict(passed=173,failed=0,log='native-tests-6bc49fa.log')
sdk=[]
for p in sorted(B.glob('sdk-*6bc49fa*-summary.json')):
    s=json.loads(p.read_text());assert s['exitCode']==0 and not s['errors'];sdk.append(s)

tabs=read('native-tabs-6bc49fa/result.json');assert tabs['sourceCommit']==FINAL and len(tabs['rows'])==16 and not tabs['errors']
assert all(r['actualText']=='Label\tValue' and r['undoRedo'] for r in tabs['rows'])
hwp_native=read('native-hwp-6bc49fa/result.json');assert hwp_native['sourceCommit']==FINAL
sop=[]
for stage in ['native-sop-6bc49fa-body22','native-sop-6bc49fa-cover-retry']:
    if not (B/stage/'result.json').exists() and not (B/stage/'result.json.gz').exists():
        sop.append(dict(stage=stage,status='pending'));continue
    s=read(stage+'/result.json');assert s['sourceRef']==FINAL and s['sourcePreserved'] and s['hidden']
    assert len(s['phases'])==24 and not s['errors']
    assert s['bodyLength']==79406 and s['images']==40 and len(s['sourceSections'])==9
    assert all(p['final']['pages']==135 and p['final']['physicalParagraphPages']==135 for p in s['phases'])
    sop.append(dict(stage=stage,status='passed',phases=24,requiredContinuousStabilityMs=s['requiredStableMs'],
        editorPages=135,physicalParagraphPages=135,printPages=s.get('printPages'),
        sourcePreserved=s['sourcePreserved'],sourceCommit=s['sourceRef'],result=stage+'/result.json'+('' if (B/stage/'result.json').exists() else '.gz')))

summary=dict(utc=datetime.datetime.now(datetime.timezone.utc).isoformat(),sourceCommit=FINAL,
    method='Separate fixed cohorts and raw-byte proofs, no denominator mixing. New native means isolated compiled Word Studio, not Microsoft Word. OS chooser is mocked; Android is browser simulation, not physical APK/device QA.',
    government=counts,
    WordDOCX10=dict(pages=[d['pages'] for d in word['documents']], fixed=word['summary'],
        freshTwoSaveDocuments=20, preservationDocumentsPassed=10,nativeOpenEditUndoRedoPrintCases=30),
    freshHWP10=dict(modelScope='Fresh production Worker -> editable model -> two real DOCX saves; not independent original HWP logical 218937-character verification',
        modelCharactersUTF16=288591,modelHardBreaks=140,modelImages=638,savePageInvariantDocuments=10,rows=fresh_rows),
    legacyPreparedHWP10=dict(pages=[r['printPages'] for r in legacy],fixed=legacy_totals,
        unresolvedStrictInventoryDocuments=5,explanation='Known header/data cell role, raw-tab/documentTab, story materialization and section/column group differences remain unproven harmless; not declared resolved.'),
    actualPDF=dict(governmentRows=pdf_rows,governmentOperators=343,governmentWhollyOutsidePaper=0,
        governmentGlyphMultiplicityEqual=28,SOPBodyGlyphMultiplicityEqual=True,SOPDeltaExactlyPageFurniture=True,
        freshHWPSaved2SourceDOMImages=638,freshHWPSaved2PDFOperators=sum(r['pdfImageOperators'] for r in hwp_pdf),
        freshHWPSaved2WhollyOutsidePaper=0,
        limitation='Positive image-paper intersection is not full visibility, correct page ownership, no overlap or independent source completeness. Repeated header images can yield more PDF operators than source images.'),
    localTests=local, officialOpenXMLSDK=sdk,
    nativeTabs=dict(repetitions=16,twoMenuSaves=True,printPages=tabs['printPages'],errors=tabs['errors']),
    nativeHWP=dict(documents=len(hwp_native['rows']),phases=18,editorPrintPages=[dict(id=r['id'],editor=[p['editorPages'] for p in r['phases']],print=r['printPages']) for r in hwp_native['rows']]),
    nativeSOP=sop,
    unresolved=['Government absolute page error remains 39 (target<=30 unmet).',
        'SOP editor135/print133/historical Word136; two-page editor-print mismatch remains.',
        'Crossed vertical-merge table height cause not proven; exact-table A/B/C/D diagnostic inputs ready for actual Word.',
        'Five old strict HWP inventory issues remain; fresh byte/logical-cell checks do not erase them.',
        'One frozen final SOP cover export exceeded 120s; retry and same-model previous/final benchmark retained.',
        'New web Word and desktop Microsoft Word not measured this run while screen locked.',
        'Post-freeze performance candidate was blocked by automatic approval review and was not created or included; pending optional user exception is not approval.'])
gov_rt=read('government-6bc49fa-two-save-invariance.json')
gov_pdf=read('government-6bc49fa-saved2-pdf-invariance.json')
assert all(gov_rt[k]==29 for k in ['modelPreservationPassed','gridStableAfterFirstSave','independentXMLGeometryPreserved','samePages','sameBodySHA256','sameImagePlacements'])
assert gov_pdf['documents']==gov_pdf['sameGlyphMultiplicity']==gov_pdf['sameImageOperatorCount']==29 and gov_pdf['outsidePaper']==0
summary['governmentTwoRealDOCXSaves']={k:v for k,v in gov_rt.items() if k!='rows'}
summary['governmentSaved2ActualPDF']={k:v for k,v in gov_pdf.items() if k!='rows'}
summary['remainingOrdinaryParagraphReservation']=read('remaining-sop-paragraph-reservation.json')
glyphs=read('final-native-full-text-geometry.json')
assert glyphs['passed'] and glyphs['modelUnchanged'] and glyphs['achievedStableMs']>=40000
summary['nativeOriginalFullTextGeometry']=glyphs
bounds=read('remaining-sop-picture-bounds.json')
assert bounds['sourceCommit']==FINAL and bounds['matchedPaintedSourcePictures']==25 and bounds['secondSaveExtentSetDifferences']==0
summary['knownNativePictureSizeDifferences']={k:v for k,v in bounds.items() if k!='rows'}
summary['knownNativePictureSizeDifferences']['example']=next(r for r in bounds['rows'] if r['sourceIndex']==4)
summary['remainingInputReduction']=read('remaining-sop-input-reduction-summary.json')
summary['originalInputIntegrity']={k:v for k,v in read('final-original-input-integrity.json').items() if k!='rows'}
summary['unresolved'].insert(1,'SOP 13/25 historically Word-painted pictures are rendered narrower than authored model extents by paragraph max-width:100%; original/second-save OOXML extents remain preserved. No post-freeze product fix included.')
summary['unresolved'].insert(2,'SOP ordinary paragraph has excessive floating-flow widgets depending on surrounding pagination. Input-reduction diagnostics are not a completed minimal regression or a full stability acceptance test.')
ink=read('remaining-sop-native-print-first-drift.json')
assert ink['sourceCommit']==FINAL and not ink['errors'] and ink['hidden'] and ink['sourcePreserved']
summary['nativeSOPFirstGlyphAndPrintAudit']={k:ink[k] for k in ['editorPages','printPages','paragraphs','firstVisibleGlyphStarts','firstGlyphStartsOutsideBodyTolerance2CSSpx','firstGlyphStartCheckIsNotAllLinesVisibility','uniqueNativePrintComparison','historicalWordLookup']}
pictures=read('docx-pdf-picture-content-6bc49fa-hwp-saved2.json')
assert pictures['totalUniqueRasterAssets']==412 and all(r['allRasterAssetsAccountedFor'] for r in pictures['rows'])
summary['actualPDF']['freshHWPSaved2UniqueRasterCoverage']={k:pictures[k] for k in ['totalUniqueRasterAssets','totalExactPixelAssets','totalExactEncodedJPEGAssets']}
summary['actualPDF']['freshHWPSaved2UniqueRasterCoverage']['limitation']='Actually painted raster assets proven by exact decoded pixels or exact encoded JPEG bytes; full visibility, clipping, page/order and original HWP completeness remain separate.'
ci=B/'ci-final-6bc49fa.json'
if ci.exists():summary['crossPlatformCI']=json.loads(ci.read_text())
write('final-validation-summary.json',summary)
print(json.dumps(dict(government=counts,WordFixed=word['summary'],freshHWP10Invariant=True,
    oldHWPFixed=legacy_totals,local=local,SOP=sop,SDK=[dict(stage=s['stage'],files=s['files']) for s in sdk]),ensure_ascii=False))
