"""Match single-source-line tab paragraphs as complete physical PDF lines.

PDF paint streams include generated leader glyphs and have no paragraph
boundaries. Matching the concatenated logical text can instead find a later
heading plus a separate numbered cell. Only exact normalized text segments
with optional tab-leader glyphs between them, on one physical line, qualify.
Ambiguous or wrapped cases remain explicitly unresolved, never fuzzy matched.
"""
import argparse, collections, json, logging, re, unicodedata
from pathlib import Path
import pdfplumber

logging.getLogger('pdfminer').setLevel(logging.ERROR)
p = argparse.ArgumentParser()
p.add_argument('source', type=Path)
p.add_argument('pdf', type=Path)
p.add_argument('prior', type=Path)
p.add_argument('output', type=Path)
a = p.parse_args()

def norm(text):
    return re.sub(r'[\s\x00-\x1f\ufffc]', '', unicodedata.normalize('NFKC', text or ''))

def lines(page):
    groups = []
    for char in sorted(page.chars, key=lambda c: ((c['top'] + c['bottom']) / 2, c['x0'])):
        text = norm(char['text'])
        if not text:
            continue
        center = (char['top'] + char['bottom']) / 2
        height = char['bottom'] - char['top']
        group = next((g for g in reversed(groups) if abs(g['center'] - center) < max(1.5, min(g['height'], height) * .35)), None)
        if group is None:
            group = {'center': center, 'height': height, 'chars': []}
            groups.append(group)
        group['chars'].append(char)
    result = []
    for group in groups:
        chars = sorted(group['chars'], key=lambda c: c['x0'])
        result.append({'text': ''.join(norm(c['text']) for c in chars),
                       'left': min(c['x0'] for c in chars), 'right': max(c['x1'] for c in chars),
                       'top': min(c['top'] for c in chars), 'bottom': max(c['bottom'] for c in chars)})
    return result

source = json.loads(a.source.read_text())
prior = {r['key']: r for r in json.loads(a.prior.read_text())['paragraphs']}
counts = collections.Counter(norm(r['text']) for r in source['paragraphs'])
physical = []
with pdfplumber.open(a.pdf) as pdf:
    for page_no, page in enumerate(pdf.pages, 1):
        physical.extend({'page': page_no, **line} for line in lines(page))
        page.close()

rows = []
for para in source['paragraphs']:
    # Source "lines" are painted style runs. Several fonts on the same
    # baseline still represent one source line, rather than several lines.
    runs = para.get('lines', [])
    single_line = bool(runs) and len({run['page'] for run in runs}) == 1 and \
        max(run['y'] for run in runs) - min(run['y'] for run in runs) < max(1, min(run['height'] for run in runs) * .35)
    if '\t' not in para['text'] or not single_line:
        continue
    parts = [norm(part) for part in para['text'].split('\t')]
    row = {'key': para['key'], 'text': para['text'], 'sourcePages': sorted({l['page'] for l in para['lines']}),
           'method': 'exact-complete-physical-line-with-generated-tab-leaders',
           'previous': {k: prior.get(para['key'], {}).get(k) for k in ['status', 'pageDifference', 'actual']}}
    if counts[norm(para['text'])] != 1 or not all(parts):
        row['status'] = 'ambiguousSource' if counts[norm(para['text'])] != 1 else 'emptyTabSegment'
        rows.append(row)
        continue
    pattern = re.compile('^' + r'[-._·•…‥⋅․─┄┈–—]*'.join(re.escape(part) for part in parts) + '$')
    matches = [line for line in physical if pattern.fullmatch(line['text'])]
    row['matches'] = matches
    row['status'] = 'matched' if len(matches) == 1 else 'ambiguousPDF' if matches else 'unresolvedPhysicalLine'
    if len(matches) == 1:
        row['actual'] = matches[0]
        row['pageDifference'] = matches[0]['page'] - min(row['sourcePages'])
    rows.append(row)

report = {'scope': __doc__, 'source': str(a.source), 'pdf': str(a.pdf),
          'summary': {'eligibleSingleLineTabParagraphs': len(rows), 'statuses': dict(collections.Counter(r['status'] for r in rows)),
                      'correctedWrongPageMatches': sum(r['status'] == 'matched' and r['previous'].get('status') == 'matched' and r['pageDifference'] != r['previous'].get('pageDifference') for r in rows)},
          'paragraphs': rows}
a.output.parent.mkdir(parents=True, exist_ok=True)
a.output.write_text(json.dumps(report, ensure_ascii=False, indent=2))
print(json.dumps(report['summary'], ensure_ascii=False))
for row in rows:
    if row['status'] == 'matched' and row['pageDifference'] != row['previous'].get('pageDifference'):
        print(json.dumps({k: row.get(k) for k in ['key', 'text', 'pageDifference', 'previous', 'actual']}, ensure_ascii=False))
