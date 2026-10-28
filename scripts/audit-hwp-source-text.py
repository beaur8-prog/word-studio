"""Compare all logical body text, including short fragments, counts and order.

Run export-hwp-audit-source.cjs first. The model directory holds fresh converted
*.wordstudio.gz files. This checks the native parser's XML against the editable
converter, not the parser against Hancom or all drawing/field pixels.
"""
import argparse
import difflib
import gzip
import hashlib
import json
import re
from pathlib import Path
from xml.etree import ElementTree as ET

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('--evidence', type=Path, default=Path('analysis/hwp-completion-fixes-2026-10-26'))
parser.add_argument('--models', type=Path)
parser.add_argument('--output', type=Path)
args = parser.parse_args()
models = args.models or args.evidence / 'converted-current'
sources = json.loads((args.evidence / 'source-sections.json').read_text())
hashes = json.loads((args.evidence / 'source-audit-hashes.json').read_text())


def normalized(text):
    # These two legacy HWP glyphs have explicit, tested Unicode replacements.
    return re.sub(r'\s|[\u0000-\u0008\u200b\ufeff]', '', text.replace('\U000f012b', '(인)').replace('\U000f081c', '┈'))


def source_body(node):
    name = node.tag.rsplit('}', 1)[-1]
    if name in ('header', 'footer', 'footNote', 'endNote', 'rect', 'ellipse', 'polygon', 'container'):
        return ''
    if name == 't':
        return ''.join(node.itertext())
    return ''.join(source_body(child) for child in node)


def model_body(node):
    if node['type'] == 'documentNote':
        return ''
    return node.get('text', '') + ''.join(model_body(c) for c in node.get('content', []))


samples = []
for sample, sections in sources.items():
    original = normalized(''.join(source_body(ET.fromstring(xml)) for xml in sections.values()))
    document = json.loads(gzip.decompress((models / (sample + '.wordstudio.gz')).read_bytes()))
    converted = normalized(model_body(document['content']))
    differences = [] if original == converted else [
        {'kind': kind, 'sourceOffset': start, 'convertedOffset': target,
         'sourceLength': end - start, 'convertedLength': tail - target,
         'sourceExcerpt': original[start:end][:80], 'convertedExcerpt': converted[target:tail][:80]}
        for kind, start, end, target, tail in difflib.SequenceMatcher(None, original, converted, autojunk=False).get_opcodes()
        if kind != 'equal'
    ]
    samples.append({'id': sample, 'sourceFileSHA256': hashes[sample], 'sourceChars': len(original), 'convertedChars': len(converted),
                    'sourceTextSHA256': hashlib.sha256(original.encode()).hexdigest(), 'convertedTextSHA256': hashlib.sha256(converted.encode()).hexdigest(),
                    'identicalLogicalBodyStream': original == converted, 'differences': differences})
result = {'scope': 'Logical body/table text, including short fragments, multiplicity and order; ignores whitespace and normalizes two documented legacy glyph aliases.',
          'limitations': ['Native source parser is shared; not an independent Hancom parsing check.', 'Drawing text, fields, notes, headers/footers and all-page visual fidelity require separate checks.'],
          'totalSourceCharacters': sum(s['sourceChars'] for s in samples), 'samples': samples}
output = args.output or args.evidence / 'source-logical-order-audit.json'
output.write_text(json.dumps(result, ensure_ascii=False, indent=2) + '\n')
print(f"{sum(s['identicalLogicalBodyStream'] for s in samples)}/{len(samples)} exact logical body streams; {result['totalSourceCharacters']} characters.")
raise SystemExit(0 if all(s['identicalLogicalBodyStream'] for s in samples) else 1)
