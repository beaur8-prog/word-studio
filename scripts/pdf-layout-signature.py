#!/usr/bin/env python3
"""Compare real PDF paint, resources and extracted text. Requires pypdf.

Use --self-test to verify that coordinate-only changes cannot pass. Stream
objects may be false-y dictionaries despite containing paint instructions;
always distinguish None from an empty content stream explicitly.
"""
from pathlib import Path
import hashlib
import io
import json
import sys
from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject, DictionaryObject, DecodedStreamObject


def digest(data):
    return hashlib.sha256(data).hexdigest()


def signature(source):
    reader = PdfReader(source)

    def resolve(value, seen=frozenset()):
        if hasattr(value, 'get_object'):
            value = value.get_object()
        if id(value) in seen:
            return 'cycle'
        if isinstance(value, dict):
            seen = seen | {id(value)}
            result = {str(k): resolve(v, seen) for k, v in sorted(value.items())
                      if str(k) not in ('/Length', '/Filter', '/DecodeParms', '/Metadata')}
            if hasattr(value, 'get_data'):
                result['decodedStreamSHA256'] = digest(value.get_data())
            return result
        if isinstance(value, list):
            return [resolve(v, seen | {id(value)}) for v in value]
        if isinstance(value, bytes):
            return {'bytesSHA256': digest(value)}
        return str(value)

    pages = []
    for page in reader.pages:
        content = page.get_contents()
        # ContentStream inherits a dictionary's truthiness: "if content" can
        # discard nonempty streams and silently hash b'' for every page.
        paint = content.get_data() if content is not None else b''
        pages.append({
            'mediaBox': list(map(float, page.mediabox)),
            'cropBox': list(map(float, page.cropbox)),
            'rotation': int(page.get('/Rotate', 0)),
            'userUnit': float(page.get('/UserUnit', 1)),
            'paintBytes': len(paint), 'paintSHA256': digest(paint),
            'pageGroupSHA256': digest(json.dumps(resolve(page.get('/Group', {})), sort_keys=True).encode()),
            'resourcesSHA256': digest(json.dumps(resolve(page.get('/Resources', {})), sort_keys=True).encode()),
            'textSHA256': digest((page.extract_text() or '').encode()),
        })
    return {'format': 'pdf-layout-signature-v2', 'pages': len(pages), 'pageSignatures': pages}


def compare(current, reference):
    a, b = signature(current), signature(reference)
    paint_fields = ('mediaBox', 'cropBox', 'rotation', 'userUnit', 'paintBytes', 'paintSHA256', 'pageGroupSHA256', 'resourcesSHA256')
    same_pages = a['pages'] == b['pages']
    paint_equal = same_pages and all(all(p[k] == q[k] for k in paint_fields)
                                     for p, q in zip(a['pageSignatures'], b['pageSignatures']))
    text_equal = same_pages and all(p['textSHA256'] == q['textSHA256']
                                    for p, q in zip(a['pageSignatures'], b['pageSignatures']))
    return {'current': a, 'reference': b, 'paintResourcesExact': paint_equal,
            'textExact': text_equal, 'fullSignatureExact': a == b,
            # Legacy callers delete duplicate PDFs on this key. Keep the
            # stronger full comparison, with real paint bytes now included.
            'paintAndResourceExact': a == b}


def self_test():
    def document(x=40, font='Helvetica', title='A'):
        writer = PdfWriter()
        page = writer.add_blank_page(width=200, height=200)
        stream = DecodedStreamObject()
        stream.set_data(f'BT /F1 12 Tf {x} 100 Td (SAME TEXT) Tj ET'.encode())
        page[NameObject('/Contents')] = writer._add_object(stream)
        page[NameObject('/Resources')] = DictionaryObject({NameObject('/Font'): DictionaryObject({
            NameObject('/F1'): DictionaryObject({NameObject('/Type'): NameObject('/Font'),
                NameObject('/Subtype'): NameObject('/Type1'), NameObject('/BaseFont'): NameObject('/' + font)})})})
        writer.add_metadata({'/Title': title})
        result = io.BytesIO(); writer.write(result); return result.getvalue()
    base = document()
    same = compare(io.BytesIO(base), io.BytesIO(document(title='Different metadata')))
    moved = compare(io.BytesIO(base), io.BytesIO(document(x=60)))
    font = compare(io.BytesIO(base), io.BytesIO(document(font='Courier')))
    assert same['fullSignatureExact']
    assert moved['textExact'] and not moved['paintResourcesExact']
    assert not font['paintResourcesExact']
    assert moved['current']['pageSignatures'][0]['paintBytes'] > 0
    assert moved['current']['pageSignatures'][0]['paintSHA256'] != digest(b'')
    print(json.dumps({'selfTest': 'passed', 'coordinateOnlyChangeDetected': True,
                      'fontChangeDetected': True, 'metadataOnlyChangeIgnored': True,
                      'nonemptyStreamVerified': True}))


if __name__ == '__main__':
    if sys.argv[1:] == ['--self-test']:
        self_test()
    elif len(sys.argv) == 3:
        source = io.BytesIO(sys.stdin.buffer.read()) if sys.argv[1] == '-' else Path(sys.argv[1])
        print(json.dumps(compare(source, Path(sys.argv[2]))))
    else:
        raise SystemExit('Usage: pdf-layout-signature.py current.pdf reference.pdf | --self-test')
