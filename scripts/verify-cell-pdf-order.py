"""Check actual browser PDFs against the logical text of the mixed-cell fixtures."""
import re
import sys
from pathlib import Path

from pypdf import PdfReader

root = Path(sys.argv[1] if len(sys.argv) > 1 else 'test-results')
expected_files = sorted(root.rglob('cell-pdf-expected.txt'))
count = int(sys.argv[2]) if len(sys.argv) > 2 else 7
if len(expected_files) != count:
    raise SystemExit(f'Expected {count} HWP/HWPX cell PDFs, found {len(expected_files)} in {root}')
for expected_file in expected_files:
    pdf_file = expected_file.with_name('cell.pdf')
    actual = ''.join(page.extract_text() or '' for page in PdfReader(pdf_file).pages)
    actual = re.sub(r'\s+', '', actual)
    expected = expected_file.read_text(encoding='utf-8')
    if actual != expected:
        first = next((i for i, pair in enumerate(zip(actual, expected)) if pair[0] != pair[1]), min(len(actual), len(expected)))
        raise SystemExit(f'PDF logical text differs at character {first}: {pdf_file}')
    print(f'{expected_file.parent.name}: PDF logical text preserved ({len(expected)} characters)')
print(f'{count}/{count} cell PDFs preserve logical text order.')
