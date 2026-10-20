"""Build a local evidence report from captured viewer pages; never modify inputs.

Requires Pillow, numpy, reportlab and fonttools. Capture Polaris through its UI
first, and run validate-hwp-samples.cjs for Studio. No viewer is a universal oracle.
"""
from pathlib import Path
import base64
import html
import hashlib
import json
import sys
import urllib.request

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / '.cache/hwp-report-python'))
import numpy as np
from PIL import Image, ImageDraw, ImageFont
from fontTools.ttLib import TTFont as FontFile
from fontTools.varLib.instancer import instantiateVariableFont
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.utils import ImageReader

BASE = ROOT / 'analysis/hwp-validation'
OUT = BASE / 'comparisons'
OUT.mkdir(exist_ok=True)
manifest = json.loads((ROOT / 'tests/hwp-samples.json').read_text())
studio = json.loads((BASE / 'studio-results.json').read_text())
polaris = json.loads((BASE / 'polaris-results.json').read_text())
by_file = {row['file']: row for row in studio['results']}
font_path = ROOT / '.cache/hwp-report-font-static.ttf'
if not font_path.exists():
    source_font = ROOT / '.cache/hwp-report-font.ttf'
    font_hash = '194018e6b2b293a7964f037b25c0249ce1418bc9ab3c971060a03aa57861e252'
    if not source_font.exists() or hashlib.sha256(source_font.read_bytes()).hexdigest() != font_hash:
        data = urllib.request.urlopen('https://raw.githubusercontent.com/google/fonts/47d54ca1ebb14fd864695f1d39dc734f39017c33/ofl/notosanskr/NotoSansKR%5Bwght%5D.ttf', timeout=30).read()
        if hashlib.sha256(data).hexdigest() != font_hash:
            raise ValueError('Report font SHA-256 mismatch')
        source_font.parent.mkdir(parents=True, exist_ok=True)
        source_font.write_bytes(data)
    variable = FontFile(source_font)
    instantiateVariableFont(variable, {'wght': 400}, inplace=True).save(font_path)
font = ImageFont.truetype(str(font_path), 18)
small = ImageFont.truetype(str(font_path), 13)
pdfmetrics.registerFont(TTFont('ReportKR', str(font_path)))


def normalized(path, viewer, width, height):
    image = Image.open(path).convert('RGB')
    if viewer == 'Polaris':
        # The actual viewer centers a single/last page vertically, but starts
        # a multipage document near the top. Find the white paper's first inner
        # column rather than mistaking viewer padding for a document margin.
        if image.size != (1280, 720):
            raise ValueError('Unexpected Polaris capture size; inspect paper bounds')
        pixels = np.asarray(image)
        probe_x = round((1265 - width * .5) / 2)
        rows = np.where((pixels[60:695, probe_x] > 245).all(axis=1))[0] + 60
        groups = np.split(rows, np.where(np.diff(rows) > 1)[0] + 1)
        paper = max(groups, key=len)
        if len(paper) < height * .5 - 15:
            raise ValueError('Paper boundary not found; inspect ' + str(path))
        columns = np.where((pixels[int(paper[0]) + 2, :1265] > 245).all(axis=1))[0]
        horizontal = next(group for group in np.split(columns, np.where(np.diff(columns) > 1)[0] + 1) if len(group) and group[0] < 632 < group[-1])
        box = (int(horizontal[0]), int(paper[0]), int(horizontal[-1]) + 1, int(paper[-1]) + 1)
        image = image.crop(box)
        if abs((image.width / image.height) / (width / height) - 1) > .03:
            raise ValueError('Paper aspect differs; inspect ' + str(path))
    elif viewer == 'Studio':
        image = image.crop((1, 1, image.width - 1, image.height - 1))
    return image.resize((560, 396) if width > height else (396, 560), Image.Resampling.LANCZOS)


def geometry(image):
    values = np.asarray(image.convert('L'))
    ink = values < 180
    ink[:3, :] = ink[-3:, :] = False
    ink[:, :3] = ink[:, -3:] = False
    y, x = np.where(ink)
    w, h = image.size
    mm_w, mm_h = (297, 210) if w > h else (210, 297)
    bounds = [round(float(x.min()) * mm_w / w, 1), round(float(y.min()) * mm_h / h, 1),
              round(float(x.max()) * mm_w / w, 1), round(float(y.max()) * mm_h / h, 1)] if len(x) else None
    rows = np.where(ink.sum(axis=1) > w * .38)[0]
    groups = np.split(rows, np.where(np.diff(rows) > 1)[0] + 1) if len(rows) else []
    return {'inkBoundsMm': bounds, 'longHorizontalInkRowsMm': [round(float(group.mean()) * mm_h / h, 1) for group in groups if len(group)]}


comparisons = []
def pair(file, page, reference, viewer, source_page=None):
    own = next(row for row in by_file[file]['rendered'] if row['number'] == page)
    left = normalized(reference, viewer, own['width'], own['height'])
    right = normalized(BASE / own['screenshot'], 'Studio', own['width'], own['height'])
    name = file.split('.')[0] + f'-{viewer.lower()}-page-{page}.png'
    right_x = 44 + left.width
    sheet = Image.new('RGB', (left.width * 2 + 56, left.height + 90), '#f2f4f7')
    draw = ImageDraw.Draw(sheet)
    title = next(row['title'] for row in manifest if row['file'] == file)
    draw.text((22, 10), title, font=font, fill='#172940')
    draw.text((22, 42), f'{viewer} · {source_page or page}쪽', font=small, fill='#4b596b')
    draw.text((right_x, 42), f'Word Studio 1.3.0 · {page}쪽', font=small, fill='#4b596b')
    sheet.paste(left, (22, 68)); sheet.paste(right, (right_x, 68))
    sheet.save(OUT / name)
    row = {'file': file, 'studioPage': page, 'referencePage': source_page or page, 'viewer': viewer,
           'image': 'comparisons/' + name, 'referenceGeometry': geometry(left), 'studioGeometry': geometry(right)}
    comparisons.append(row)
    return row


for reference in polaris['results']:
    file = reference['file']
    for number in reference['captured']:
        own_number = by_file[file]['pageCount'] if file.startswith('02-') and number == 163 else number
        pair(file, own_number, BASE / f"polaris/{file[:2]}-page-{number}.png", 'Polaris', number)
for file, prefix in [('03-health-duties.hwp', '03'), ('04-language-plan.hwp', '04'), ('10-text-align.hwp', '10')]:
    for own in by_file[file]['rendered']:
        number = own['number']
        reference = BASE / f'hancom-pdf/{prefix}-{number:02d}.png' if prefix == '04' else BASE / f'hancom-pdf/{prefix}-{number}.png'
        pair(file, number, reference, 'Hancom PDF')

report = {'date': '2026-10-20', 'version': 'Word Studio 1.3.0 (not publicly released)',
          'method': 'Same public input; Polaris UI at 50%; Studio SVG at 50%; Hancom PDF at 48 DPI. Detect paper frames, preserve orientation, normalize to 396x560 or 560x396. Bounds are approximate raster measurements, not accuracy scores.',
          'comparisons': comparisons}
(BASE / 'comparison-results.json').write_text(json.dumps(report, ensure_ascii=False, indent=2) + '\n')

table_rows = [
    ('표준근로계약서 HWP', 'Polaris', 7, 7, '7쪽 비교'),
    ('표준취업규칙 HWP', 'Polaris', 163, 154, '9쪽 차이 · 첫/마지막 쪽'),
    ('보건소 분장사무 HWP', 'Polaris / Hancom PDF', 4, 4, '4쪽 비교'),
    ('그림 HWP', 'Polaris', 5, 5, '5쪽 비교'),
    ('수식 HWP', 'Polaris', 1, 1, '수식·기준선 비교'),
    ('대각선 표 HWPX', 'Polaris', 1, 1, '대각선·테두리 비교'),
    ('국립국어원 업무계획 HWP', 'Hancom PDF', 35, 35, '1·2·10·35쪽 비교'),
    ('문단 맞춤 HWP', 'Hancom PDF', 1, 1, '맞춤·간격 비교'),
]

style = '''body{font:15px system-ui,sans-serif;background:#f3f5f8;color:#203047;max-width:1100px;margin:35px auto;padding:0 22px}h1{font-size:30px}p{line-height:1.7}.warning{background:#fff3d6;border-left:4px solid #cb8c00;padding:14px}table{border-collapse:collapse;width:100%;background:white}th,td{border:1px solid #d6dce5;padding:10px;text-align:left}th{background:#e9eef5}details{margin:20px 0;padding:16px;background:white;border-radius:9px}summary{cursor:pointer;font-weight:650}img{display:block;width:100%;max-width:848px;margin:14px auto}code{background:#e8edf4;padding:2px 5px}a{color:#225a9c}small{color:#58687d}'''
body = ['<!doctype html><html lang="ko"><meta charset="utf-8"><title>HWP 레이아웃 비교</title><style>' + style + '</style>',
        '<h1>HWP/HWPX 읽기 · 레이아웃 비교</h1><p>Word Studio 1.3.0 · 2026-10-20 · 소스·로컬 검증본. 공개 배포는 하지 않았습니다.</p>',
        '<p>공개 문서 11개 열기·화면 37쪽 검사, 암호/구형 문서 1개 거부. 폴라리스 웹의 실제 화면 6종과 한컴 PDF 참고본 3종을 비교했습니다. 표·그림·수식을 읽고 PDF로 저장할 수 있습니다.</p>',
        '<p class="warning"><b>남은 차이:</b> 긴 표준취업규칙은 폴라리스 163쪽, Studio 154쪽입니다. 대각선 표에서는 빨간 대각선·색이 일부 누락되며, 가는 표 테두리는 더 희미합니다. 같은 쪽수인 문서도 글꼴·기준선·표 분할·수식 배치가 완전히 같지는 않습니다. 제출용 인쇄와 정확한 페이지 참조는 원본 뷰어에서도 확인하세요.</p>',
        '<table><tr><th>문서</th><th>기준</th><th>기준 쪽수</th><th>Studio 쪽수</th><th>확인</th></tr>']
for row in table_rows:
    body.append('<tr>' + ''.join('<td>' + html.escape(str(cell)) + '</td>' for cell in row) + '</tr>')
body.extend(['</table><p>폴라리스: 같은 공개 HWP/HWPX를 실제 서비스에서 열어 캡처. 한컴 PDF: 같은 이름으로 동봉된 참고본이며 Hwp 2022 / Hancom PDF 메타데이터 확인. 이 Mac에서 한컴 앱으로 새로 출력한 것은 아닙니다.</p>',
             '<p>아래 그림은 종이 부분만 잘라 같은 크기로 맞췄습니다. 원본 전체 화면과 파일 SHA-256, 글자 영역의 근사 측정은 검증 폴더에 보존했습니다. 측정치를 호환성 정확도 점수로 사용하지 않습니다.</p>'])
for row in comparisons:
    encoded = base64.b64encode((BASE / row['image']).read_bytes()).decode()
    title = f"{row['file']} · {row['viewer']} {row['referencePage']}쪽 / Studio {row['studioPage']}쪽"
    body.append(f'<details><summary>{html.escape(title)}</summary><img alt="{html.escape(title)}" src="data:image/png;base64,{encoded}"><small>글자 영역(mm): 기준 {row["referenceGeometry"]["inkBoundsMm"]} / Studio {row["studioGeometry"]["inkBoundsMm"]}</small></details>')
body.append('<h2>기능 검사</h2><p>Windows·Ubuntu 22.04에서 전체 브라우저 검사 88개, 네이티브 단위 검사 4개가 각각 통과했습니다. Mac 패키지와 Windows·Ubuntu 22.04/24.04 설치 앱에서 실제 파일 열기·전체 텍스트 복사·PDF 저장·기존 Word 문서 보존을 확인했습니다. PDF는 합성 HWP/HWPX 각각 1쪽, Mac 공개 근로계약서 7쪽, 보건소 표 4쪽으로 화면과 일치합니다.</p><p>중첩 표·글상자 내용 누락을 보완하여 긴 취업규칙의 텍스트 복사가 5,901자에서 129,913자로 늘었습니다. 파일은 기기 안에서 처리하며 테스트의 외부 요청과 렌더러 오류는 0개입니다.</p>')
body.append('<h2>출처</h2><ul><li><a href="https://www.moel.go.kr/info/etc/dataroom/view.do?bbs_seq=20250300356">고용노동부 공식 파일</a></li><li><a href="https://hwp.polarisoffice.com/">폴라리스 웹 뷰어</a></li><li><a href="https://github.com/edwardkim/rhwp">RHWP 공개 샘플·해석기 (0.8.7, MIT)</a></li><li><a href="https://github.com/beaur8-prog/word-studio/blob/main/HWP_VALIDATION.md">전체 검증·제한·재현 방법</a></li></ul></html>')
(BASE / 'report.html').write_text('\n'.join(body))

# One final PDF deliverable. Raw captures and native app PDFs are test evidence.
pdf_path = ROOT / 'output/pdf/HWP-Layout-Comparison.pdf'
pdf_path.parent.mkdir(parents=True, exist_ok=True)
pdf = canvas.Canvas(str(pdf_path), pagesize=landscape(A4))
pdf.setTitle('HWP/HWPX 읽기 및 레이아웃 비교 — Word Studio 1.3.0')
W, H = landscape(A4)
def text(x, y, value, size=10, color='#25354c'):
    pdf.setFillColor(color); pdf.setFont('ReportKR', size); pdf.drawString(x, y, value)

text(36, H - 46, 'HWP/HWPX 읽기 · 레이아웃 비교', 22)
text(36, H - 69, 'Word Studio 1.3.0  |  2026-10-20  |  소스·로컬 검증본, 공개 배포 전', 10)
text(36, H - 97, '공개 문서 11개 열기·화면 37쪽 검사 / 구형 문서 1개 거부 / 폴라리스 문서 6종 + 한컴 PDF 참고본 3종', 10)
y = H - 131
columns = [36, 224, 418, 495, 570]
for x, value in zip(columns, ['문서', '비교 기준', '기준 쪽수', 'Studio', '비교 범위']): text(x, y, value, 10)
pdf.setStrokeColorRGB(.78, .82, .88); pdf.line(36, y - 10, W - 36, y - 10)
for row in table_rows:
    y -= 26
    for x, value in zip(columns, row): text(x, y, str(value), 9)
text(36, 190, '주요 차이와 해석', 13)
for i, value in enumerate([
    '• 긴 취업규칙은 163 / 154쪽 차이가 남습니다. 같은 쪽수도 동일한 글꼴·줄바꿈·조판을 뜻하지 않습니다.',
    '• 대각선 표의 빨간 선·색 일부 누락과 가는 테두리의 흐린 표시가 남아 있습니다.',
    '• 원본 HWP를 편집하지 않습니다. 기존 Word 문서를 유지하고, 읽기 화면을 닫으면 편집에 돌아갑니다.',
    '• 중첩 표·글상자 텍스트 누락과 여러 SVG 쪽의 표 잘림 자원 충돌을 보완했습니다.',
    '• Mac 실제 앱 PDF: 합성 HWP/HWPX 1쪽씩, 근로계약서 7쪽, 보건소 표 4쪽. 렌더러 오류 0개.',
    '• Windows·Ubuntu: 브라우저 검사 88개씩 통과. 실제 설치 앱의 HWP/HWPX 열기·복사·PDF도 통과.',
    '• 폴라리스는 실제 서비스 화면, 한컴 PDF는 동봉 참고본입니다. 이 Mac에서 한컴 앱을 직접 실행하지 않았습니다.',
    '• 원본 전체 화면·공개 URL·SHA-256·전체 비교는 로컬 report.html 및 Git의 HWP_VALIDATION.md에 있습니다.',
]): text(36, 165 - i * 17, value, 9)
text(36, 28, 'https://github.com/beaur8-prog/word-studio/blob/main/HWP_VALIDATION.md', 8)
pdf.showPage()
selection = [('01-employment-contract.hwp', 'Polaris', 1), ('03-health-duties.hwp', 'Polaris', 1),
             ('06-image.hwp', 'Polaris', 3), ('08-equation.hwp', 'Polaris', 1), ('09-diagonal.hwpx', 'Polaris', 1),
             ('02-work-rules.hwp', 'Polaris', 1), ('03-health-duties.hwp', 'Hancom PDF', 4),
             ('04-language-plan.hwp', 'Hancom PDF', 10), ('10-text-align.hwp', 'Hancom PDF', 1)]
for file, viewer, number in selection:
    row = next(row for row in comparisons if row['file'] == file and row['viewer'] == viewer and row['studioPage'] == number)
    pdf.drawImage(ImageReader(str(BASE / row['image'])), 38, 22, width=W - 76, height=H - 44, preserveAspectRatio=True, anchor='c')
    pdf.showPage()
pdf.save()
print(json.dumps({'comparisons': len(comparisons), 'html': str(BASE / 'report.html'), 'pdf': str(pdf_path)}, ensure_ascii=False))
