import type { HwpDocument } from '@rhwp/core';
import type { PageSettings } from '../types';
import documentLimits from '../../electron/document-limits.json';
import type { HwpProgressReporter } from './progress';

export interface EditableHwp { html: string; page: PageSettings; warnings: string[]; editorPages?: number }
export const HWP_CONVERSION_NOTICE = '본문, 기본 글자·문단 서식, 표와 그림을 변환합니다. 한글과 Word의 조판 방식이 달라 줄바꿈과 쪽 수가 달라질 수 있습니다. 원본 파일은 유지하고 변환한 문서는 새 파일로 저장하세요.';
// Expanded HTML includes base64 pictures, so a 21 MB image document already
// needs about 28 MB before markup. This is separate from the source-file cap.
export const HWP_HTML_LIMIT = documentLimits.hwpEditableHtmlBytes;
export function hwpHtmlBytes(value: string): number {
  // Do not allocate another document-sized buffer just to measure the Worker input.
  let bytes = value.length;
  for (let i = 0; i < value.length; i++) {
    const c = value.charCodeAt(i);
    if (c <= 0x7f) continue;
    if (c <= 0x7ff) bytes++;
    else {
      bytes += 2;
      if (c >= 0xd800 && c <= 0xdbff && i + 1 < value.length) {
        const next = value.charCodeAt(i + 1);
        if (next >= 0xdc00 && next <= 0xdfff) i++;
      }
    }
  }
  return bytes;
}
export function checkHwpHtmlSize(bytes: number) {
  if (bytes > HWP_HTML_LIMIT) throw new Error(`편집용 변환 데이터가 약 ${Math.ceil(bytes / 1024 / 1024 * 10) / 10} MiB로 한도 ${HWP_HTML_LIMIT / 1024 / 1024} MiB를 넘습니다. 원본 파일 크기가 아니라 압축을 푼 본문·그림을 편집 형식으로 변환한 크기입니다. 그림을 줄이거나 문서를 나눠 주세요.`);
}
const escape = (text: string) => text.replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]!);
const fragment = (html: string) => html.match(/<!--StartFragment-->([\s\S]*?)<!--EndFragment-->/)?.[1].trim() ?? html;
type Control = { list: number; para: number; controlIndex: number; ctrlId: string; userDesc: string; props: { TreatAsChar?: number } };
type PageDef = { width: number; height: number; landscape: boolean; marginTop: number; marginRight: number; marginBottom: number; marginLeft: number; marginHeader: number; marginFooter: number; marginGutter: number };
const mm = (value: number) => Math.round(value * 25.4 / 7200 * 100) / 100;
const hu = (value: number) => Math.round(value * 7200 / 25.4);
const px = (value: number) => Math.round(value / 75 * 100) / 100;

function tableGeometry(doc: HwpDocument, section: number, para: number, control: number, html: string): string {
  const dims = JSON.parse(doc.getTableDimensions(section, para, control));
  if (dims.cellCount > 50000 || dims.colCount > 1000 || dims.rowCount > 10000) throw new Error('편집할 표의 셀 수가 너무 많습니다.');
  const cells = Array.from({ length: dims.cellCount }, (_, index) => ({
    ...JSON.parse(doc.getCellInfo(section, para, control, index)), ...JSON.parse(doc.getCellProperties(section, para, control, index)),
  })).sort((a, b) => a.row - b.row || a.col - b.col);
  // Solve cumulative grid edges from both merged and unmerged cells. Dividing
  // every merged width equally expands the fine grids used by Hancom forms.
  const grid = (axis: 'col' | 'row', length: number) => {
    const edges: (number | undefined)[] = Array(length + 1); edges[0] = 0;
    for (let pass = 0; pass < 4; pass++) for (const cell of cells) {
      const start = cell[axis], end = start + cell[axis + 'Span'], size = cell[axis === 'col' ? 'width' : 'height'];
      if (edges[start] !== undefined && edges[end] === undefined) edges[end] = edges[start]! + size;
      else if (edges[end] !== undefined && edges[start] === undefined) edges[start] = Math.max(0, edges[end]! - size);
    }
    for (let i = 1; i <= length; i++) if (edges[i] === undefined) {
      let end = i + 1; while (end <= length && edges[end] === undefined) end++;
      const step = end <= length ? (edges[end]! - edges[i - 1]!) / (end - i + 1) : axis === 'col' ? 6000 : 1000;
      for (let j = i; j < end && j <= length; j++) edges[j] = edges[i - 1]! + step * (j - i + 1);
      i = end - 1;
    }
    return Array.from({ length }, (_, i) => Math.max(1, px(edges[i + 1]! - edges[i]!)));
  };
  const columns = grid('col', dims.colCount), rows = grid('row', dims.rowCount);
  const borderWidths = [.1, .12, .15, .2, .25, .3, .4, .5, .6, .7, 1, 1.5, 2, 3, 4, 5];
  const style = (tag: string, css: string) => tag.includes('style="') ? tag.replace(/style="([^"]*)"/, (_, old) => `style="${old}${css}"`) : tag.replace(/>$/, ` style="${css}">`);
  let depth = 0, cellIndex = 0, row = 0;
  return html.replace(/<\/?(?:table|tr|td)\b[^>]*>/g, tag => {
    if (/^<table\b/.test(tag)) { depth++; return depth === 1 ? style(tag, `width:${columns.reduce((sum, n) => sum + n, 0)}px;`) : tag; }
    if (tag === '</table>') { depth--; return tag; }
    if (depth !== 1) return tag;
    if (/^<tr\b/.test(tag)) return style(tag, `height:${rows[row++] ?? 1}px;`);
    if (!/^<td\b/.test(tag)) return tag;
    const c = cells[cellIndex++]; if (!c) return tag;
    const width = columns.slice(c.col, c.col + c.colSpan).map(n => Math.max(1, Math.round(n)));
    const borders = ['Left', 'Right', 'Top', 'Bottom'].map(side => {
      const b = c['border' + side], kind = !b?.type ? 'none' : b.type >= 8 && b.type <= 11 ? 'double' : b.type === 3 || b.type === 7 ? 'dotted' : b.type >= 2 && b.type <= 6 ? 'dashed' : 'solid';
      return `border-${side.toLowerCase()}:${borderWidths[b?.width] ?? .1}mm ${kind} ${/^#[a-f0-9]{6}$/i.test(b?.color) ? b.color : '#000000'};`;
    }).join('');
    return style(tag.replace(/>$/, ` colwidth="${width.join(',')}">`), `width:${px(c.width)}px;height:${px(c.height)}px;padding:${px(c.paddingTop)}px ${px(c.paddingRight)}px ${px(c.paddingBottom)}px ${px(c.paddingLeft)}px;vertical-align:${['top', 'middle', 'bottom'][c.verticalAlign] ?? 'top'};${borders}`);
  });
}

function pageSettings(doc: HwpDocument, section: number, warnings: Set<string>): PageSettings {
  const p: PageDef = JSON.parse(doc.getPageDef(section));
  const letter = Math.abs(p.width - 61200) < 150 && Math.abs(p.height - 79200) < 150;
  if (!letter && (Math.abs(p.width - 59528) > 150 || Math.abs(p.height - 84188) > 150)) warnings.add('A4/Letter 이외의 용지는 A4로 맞춥니다.');
  const page: PageSettings = { size: letter ? 'Letter' : 'A4', orientation: p.landscape ? 'landscape' : 'portrait',
    margin: mm(p.marginLeft), margins: { top: mm(p.marginTop + p.marginHeader), right: mm(p.marginRight), bottom: mm(p.marginBottom + p.marginFooter), left: mm(p.marginLeft) },
    header: '', footer: '', pageNumbers: false, headerDistance: mm(p.marginHeader), footerDistance: mm(p.marginFooter), gutter: mm(p.marginGutter) };
  for (const isHeader of [true, false]) {
    const hf = JSON.parse(doc.getHeaderFooter(section, isHeader, 0));
    if (hf.exists) {
      page[isHeader ? 'header' : 'footer'] = hf.text ?? '';
      warnings.add('머리말·꼬리말은 기본 텍스트로 가져옵니다. 홀짝·첫 쪽 구분과 개체는 원본 보기에서 확인하세요.');
    }
  }
  return page;
}

/** The engine's selection HTML intentionally omits controls. Interleave their
 * HTML at the raw paragraph offsets instead of losing tables and pictures. */
export function editableHwp(doc: HwpDocument, onProgress?: HwpProgressReporter): EditableHwp {
  onProgress?.({ label: '문서의 표·그림 구조를 확인하는 중…' });
  const warnings = new Set<string>();
  const controls: Control[] = JSON.parse(doc.getControls());
  const byParagraph = new Map<number, Control[]>();
  for (const c of controls) {
    if (c.list === 0) { const list = byParagraph.get(c.para) ?? []; list.push(c); byParagraph.set(c.para, list); }
    if (c.props.TreatAsChar === 0 && ['tbl', 'gso'].includes(c.ctrlId)) warnings.add('떠 있는 표·그림은 본문 흐름 안으로 옮깁니다. 겹침·회전·자르기·캡션은 원본과 다를 수 있습니다.');
  }
  const pieces: string[] = [];
  let globalPara = 0, size = 0;
  const append = (value: string) => { size += hwpHtmlBytes(value) + (pieces.length ? 1 : 0); checkHwpHtmlSize(size); pieces.push(value); };
  const firstPage = pageSettings(doc, 0, warnings);
  const total = Array.from({ length: doc.getSectionCount() }, (_, section) => doc.getParagraphCount(section)).reduce((sum, count) => sum + count, 0);
  onProgress?.({ label: '본문·표·그림을 변환하는 중…', completed: 0, total, unit: '문단' });
  for (let section = 0; section < doc.getSectionCount(); section++) {
    if (section) { const page = pageSettings(doc, section, warnings); append(`<div data-section-page="${escape(JSON.stringify(page))}"></div>`); }
    const column = JSON.parse(doc.getColumnDef(section));
    if ((column.columnCount ?? column.count ?? 1) > 1) warnings.add('한글의 다단 배치는 편집본에서 단일 본문으로 변환합니다.');
    for (let para = 0; para < doc.getParagraphCount(section); para++, globalPara++) {
      const properties = JSON.parse(doc.getParaPropertiesAt(section, para));
      if (properties.pageBreakBefore) append('<div data-page-break="true"></div>');
      const positions: number[] = JSON.parse(doc.getControlTextPositions(section, para));
      const range = (start: number, end: number) => fragment(doc.exportSelectionHtml(section, para, start, para, end));
      const whole = range(0, 0xffffffff);
      let open = whole.match(/^<p\b[^>]*>/)?.[0] ?? '<p style="">';
      open = open.replace(/style="([^"]*)"/, (_, css) => `style="${css}margin-top:${Number(properties.spacingBefore) || 0}px;margin-bottom:${Number(properties.spacingAfter) || 0}px;"`);
      const inline = (html: string) => html.replace(/^<p\b[^>]*>\s*/, '').replace(/<\/p>\s*$/, '');
      let cursor = 0, content = '', contentBytes = 0;
      const appendInline = (value: string) => { contentBytes += hwpHtmlBytes(value); checkHwpHtmlSize(size + contentBytes); content += value; };
      for (const c of byParagraph.get(globalPara) ?? []) {
        if (['secd', 'cold', 'head', 'foot', 'pgnp', 'pgct', 'pghd', 'nwno'].includes(c.ctrlId)) continue;
        const pos = Math.max(cursor, positions[c.controlIndex] ?? cursor);
        appendInline(inline(range(cursor, pos))); cursor = pos;
        onProgress?.({ label: c.ctrlId === 'tbl' ? '표를 변환하는 중…' : c.ctrlId === 'gso' ? '그림·도형을 변환하는 중…' : '문서 개체를 변환하는 중…' });
        let controlHtml = fragment(doc.exportControlHtml(section, para, '[]', c.controlIndex));
        if (c.ctrlId === 'tbl') controlHtml = tableGeometry(doc, section, para, c.controlIndex, controlHtml);
        if (!controlHtml || /^<!--/.test(controlHtml)) {
          let detail = '';
          if (c.ctrlId === 'eqed') detail = JSON.parse(doc.getEquationProperties(section, para, c.controlIndex, -1, -1)).script ?? '';
          controlHtml = `<span>[${escape(c.userDesc || c.ctrlId)}${detail ? ': ' + escape(detail) : ' — 원본 보기에서 확인'}]</span>`;
          if (c.ctrlId === 'gso') {
            try {
              const count = doc.getCellParagraphCount(section, para, c.controlIndex, 0);
              if (count) controlHtml = fragment(doc.exportSelectionInCellHtml(section, para, c.controlIndex, 0, 0, 0, count - 1, 0xffffffff));
            } catch { /* Non-textbox shapes retain a visible placeholder. */ }
          }
          warnings.add(`${c.userDesc || c.ctrlId} 개체는 편집 가능한 원본 개체로 변환하지 못해 위치에 텍스트 표시를 남깁니다.`);
        }
        if (/^<(?:table|p)\b/.test(controlHtml)) {
          if (content.trim()) append(open + content + '</p>');
          append(controlHtml); content = ''; contentBytes = 0;
        } else appendInline(controlHtml.replace(/<img /g, '<img data-inline-image="true" '));
      }
      appendInline(inline(range(cursor, 0xffffffff)));
      // Keep real blank paragraphs, but avoid an extra empty paragraph after a table.
      if (content.trim() || !(byParagraph.get(globalPara) ?? []).some(c => c.ctrlId === 'tbl')) append(open + content + '</p>');
      onProgress?.({ label: '본문·표·그림을 변환하는 중…', completed: globalPara + 1, total, unit: '문단' });
    }
  }
  onProgress?.({ label: '변환한 본문과 그림을 정리하는 중…' });
  let html = pieces.join('\n').replace(/<!--\s*rhwp:([\s\S]*?)-->/g, () => {
    warnings.add('표 안에 변환하지 못한 개체가 있습니다. 해당 위치의 표시와 원본을 확인하세요.');
    return '<span>[변환하지 못한 개체 — 원본 보기에서 확인]</span>';
  });
  // Some HWPX manifests label embedded BMPs as application/octet-stream.
  // Sniff only known raster signatures; never permit arbitrary data resources.
  html = html.replace(/src="data:application\/octet-stream;base64,([A-Za-z0-9+/=\s]+)"/g, (source, base64: string) => {
    const b = base64.replace(/\s/g, '');
    const mime = b.startsWith('Qk') ? 'bmp' : b.startsWith('iVBORw0KGgo') ? 'png' : b.startsWith('/9j/') ? 'jpeg' : b.startsWith('R0lGOD') ? 'gif' : '';
    return mime ? `src="data:image/${mime};base64,${b}"` : source;
  });
  // A paragraph's CSS strut otherwise inherits Word's 11pt font even when every
  // HWP run is 9pt, inflating line-height percentages throughout long tables.
  html = html.replace(/<p style="([^"]*)">(\s*<span style="([^"]*)">)/g, (_, css: string, span: string, runCss: string) => {
    const font = ['font-size', 'font-family'].map(key => runCss.match(new RegExp(`(?:^|;)${key}:([^;]+)`))?.[0]).filter(Boolean).join(';');
    return `<p style="${css}${font};">${span}`;
  });
  checkHwpHtmlSize(hwpHtmlBytes(html));
  warnings.add('복잡한 표 테두리·셀 크기, 글상자·수식·각주·필드 등 고급 서식은 완전히 호환되지 않습니다. 편집본과 원본을 비교하세요.');
  return { html: html || '<p></p>', page: firstPage, warnings: [...warnings] };
}

export function setHwpPage(doc: HwpDocument, page: PageSettings) {
  const m = page.margins ?? { top: page.margin, right: page.margin, bottom: page.margin, left: page.margin };
  const letter = page.size === 'Letter';
  const header = Math.min(m.top, page.headerDistance ?? Math.max(0, m.top / 2 - 3));
  const footer = Math.min(m.bottom, page.footerDistance ?? Math.max(0, m.bottom / 2 - 3));
  doc.setPageDef(0, JSON.stringify({ width: letter ? 61200 : 59528, height: letter ? 79200 : 84188, landscape: page.orientation === 'landscape',
    // HWP body top = marginHeader + marginTop, unlike Word's body margin.
    marginTop: hu(m.top - header), marginRight: hu(m.right), marginBottom: hu(m.bottom - footer), marginLeft: hu(m.left),
    marginHeader: hu(header), marginFooter: hu(footer), marginGutter: hu(page.gutter ?? 0) }));
  if (page.columns) doc.setColumnDef(0, page.columns.count, 0, 1, hu(page.columns.space));
  for (const isHeader of [true, false]) {
    const text = page[isHeader ? 'header' : 'footer'];
    if (text) { doc.createHeaderFooter(0, isHeader, 0); doc.insertTextInHeaderFooter(0, isHeader, 0, 0, 0, text); }
  }
}

export function serializeHwp(Core: typeof HwpDocument, html: string, page: PageSettings, format: 'hwp' | 'hwpx', onProgress?: HwpProgressReporter) {
  if (typeof html !== 'string' || !html) throw new Error('저장할 편집 문서가 없습니다.');
  checkHwpHtmlSize(hwpHtmlBytes(html));
  if (!['hwp', 'hwpx'].includes(format)) throw new Error('지원하지 않는 한글 저장 형식입니다.');
  const doc = Core.createEmpty();
  try {
    // A real Hancom template supplies valid HWP 5.1 DocInfo and default styles.
    doc.createBlankDocument();
    const breakToken = 'WORDSTUDIO-PAGE-BREAK-' + crypto.randomUUID();
    const converted = html.replace(/<div\b[^>]*data-page-break=["'][^"']*["'][^>]*>[\s\S]*?<\/div>/gi, `<p>${breakToken}</p>`);
    // The first pasted paragraph inherits the host's paragraph shape even on
    // the multi-paragraph path. Sacrifice a leading blank, then remove both
    // sentinels. The writer rebuilds section controls from the section model.
    onProgress?.({ label: '본문·표·그림을 한글 문서로 변환하는 중…' });
    const pasted = JSON.parse(doc.pasteHtml(0, 0, 0, '<p></p>' + converted + '<p></p>'));
    if (!pasted.ok) throw new Error('문서 내용을 한글 형식으로 변환하지 못했습니다.');
    onProgress?.({ label: '한글 문서의 쪽 나누기를 반영하는 중…' });
    const last = doc.getParagraphCount(0) - 1;
    if (last > 0 && doc.getParagraphLength(0, last) === 0) doc.deleteParagraph(0, last);
    if (doc.getParagraphCount(0) > 1) doc.deleteParagraph(0, 0);
    for (let p = 0; p < doc.getParagraphCount(0); p++) {
      if (doc.getTextRange(0, p, 0, 100) !== breakToken) continue;
      if (p + 1 < doc.getParagraphCount(0) && doc.getTextRange(0, p + 1, 0, 100) !== breakToken) doc.deleteParagraph(0, p);
      else doc.deleteText(0, p, 0, breakToken.length);
      doc.applyParaFormat(0, p, JSON.stringify({ pageBreakBefore: true }));
    }
    setHwpPage(doc, page);
    onProgress?.({ label: `${format.toUpperCase()} 파일을 만드는 중…` });
    const output = format === 'hwp' ? doc.exportHwpWithReport() : doc.exportHwpxWithReport();
    try {
      const report = JSON.parse(output.contentLoss());
      const bytes = output.takeBytes();
      if (bytes.byteLength > documentLimits.documentFileBytes) throw new Error(`변환한 한글 파일이 ${documentLimits.documentFileBytes / 1024 / 1024} MiB를 넘습니다.`);
      // A successful writer alone is insufficient: actually parse its result.
      onProgress?.({ label: '변환한 한글 파일을 다시 읽어 확인하는 중…' });
      const check = new Core(bytes);
      let pages = 0;
      try { pages = check.pageCount(); if (!check.getSectionCount() || !pages) throw new Error('저장한 한글 문서를 다시 읽지 못했습니다.'); }
      finally { check.free(); }
      return { bytes, pages, warnings: report.count ? [`한글 저장 엔진에서 내용 손실 ${report.count}건을 보고했습니다. ${JSON.stringify(report.losses).slice(0, 1600)}`] : [] };
    } finally { output.free(); }
  } finally { doc.free(); }
}
