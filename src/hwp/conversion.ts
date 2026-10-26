import type { HwpDocument } from '@rhwp/core';
import type { PageSettings } from '../types';
import documentLimits from '../../electron/document-limits.json';
import type { HwpProgressReporter } from './progress';
import type { HwpLayoutMetadata } from './layout-metadata';

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
const cssStyle = (tag: string, css: string) => tag.includes('style="') ? tag.replace(/style="([^"]*)"/, (_, old) => `style="${old};${css}"`) : tag.replace(/>$/, ` style="${css}">`);

function paragraphCss(properties: Record<string, any>, char?: Record<string, any>, cap?: number, includeSpacing = true): string {
  const fontPx = Number(char?.fontSize) / 75;
  const line = Number(properties.lineSpacing);
  // The query API mixes units: margins/indent and fixed leading are
  // already physical pixels, but before/after retain ParaShape's 2x scale.
  // See build_para_properties_json and style_resolver at the corpus revision.
  const height = properties.lineSpacingType === 'Percent' ? fontPx * line / 100
    : properties.lineSpacingType === 'Fixed' ? line
    : properties.lineSpacingType === 'SpaceOnly' ? fontPx + line
    : properties.lineSpacingType === 'Minimum' ? Math.max(fontPx, line) : NaN;
  const indent = Number(properties.indent) || 0;
  const left = (Number(properties.marginLeft) || 0) + Math.max(0, -indent);
  const spacing = Number.isFinite(cap) && Number.isFinite(height) && height > cap! && cap! >= fontPx ? `${cap}px`
    : properties.lineSpacingType === 'Percent' ? `${line}%` : Number.isFinite(height) ? `${height}px` : '';
  return `margin-left:${left}px;margin-right:${(Number(properties.marginRight) || 0)}px;text-indent:${indent}px;` +
    (includeSpacing ? `margin-top:${(Number(properties.spacingBefore) || 0) / 2}px;margin-bottom:${(Number(properties.spacingAfter) || 0) / 2}px;` : '') +
    (char?.fontSize ? `font-size:${char.fontSize / 100}pt;font-family:${escape(JSON.stringify(char.fontFamily || 'serif'))};` : '') +
    (spacing ? `line-height:${spacing};` : '');
}

function sourceLinePitches(doc: HwpDocument): Map<string, number> {
  const lines = new Map<string, { y: number; h: number; page: number }[]>();
  const pitches = new Map<string, number>();
  for (let page = 0; page < doc.pageCount(); page++) for (const run of JSON.parse(doc.getPageTextLayout(page)).runs ?? []) {
    if (run.parentParaIdx === undefined && typeof run.secIdx === 'number' && typeof run.paraIdx === 'number') pitches.set(`flow:${run.secIdx}:${run.paraIdx}`, 1);
    if (typeof run.secIdx !== 'number' || typeof run.paraIdx !== 'number' || !run.text?.trim() || run.cellPath?.length > 1) continue;
    const key = run.parentParaIdx === undefined ? `${run.secIdx}:${run.paraIdx}` : `${run.secIdx}:${run.parentParaIdx}:${run.controlIdx}:${run.cellIdx}:${run.cellParaIdx}`;
    const group = lines.get(key) ?? [], previous = group.at(-1);
    if (!previous || previous.page !== page || Math.abs(previous.y - run.y) > Math.min(previous.h, run.h) * .45) group.push({ y: run.y, h: run.h, page });
    lines.set(key, group);
  }
  for (const [key, group] of lines) {
    const values = group.slice(1).flatMap((line, i) => {
      const previous = group[i], value = line.y - previous.y;
      return line.page === previous.page && value >= Math.max(line.h, previous.h) * .9 && value < Math.max(line.h, previous.h) * 4 ? [value] : [];
    }).sort((a, b) => a - b);
    if (values.length) pitches.set(key, Math.round(values[Math.floor(values.length / 2)] * 100) / 100);
  }
  return pitches;
}

function tableGeometry(doc: HwpDocument, section: number, para: number, control: number, html: string, pitches: Map<string, number>): string {
  const dims = JSON.parse(doc.getTableDimensions(section, para, control));
  if (dims.cellCount > 50000 || dims.colCount > 1000 || dims.rowCount > 10000) throw new Error('편집할 표의 셀 수가 너무 많습니다.');
  const cells = Array.from({ length: dims.cellCount }, (_, index) => ({
    index, ...JSON.parse(doc.getCellInfo(section, para, control, index)), ...JSON.parse(doc.getCellProperties(section, para, control, index)),
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
  const tableProperties = JSON.parse(doc.getTableProperties(section, para, control));
  const borderWidths = [.1, .12, .15, .2, .25, .3, .4, .5, .6, .7, 1, 1.5, 2, 3, 4, 5];
  const style = (tag: string, css: string) => tag.includes('style="') ? tag.replace(/style="([^"]*)"/, (_, old) => `style="${old}${css}"`) : tag.replace(/>$/, ` style="${css}">`);
  let depth = 0, cellIndex = 0, row = 0, cellParagraph = 0, header = false, paragraphStyle = '';
  return html.replace(/<\/?(?:table|tr|td|p)\b[^>]*>/g, tag => {
    if (/^<table\b/.test(tag)) { depth++; return depth === 1 ? style(tag, `width:${columns.reduce((sum, n) => sum + n, 0)}px;`) : tag; }
    if (tag === '</table>') { depth--; return tag; }
    if (depth !== 1) return tag;
    if (/^<p\b/.test(tag)) {
      try {
        const c = cells[cellIndex - 1], index = c.index;
        const properties = JSON.parse(doc.getCellParaPropertiesAt(section, para, control, index, cellParagraph));
        const char = JSON.parse(doc.getCellCharPropertiesAt(section, para, control, index, cellParagraph, 0));
        const pitch = pitches.get(`${section}:${para}:${control}:${index}:${cellParagraph}`);
        cellParagraph++;
        // A one-line source cell can be shorter than CSS's em * percentage.
        // Preserve its measured content pitch, without enforcing a fixed row.
        let cap: number | undefined;
        if (c.rowSpan === 1 && doc.getCellParagraphCount(section, para, control, index) === 1 &&
          JSON.parse(doc.getLineInfoInCell(section, para, control, index, 0, 0)).lineCount === 1)
          cap = px(c.height - c.paddingTop - c.paddingBottom);
        return cssStyle(tag, paragraphCss(properties, char, cap, false) + (pitch ? `line-height:${pitch}px;` : ''));
      } catch { return tag; }
    }
    if (tag === '</td>') return `${cellParagraph ? '' : `<p style="${paragraphStyle}"></p>`}</${header ? 'th' : 'td'}>`;
    if (/^<tr\b/.test(tag)) return style(tag, `height:${rows[row++] ?? 1}px;`);
    if (!/^<td\b/.test(tag)) return tag;
    const c = cells[cellIndex++]; if (!c) return tag;
    cellParagraph = 0; header = !!c.isHeader && !!tableProperties.repeatHeader;
    try { paragraphStyle = paragraphCss(JSON.parse(doc.getCellParaPropertiesAt(section, para, control, c.index, 0)), JSON.parse(doc.getCellCharPropertiesAt(section, para, control, c.index, 0, 0)),
      c.rowSpan === 1 ? px(c.height - c.paddingTop - c.paddingBottom) : undefined, false); }
    catch { paragraphStyle = 'line-height:100%;font-size:10pt;'; }
    const width = columns.slice(c.col, c.col + c.colSpan).map(n => Math.max(1, Math.round(n)));
    const borders = ['Left', 'Right', 'Top', 'Bottom'].map(side => {
      const b = c['border' + side], kind = !b?.type ? 'none' : b.type >= 8 && b.type <= 11 ? 'double' : b.type === 3 || b.type === 7 ? 'dotted' : b.type >= 2 && b.type <= 6 ? 'dashed' : 'solid';
      return `border-${side.toLowerCase()}:${borderWidths[b?.width] ?? .1}mm ${kind} ${/^#[a-f0-9]{6}$/i.test(b?.color) ? b.color : '#000000'};`;
    }).join('');
    let noWrap = false;
    if (doc.getCellParagraphCount(section, para, control, c.index) === 1) {
      const text = doc.exportSelectionInCellHtml(section, para, control, c.index, 0, 0, 0, 0xffffffff).replace(/<[^>]*>|<!--.*?-->/g, '').trim();
      noWrap = /^\d+[.)]?$/.test(text);
    }
    return style(tag.replace(/^<td/, header ? '<th' : '<td').replace(/>$/, ` colwidth="${width.join(',')}">`), `width:${px(c.width)}px;height:${px(c.height)}px;padding:${px(c.paddingTop)}px ${px(c.paddingRight)}px ${px(c.paddingBottom)}px ${px(c.paddingLeft)}px;vertical-align:${['top', 'middle', 'bottom'][c.verticalAlign] ?? 'top'};${noWrap ? 'white-space:nowrap;' : ''}${header ? 'background-color:transparent;color:#000;font-weight:normal;' : ''}${borders}`);
  });
}

function pageSettings(doc: HwpDocument, section: number, warnings: Set<string>): PageSettings {
  const p: PageDef = JSON.parse(doc.getPageDef(section));
  const letter = Math.abs(p.width - 61200) < 150 && Math.abs(p.height - 79200) < 150;
  const custom = !letter && (Math.abs(p.width - 59528) > 150 || Math.abs(p.height - 84188) > 150);
  const column = JSON.parse(doc.getColumnDef(section));
  const page: PageSettings = { size: custom ? 'Custom' : letter ? 'Letter' : 'A4', ...(custom ? { width: mm(p.width), height: mm(p.height) } : {}), orientation: p.landscape ? 'landscape' : 'portrait',
    ...(column.columnCount > 1 ? { columns: { count: column.columnCount, space: px(column.spacing), balance: false } } : {}),
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
export function editableHwp(doc: HwpDocument, onProgress?: HwpProgressReporter, metadata?: HwpLayoutMetadata): EditableHwp {
  onProgress?.({ label: '문서의 표·그림 구조를 확인하는 중…' });
  const warnings = new Set<string>(metadata?.warnings);
  const controls: Control[] = JSON.parse(doc.getControls());
  const pitches = sourceLinePitches(doc);
  let picturePositions: Map<string, any> | undefined;
  const overlays = new Map<number, any[]>();
  const pictureHtml = (section: number, para: number, control: Control): string | undefined => {
    let props: any;
    try { props = JSON.parse(doc.getPictureProperties(section, para, control.controlIndex)); } catch { return; }
    if (props.treatAsChar) return;
    if (!picturePositions) {
      picturePositions = new Map();
      for (let page = 0; page < doc.pageCount(); page++) for (const c of JSON.parse(doc.getPageControlLayout(page)).controls ?? [])
        if (c.type === 'image') picturePositions.set(`${c.secIdx}:${c.paraIdx}:${c.controlIdx}`, { ...c, page });
    }
    const position = picturePositions.get(`${section}:${para}:${control.controlIndex}`);
    if (!position) return;
    let images = overlays.get(position.page);
    if (!images) {
      const page = JSON.parse(doc.getPageOverlayImages(position.page));
      images = [...page.behind ?? [], ...page.front ?? []]; overlays.set(position.page, images);
    }
    const image = images.find(i => Math.abs(i.bbox.x - position.x) < 1 && Math.abs(i.bbox.y - position.y) < 1 && Math.abs(i.bbox.width - position.w) < 1);
    if (!image || !/^image\/(png|jpeg|bmp|gif|webp)$/.test(image.mime)) return;
    const geometry = { horizontalPosition: { relative: 'page', offset: Math.round(position.x * 9525) },
      verticalPosition: { relative: 'page', offset: Math.round(position.y * 9525) },
      wrap: { type: ['BehindText', 'InFrontOfText'].includes(props.textWrap) ? 0 : props.textWrap === 'TopAndBottom' ? 3 : 1 },
      behindDocument: props.textWrap === 'BehindText', allowOverlap: props.allowOverlap, layoutInCell: true };
    return `<img data-inline-image="true" data-image-floating="${escape(JSON.stringify(geometry))}" src="data:${image.mime};base64,${image.base64}" width="${position.w}" height="${position.h}" alt="${escape(props.description || '그림')}" />`;
  };
  const byParagraph = new Map<number, Control[]>();
  for (const c of controls) {
    if (c.list === 0) { const list = byParagraph.get(c.para) ?? []; list.push(c); byParagraph.set(c.para, list); }
    if (c.props.TreatAsChar === 0 && c.ctrlId === 'tbl') warnings.add('떠 있는 표는 본문 흐름 안으로 옮깁니다. 표의 앵커·겹침·캡션은 원본과 다를 수 있습니다.');
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
    let columnCount = column.columnCount ?? column.count ?? 1, columnSpace = px(column.spacing ?? 0);
    const hasColumns = columnCount > 1 || Object.entries(metadata?.paragraphs ?? {}).some(([key, p]) => key.startsWith(section + ':') && (p.columns?.count ?? 1) > 1);
    let columnOpen = false;
    const openColumns = () => { if (hasColumns && !columnOpen) { append(`<div data-column-count="${columnCount}" data-column-space="${columnSpace}" data-column-balance="false">`); columnOpen = true; } };
    const closeColumns = () => { if (columnOpen) { append('</div>'); columnOpen = false; } };
    openColumns();
    for (let para = 0; para < doc.getParagraphCount(section); para++, globalPara++) {
      const properties = JSON.parse(doc.getParaPropertiesAt(section, para));
      const structure = metadata?.paragraphs[`${section}:${para}`];
      if (structure?.columns) { closeColumns(); columnCount = structure.columns.count; columnSpace = structure.columns.space; openColumns(); }
      if (properties.pageBreakBefore || structure?.pageBreak) { closeColumns(); append('<div data-page-break="true"></div>'); openColumns(); }
      else if (structure?.columnBreak) append('<div data-column-break="true"></div>');
      const positions: number[] = JSON.parse(doc.getControlTextPositions(section, para));
      const range = (start: number, end: number) => fragment(doc.exportSelectionHtml(section, para, start, para, end));
      const whole = range(0, 0xffffffff);
      let open = whole.match(/^<p\b[^>]*>/)?.[0] ?? '<p style="">';
      open = cssStyle(open, paragraphCss(properties, JSON.parse(doc.getCharPropertiesAt(section, para, 0))));
      const pitch = pitches.get(`${section}:${para}`);
      if (pitch) open = cssStyle(open, `line-height:${pitch}px;`);
      const inline = (html: string) => html.replace(/^<p\b[^>]*>\s*/, '').replace(/<\/p>\s*$/, '');
      let cursor = 0, content = '', contentBytes = 0;
      const appendInline = (value: string) => { contentBytes += hwpHtmlBytes(value); checkHwpHtmlSize(size + contentBytes); content += value; };
      for (const c of byParagraph.get(globalPara) ?? []) {
        if (['secd', 'cold', 'head', 'foot', 'pgnp', 'pgct', 'pghd', 'nwno'].includes(c.ctrlId)) continue;
        const pos = Math.max(cursor, positions[c.controlIndex] ?? cursor);
        appendInline(inline(range(cursor, pos))); cursor = pos;
        onProgress?.({ label: c.ctrlId === 'tbl' ? '표를 변환하는 중…' : c.ctrlId === 'gso' ? '그림·도형을 변환하는 중…' : '문서 개체를 변환하는 중…' });
        let controlHtml = fragment(doc.exportControlHtml(section, para, '[]', c.controlIndex));
        if (c.ctrlId === 'gso') controlHtml = pictureHtml(section, para, c) ?? controlHtml;
        if (c.ctrlId === 'tbl') controlHtml = tableGeometry(doc, section, para, c.controlIndex, controlHtml, pitches);
        if (c.ctrlId === 'eqed') {
          const equation = JSON.parse(doc.getEquationProperties(section, para, c.controlIndex, -1, -1));
          const script = equation.script || '';
          try {
            const svg = doc.renderEquationPreview(script, equation.fontSize || 1200, 0);
            const width = equation.width > 0 ? px(equation.width) : Number(svg.match(/\bwidth="([\d.]+)"/)?.[1]);
            const height = equation.height > 0 ? px(equation.height) : Number(svg.match(/\bheight="([\d.]+)"/)?.[1]);
            if (!(width > 0 && height > 0)) throw new Error('Empty equation preview');
            controlHtml = `<img data-inline-image="true" data-hwp-svg="${escape(svg)}" src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQIHWP4z8DwHwAFgAI/ScLbtAAAAABJRU5ErkJggg==" width="${width}" height="${height}" alt="${escape('수식: ' + script)}" />`;
            warnings.add('수식은 원래 모양의 그림으로 가져옵니다. 수식 명령어는 그림 설명에 보존하지만 수식 자체는 편집할 수 없습니다.');
          } catch { controlHtml = ''; }
        }
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
          const tableWidth = Number(controlHtml.match(/^<table[^>]*style="[^"]*width:([\d.]+)px/)?.[1]);
          const pd: PageDef = JSON.parse(doc.getPageDef(section));
          const fullWidth = columnCount > 1 && tableWidth > px((pd.landscape ? pd.height : pd.width) - pd.marginLeft - pd.marginRight) / columnCount * 1.35;
          if (fullWidth) { closeColumns(); append('<div data-column-count="1" data-column-space="0">'); }
          append(controlHtml);
          if (fullWidth) { append('</div>'); openColumns(); }
          content = ''; contentBytes = 0;
        } else appendInline(controlHtml.replace(/<img /g, '<img data-inline-image="true" '));
      }
      appendInline(inline(range(cursor, 0xffffffff)));
      // Keep real blank paragraphs, but avoid an extra empty paragraph after a table.
      if (content.trim() || !(byParagraph.get(globalPara) ?? []).some(c => c.ctrlId === 'tbl')) {
        if (content.includes('data-image-floating') && !content.replace(/<[^>]*>/g, '').trim() && !pitches.has(`flow:${section}:${para}`))
          open = cssStyle(open, 'line-height:1px;margin-top:0;margin-bottom:0;');
        append(open + content + '</p>');
      }
      onProgress?.({ label: '본문·표·그림을 변환하는 중…', completed: globalPara + 1, total, unit: '문단' });
    }
    closeColumns();
  }
  onProgress?.({ label: '변환한 본문과 그림을 정리하는 중…' });
  let html = pieces.join('\n').replace(/<div data-column-count="[^"]*"[^>]*>\s*<\/div>/g, '').replace(/<!--\s*rhwp:([\s\S]*?)-->/g, () => {
    warnings.add('표 안에 변환하지 못한 개체가 있습니다. 해당 위치의 표시와 원본을 확인하세요.');
    return '<span>[변환하지 못한 개체 — 원본 보기에서 확인]</span>';
  });
  // Verified Hancom glyph meanings, also used by the core reader's text paint.
  // Limit this to the two forms symbols; never guess arbitrary PUA meanings.
  html = html.replace(/\u{F081C}/gu, '┈').replace(/\u{F012B}/gu, '(인)');
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
  // HWP percentage leading is based on the character em. Word's AUTO leading
  // multiplies a font's full line box instead, so preserve the physical pitch
  // for these imported paragraphs. Ordinary authored HTML keeps its own rule.
  html = html.replace(/<p style="([^"]*)">/g, (tag, css: string) => {
    const line = [...css.matchAll(/(?:^|;)line-height:([\d.]+)%;/g)].at(-1);
    const font = [...css.matchAll(/(?:^|;)font-size:([\d.]+)(pt|px);/g)].at(-1);
    if (!line || !font || [...css.matchAll(/(?:^|;)line-height:([^;]+);/g)].at(-1)?.[1] !== line[1] + '%') return tag;
    const pitch = Number(font[1]) * (font[2] === 'pt' ? 4 / 3 : 1) * Number(line[1]) / 100;
    return cssStyle(tag, `line-height:${Math.max(.1, Math.round(pitch * 100) / 100)}px;`);
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
  doc.setPageDef(0, JSON.stringify({ width: page.size === 'Custom' ? hu(page.width!) : letter ? 61200 : 59528, height: page.size === 'Custom' ? hu(page.height!) : letter ? 79200 : 84188, landscape: page.orientation === 'landscape',
    // HWP body top = marginHeader + marginTop, unlike Word's body margin.
    marginTop: hu(m.top - header), marginRight: hu(m.right), marginBottom: hu(m.bottom - footer), marginLeft: hu(m.left),
    marginHeader: hu(header), marginFooter: hu(footer), marginGutter: hu(page.gutter ?? 0) }));
  if (page.columns) doc.setColumnDef(0, page.columns.count, 0, 1, Math.round(page.columns.space * 75));
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
