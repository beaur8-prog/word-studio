import { generateHTML } from '@tiptap/core';
import { editorExtensions, parseDocumentHtml } from '../extensions';
import { sanitizeImportedHtml } from '../document-html';
import type { StudioDocument } from '../types';
import { HwpReaderSession } from './reader';
import { checkHwpHtmlSize, hwpHtmlBytes, type EditableHwp } from './conversion';
import { createPrintLayout } from '../print-layout';
import { pageMargins, paperDimensions } from '../types';
import { paintHwpProgress, type HwpProgressReporter } from './progress';

function renderedBody(root: HTMLElement, doc: StudioDocument): string {
  // HTML does not carry editor CSS. Resolve the document defaults and heading
  // appearance before handing it to a parser that understands inline styles.
  const margins = pageMargins(doc.page);
  const [width] = paperDimensions(doc.page);
  root.className = 'tiptap';
  root.style.cssText = `position:fixed;visibility:hidden;pointer-events:none;left:-100000px;top:0;padding:0;min-height:0;width:${(width - margins.left - margins.right) * 96 / 25.4}px;`;
  document.body.append(root);
  try {
    for (const element of root.querySelectorAll<HTMLElement>('p,h1,h2,h3,h4,h5,h6')) {
      const computed = getComputedStyle(element), paragraph = document.createElement('p'), span = document.createElement('span');
      paragraph.style.cssText = element.style.cssText;
      for (const property of ['text-align', 'line-height', 'margin-top', 'margin-bottom', 'margin-left', 'margin-right', 'text-indent']) paragraph.style.setProperty(property, computed.getPropertyValue(property));
      // The engine's fixed-pixel line spacing uses a different internal scale.
      // Relative spacing preserves the measured CSS line box without halving it.
      const line = parseFloat(computed.lineHeight), font = parseFloat(computed.fontSize);
      if (line > 0 && font > 0) paragraph.style.lineHeight = `${Math.round(line / font * 100)}%`;
      for (const property of ['font-family', 'font-size', 'font-weight', 'font-style', 'color']) span.style.setProperty(property, computed.getPropertyValue(property));
      span.append(...element.childNodes); paragraph.append(span); element.replaceWith(paragraph);
    }
    // The Hancom parser reads cell dimensions, rather than Tiptap's grid attrs
    // and TR heights. Materialize the same grid in TD CSS before serialization.
    root.querySelectorAll<HTMLTableElement>('table').forEach(table => {
      const rows = [...table.rows].filter(row => row.closest('table') === table);
      rows.forEach((row, index) => [...row.cells].forEach(cell => {
        const widths = cell.getAttribute('colwidth')?.split(',').map(Number);
        if (widths?.length && widths.every(n => Number.isFinite(n) && n > 0)) cell.style.width = `${widths.reduce((sum, n) => sum + n, 0)}px`;
        else cell.style.width = `${cell.getBoundingClientRect().width}px`;
        const height = rows.slice(index, index + cell.rowSpan).reduce((sum, r) => sum + (parseFloat(r.style.height) || r.getBoundingClientRect().height), 0);
        if (height) cell.style.height = `${height}px`;
      }));
    });
    // The engine's CSS attribute reader does not decode HTML attribute entities.
    // Single quotes keep family names readable without changing document text.
    return root.innerHTML.replace(/style="([^"]*)"/g, (_, css: string) => `style="${css.replace(/&quot;/g, "'")}"`) || '<p></p>';
  } finally { root.remove(); }
}

export function importEditableHwp(source: EditableHwp, name: string): StudioDocument {
  checkHwpHtmlSize(hwpHtmlBytes(source.html));
  const clean = sanitizeImportedHtml(source.html);
  const root = document.createElement('div'); root.innerHTML = clean.html;
  // Engine HTML is pretty-printed. Its indentation/newlines between P/SPAN
  // tags are markup, not HWP text. Preserve actual run whitespace inside SPAN.
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT), remove: Text[] = [];
  while (walker.nextNode()) {
    const text = walker.currentNode as Text;
    if (/^[\t\r\n ]*[\r\n][\t\r\n ]*$/.test(text.data) && /^(P|TD|TH|TR|TABLE|TBODY|DIV)$/.test(text.parentElement?.tagName ?? '')) remove.push(text);
  }
  remove.forEach(text => text.remove());
  return { format: 'word-studio', version: 1, id: crypto.randomUUID(), title: name.replace(/\.(hwp|hwpx)$/i, ''),
    content: parseDocumentHtml(root.innerHTML), page: source.page, comments: [], updatedAt: new Date().toISOString() };
}

export async function previewEditableHwp(source: EditableHwp, name: string, onProgress?: HwpProgressReporter, signal?: AbortSignal): Promise<EditableHwp> {
  signal?.throwIfAborted();
  if (onProgress) { onProgress({ label: 'Word 편집 형식과 서식을 준비하는 중…' }); await paintHwpProgress(); signal?.throwIfAborted(); }
  const doc = importEditableHwp(source, name);
  const html = generateHTML(doc.content, editorExtensions);
  const layout = await createPrintLayout(html, doc.page, onProgress, signal);
  try { return { ...source, html, editorPages: layout.pages, warnings: [...new Set([...source.warnings, ...layout.warnings])] }; }
  finally { layout.dispose(); }
}

/** Explicitly inventory information that the shared HTML bridge cannot carry.
 * These warnings precede the save dialog, including when the writer reports no loss. */
export async function exportEditableHwp(doc: StudioDocument, format: 'hwp' | 'hwpx', onProgress?: HwpProgressReporter) {
  if (onProgress) { onProgress({ label: '저장할 본문과 서식을 준비하는 중…' }); await paintHwpProgress(); }
  const root = document.createElement('div');
  const source = generateHTML(doc.content, editorExtensions);
  checkHwpHtmlSize(hwpHtmlBytes(source));
  root.innerHTML = sanitizeImportedHtml(source).html;
  const warnings: string[] = [];
  const warn = (selector: string, message: string) => { if (root.querySelector(selector)) warnings.push(message); };
  if ([...root.querySelectorAll<HTMLElement>('[data-section-page]')].some(el => JSON.parse(el.dataset.sectionPage!).breakType === 'continuous')) warnings.push('같은 쪽의 연속 구역은 인쇄 쪽 단위로 저장합니다. 쪽 안에서 바뀌는 여백·머리말은 첫 구역 설정을 따를 수 있습니다.');
  const notes = [...root.querySelectorAll<HTMLElement>('[data-document-note]')].map(note => note.cloneNode(true) as HTMLElement);
  if (notes.some(note => !/^\d+$/.test(note.dataset.noteLabel || ''))) warnings.push('각주·미주의 사용자 지정 표시는 한글의 숫자 번호로 바뀝니다. 본문 참조와 각주·미주 내용은 유지합니다.');
  warn('[data-docx-bookmark], a', '책갈피·링크는 표시된 텍스트만 저장합니다.');
  warn('[data-image-floating]', '떠 있는 그림은 본문 흐름 안으로 옮깁니다. 자르기·회전·겹침은 유지되지 않습니다.');
  warn('[data-word-overlap]', '겹쳐 쓰기와 특수 문자는 표시된 텍스트로 저장합니다.');
  warn('table', '표의 병합과 내용은 저장하지만 복잡한 셀 너비·테두리·쪽 나눔은 달라질 수 있습니다.');
  warn('h1,h2,h3,h4,h5,h6', '제목은 글자·문단 서식으로 저장합니다. Word 제목 수준은 유지되지 않습니다.');
  if (doc.comments.length) warnings.push(`검토 메모 ${doc.comments.length}개는 한글 파일에 저장되지 않습니다. Word Studio 백업을 함께 보관하세요.`);
  if (doc.page.differentFirstPage || doc.page.differentOddEven || doc.page.headerHtml || doc.page.footerHtml || doc.page.pageNumbers) warnings.push('머리말·꼬리말은 기본 텍스트만 저장하며 자동 쪽 번호와 첫 쪽·홀짝 구분은 유지되지 않습니다.');
  const layout = await createPrintLayout(source, doc.page, onProgress);
  let html: string, exportPage = doc.page;
  const wordPages = layout.pages;
  try {
    if (onProgress) { onProgress({ label: '한글 문서의 쪽 나누기를 준비하는 중…' }); await paintHwpProgress(); }
    warnings.push(...layout.warnings);
    if (wordPages > 1) warnings.push('Word 인쇄 쪽 경계를 유지하도록 표·문단을 쪽 단위로 나누어 저장합니다. 이어진 표가 여러 표로 나뉘고 반복 머리행이 각 표에 포함될 수 있습니다.');
    // Preserve Word's measured page boundaries. Letting one enormous HTML table
    // reflow in the engine can collapse hundreds of pages into a few dozen.
    let previousPage = doc.page;
    root.innerHTML = [...layout.element.querySelectorAll('.print-page')].map((sheet, index) => {
      const body = sheet.querySelector('.print-content')?.innerHTML || '<p></p>';
      const settings = layout.pageSettings.get(sheet) ?? previousPage;
      if (!index) exportPage = settings;
      const section = JSON.stringify(settings) !== JSON.stringify(previousPage);
      const marker = document.createElement('div');
      if (section) marker.dataset.sectionPage = JSON.stringify(settings);
      else marker.dataset.pageBreak = 'true';
      previousPage = settings;
      return (index ? marker.outerHTML : '') + body + (sheet.querySelector('.print-footnotes')?.innerHTML || '');
    }).join('');
    root.innerHTML = sanitizeImportedHtml(root.innerHTML).html;
    // Print places footnotes at page bottoms. Save their complete source
    // stories once, including notes that continued onto another print page.
    root.querySelectorAll('[data-document-note]').forEach(note => note.remove());
    root.append(...notes);
    root.querySelectorAll('[data-docx-tab]').forEach(el => el.replaceWith(document.createTextNode('\t')));
    root.querySelectorAll('img').forEach(img => { img.removeAttribute('data-image-floating'); });
    html = renderedBody(root, doc);
  } finally { layout.dispose(); }
  const session = new HwpReaderSession();
  try {
    const result = await session.export(html, exportPage, format, onProgress);
    warnings.push(`쪽 수 비교: Word 인쇄 ${wordPages}쪽 → ${format.toUpperCase()} 원본 보기 ${result.pages}쪽.` + (wordPages !== result.pages ? ' 쪽 배치가 달라졌습니다. 정확한 쪽 참조가 필요한 문서는 원본과 대조하세요.' : ' 쪽 수가 같아도 세부 배치는 다를 수 있습니다.'));
    return { blob: new Blob([new Uint8Array(result.bytes)], { type: format === 'hwp' ? 'application/x-hwp' : 'application/vnd.hancom.hwpx' }), warnings: [...new Set([...warnings, ...result.warnings])] };
  } finally { session.close(); }
}
