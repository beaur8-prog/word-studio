import { generateHTML } from '@tiptap/core';
import { editorExtensions, parseDocumentHtml } from '../extensions';
import { sanitizeImportedHtml } from '../document-html';
import type { StudioDocument } from '../types';
import { HwpReaderSession } from './reader';
import { checkHwpHtmlSize, hwpHtmlBytes, type EditableHwp } from './conversion';
import { createPrintLayout } from '../print-layout';
import { pageMargins, paperDimensions } from '../types';
import { imageFloating } from '../image-floating';
import { parseTabStops } from '../document-tabs';
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
      if (element.dataset.tabStops) paragraph.dataset.tabStops = element.dataset.tabStops;
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
  const textNodes: Text[] = [], tabs = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  while (tabs.nextNode()) if (tabs.currentNode.textContent?.includes('\t') && tabs.currentNode.parentElement?.closest('[data-tab-stops]')) textNodes.push(tabs.currentNode as Text);
  for (const node of textNodes) {
    const parts = node.data.split('\t'), fragment = document.createDocumentFragment();
    parts.forEach((part, i) => { if (i) { const tab = document.createElement('span'); tab.dataset.docxTab = 'true'; tab.textContent = '\t'; fragment.append(tab); } fragment.append(document.createTextNode(part)); });
    node.replaceWith(fragment);
  }
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
  warn('a', '링크는 표시된 텍스트만 저장합니다. 책갈피 위치는 유지합니다.');
  if ([...root.querySelectorAll<HTMLElement>('[data-tab-stops]')].some(p => parseTabStops(p.dataset.tabStops).some(s => s.position < 0)))
    warnings.push('한글 형식에는 음수 탭 위치를 저장하지 않습니다. 양수 탭 위치와 본문은 유지합니다. 원래 탭 설정은 DOCX 또는 백업으로 보관하세요.');
  warn('[data-image-floating]', '떠 있는 그림은 현재 인쇄 위치를 용지 기준 좌표로 저장합니다. 원래 문단·단에 연결된 이동 규칙은 바뀝니다.');
  if ([...root.querySelectorAll<HTMLElement>('[data-image-floating]')].some(el => { const g = imageFloating(el.dataset.imageFloating); return g?.polygon || g?.wrap?.type === 2 || g?.wrap?.type === 4; })) warnings.push('그림 윤곽을 따르는 감싸기는 사각형 감싸기로 저장합니다. 표시된 그림과 쪽 위치는 유지합니다.');
  warn('[data-word-overlap]', '겹쳐 쓰기와 특수 문자는 표시된 텍스트로 저장합니다.');
  warn('table', '표의 병합과 내용은 저장하지만 복잡한 셀 너비·테두리·쪽 나눔은 달라질 수 있습니다.');
  warn('h1,h2,h3,h4,h5,h6', '제목은 글자·문단 서식으로 저장합니다. Word 제목 수준은 유지되지 않습니다.');
  if (doc.comments.length) warnings.push(`검토 메모 ${doc.comments.length}개는 한글 파일에 저장되지 않습니다. Word Studio 백업을 함께 보관하세요.`);
  const allSettings = [doc.page, ...[...root.querySelectorAll<HTMLElement>('[data-section-page]')].map(el => JSON.parse(el.dataset.sectionPage!))];
  if (allSettings.some(page => page.differentFirstPage)) warnings.push('첫 쪽만 다른 머리말·꼬리말은 별도의 한글 구역으로 저장해 현재 쪽의 모양을 유지합니다.');
  if (allSettings.some(page => [page.headerHtml, page.footerHtml, page.evenHeaderHtml, page.evenFooterHtml].some(html => html?.includes('data-image-floating')))) warnings.push('머리말·꼬리말의 그림은 저장하지만 떠 있는 개체의 좌표·겹침은 달라질 수 있습니다.');
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
      const paper = sheet.getBoundingClientRect();
      for (const image of sheet.querySelectorAll<HTMLElement>('.print-content img[data-image-floating]')) {
        const g = imageFloating(image.dataset.imageFloating); if (!g) continue;
        const box = image.getBoundingClientRect();
        // Resolve paragraph/column/margin anchors on the printed sheet before
        // its body is detached. Native page coordinates also survive section
        // sizes and the explicit page fragments used by this writer.
        image.dataset.imageFloating = JSON.stringify({ ...g,
          horizontalPosition: { relative: 'page', offset: Math.round((box.left - paper.left) * 9525) },
          verticalPosition: { relative: 'page', offset: Math.round((box.top - paper.top) * 9525) } });
      }
      const body = sheet.querySelector('.print-content')?.innerHTML || '<p></p>';
      const original = layout.pageSettings.get(sheet) ?? previousPage;
      const first = sheet.getAttribute('data-section-first') === 'true';
      const settings = original.differentFirstPage ? first ? { ...original, differentFirstPage: false, differentOddEven: false,
        header: original.firstHeader ?? '', footer: original.firstFooter ?? '', headerHtml: original.firstHeaderHtml, footerHtml: original.firstFooterHtml, pageNumbers: original.firstPageNumbers ?? false }
        : { ...original, differentFirstPage: false } : original;
      if (!index) exportPage = settings;
      const section = JSON.stringify(settings) !== JSON.stringify(previousPage);
      const marker = document.createElement('div');
      if (section) marker.dataset.sectionPage = JSON.stringify(settings);
      else marker.dataset.pageBreak = 'true';
      previousPage = settings;
      return (index ? marker.outerHTML : '') + body + (sheet.querySelector('.print-footnotes')?.innerHTML || '');
    }).join('');
    root.querySelectorAll('[data-flow-widget]').forEach(widget => widget.remove());
    root.innerHTML = sanitizeImportedHtml(root.innerHTML).html;
    // Print places footnotes at page bottoms. Save their complete source
    // stories once, including notes that continued onto another print page.
    root.querySelectorAll('[data-document-note]').forEach(note => note.remove());
    root.append(...notes);
    root.querySelectorAll('[data-docx-tab]').forEach(el => el.replaceWith(document.createTextNode('\t')));
    html = renderedBody(root, doc);
  } finally { layout.dispose(); }
  const session = new HwpReaderSession();
  try {
    const result = await session.export(html, exportPage, format, onProgress);
    warnings.push(`쪽 수 비교: Word 인쇄 ${wordPages}쪽 → ${format.toUpperCase()} 원본 보기 ${result.pages}쪽.` + (wordPages !== result.pages ? ' 쪽 배치가 달라졌습니다. 정확한 쪽 참조가 필요한 문서는 원본과 대조하세요.' : ' 쪽 수가 같아도 세부 배치는 다를 수 있습니다.'));
    return { blob: new Blob([new Uint8Array(result.bytes)], { type: format === 'hwp' ? 'application/x-hwp' : 'application/vnd.hancom.hwpx' }), warnings: [...new Set([...warnings, ...result.warnings])] };
  } finally { session.close(); }
}
