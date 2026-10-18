import { pageMargins, type PageSettings } from './types';

export type PrintLayout = { element: HTMLElement; pages: number; warnings: string[]; dispose(): void };
const PX = 96 / 25.4;
const dimensions = (p: PageSettings) => {
  const size = p.size === 'Letter' ? [215.9, 279.4] : [210, 297];
  return p.orientation === 'landscape' ? size.reverse() : size;
};
function splitText(element: HTMLElement, offset: number): [HTMLElement, HTMLElement] {
  const walker = document.createTreeWalker(element, NodeFilter.SHOW_TEXT);
  let remaining = offset, node: Node | null = walker.nextNode();
  while (node && remaining > (node.textContent?.length ?? 0)) {
    remaining -= node.textContent!.length; node = walker.nextNode();
  }
  const before = document.createRange(), after = document.createRange();
  before.selectNodeContents(element); after.selectNodeContents(element);
  if (node) { before.setEnd(node, remaining); after.setStart(node, remaining); }
  else after.collapse(false);
  const a = element.cloneNode(false) as HTMLElement, b = element.cloneNode(false) as HTMLElement;
  a.append(before.cloneContents()); b.append(after.cloneContents());
  a.style.marginBottom = '0'; b.style.marginTop = '0'; b.style.textIndent = '0';
  return [a, b];
}

/** Build one measured page tree used unchanged by preview, PDF and native printing. */
export async function createPrintLayout(html: string, initialPage: PageSettings): Promise<PrintLayout> {
  const host = document.createElement('div');
  host.className = 'print-layout measuring';
  host.setAttribute('aria-label', '인쇄 미리보기 문서');
  document.body.append(host);
  const style = document.createElement('style'); host.append(style);
  const source = document.createElement('div'); source.innerHTML = html;
  const warnings = new Set<string>();
  let settings = initialPage, pageNumber = 0, content: HTMLElement;
  let sectionNumber = 0;
  const pageStyles: string[] = [];
  function newPage() {
    const [width, height] = dimensions(settings), margins = pageMargins(settings);
    const pageName = `studioSection${sectionNumber}`;
    pageStyles.push(`@page ${pageName} { size:${width}mm ${height}mm; margin:0; }`);
    style.textContent = [...new Set(pageStyles)].join('\n');
    const paper = document.createElement('section'); paper.className = 'print-page';
    paper.dataset.page = String(++pageNumber);
    Object.assign(paper.style, { width: `${width * PX}px`, height: `${height * PX}px`, page: pageName });
    content = document.createElement('div'); content.className = 'tiptap print-content';
    Object.assign(content.style, { left: `${margins.left * PX}px`, top: `${margins.top * PX}px`,
      width: `${(width - margins.left - margins.right) * PX}px`, height: `${(height - margins.top - margins.bottom) * PX}px` });
    const header = document.createElement('div'); header.className = 'print-header'; header.textContent = settings.header;
    const footer = document.createElement('div'); footer.className = 'print-footer';
    const footerText = document.createElement('span'); footerText.textContent = settings.footer; footer.append(footerText);
    if (settings.pageNumbers) { const number = document.createElement('span'); number.textContent = String(pageNumber); footer.append(number); }
    for (const el of [header, footer]) { el.style.left = `${margins.left * PX}px`; el.style.right = `${margins.right * PX}px`; }
    header.style.top = `${Math.max(0, margins.top / 2 - 3) * PX}px`;
    footer.style.bottom = `${Math.max(0, margins.bottom / 2 - 3) * PX}px`;
    header.style.maxHeight = `${Math.max(0, margins.top - 4) * PX}px`;
    footer.style.maxHeight = `${Math.max(0, margins.bottom - 4) * PX}px`;
    paper.append(header, content, footer); host.append(paper);
  }
  function fits() {
    const last = content.lastElementChild as HTMLElement | null;
    return !last || last.getBoundingClientRect().bottom <= content.getBoundingClientRect().bottom + 0.5;
  }
  function scaleOversized(element: HTMLElement) {
    const available = content.getBoundingClientRect().bottom - element.getBoundingClientRect().top;
    const height = element.getBoundingClientRect().height;
    if (height > available && available > 0) {
      element.style.zoom = String(Math.min(1, available / height));
      warnings.add('한 페이지보다 큰 그림·병합 셀 묶음은 잘리지 않도록 축소했습니다.');
    }
  }
  // Split editable form rows at paragraph or nested-table row boundaries. A
  // large single cell must not shrink all its text just to fit one sheet.
  function splitElement(element: HTMLElement, boundary: number, depth = 0): HTMLElement | null {
    if (depth > 30) return null;
    if (element.matches('p,h1,h2,h3,h4,h5,h6') && (element.textContent?.length ?? 0) > 1) {
      const offsets = [...new Intl.Segmenter(undefined, { granularity: 'grapheme' }).segment(element.textContent!)].map(s => s.index).slice(1);
      let low = 0, high = offsets.length - 1, best = 0;
      const marker = document.createComment('split'); element.replaceWith(marker);
      while (low <= high) {
        const mid = (low + high) >> 1, [head] = splitText(element, offsets[mid]); marker.before(head);
        if (head.getBoundingClientRect().bottom <= boundary + .5) { best = offsets[mid]; low = mid + 1; } else high = mid - 1;
        head.remove();
      }
      marker.replaceWith(element);
      if (!best) return null;
      const [head, tail] = splitText(element, best); element.replaceChildren(...head.childNodes); element.style.marginBottom = '0'; return tail;
    }
    if (element.matches('tr') && (element as HTMLTableRowElement).cells.length !== 1) return null;
    if (element.matches('table')) {
      const table = element as HTMLTableElement;
      if ([...table.rows].some(row => [...row.cells].some(cell => cell.rowSpan > 1))) return null;
    }
    const children = [...element.children] as HTMLElement[];
    const style = getComputedStyle(element);
    const innerBoundary = boundary - (parseFloat(style.paddingBottom) || 0) - (parseFloat(style.borderBottomWidth) || 0);
    const index = children.findIndex(child => !child.matches('colgroup') && child.getBoundingClientRect().bottom + (parseFloat(getComputedStyle(child).marginBottom) || 0) > innerBoundary + .5);
    if (index < 0) return null;
    const child = children[index], tail = element.cloneNode(false) as HTMLElement;
    const childTail = splitElement(child, innerBoundary - (parseFloat(getComputedStyle(child).marginBottom) || 0), depth + 1);
    if (childTail) tail.append(childTail);
    else if (index === 0 || children.slice(0, index).every(c => c.matches('colgroup'))) return null;
    else tail.append(child);
    children.slice(index + 1).forEach(c => tail.append(c));
    if (element.matches('table')) element.querySelectorAll(':scope > colgroup').forEach(c => tail.prepend(c.cloneNode(true)));
    if (element.matches('tr')) { element.style.removeProperty('height'); tail.style.removeProperty('height'); }
    return tail;
  }
  function appendTable(table: HTMLTableElement) {
    const rows = Array.from(table.rows);
    const headers: HTMLTableRowElement[] = [];
    while (rows.length && rows[0].cells.length > 0 && [...rows[0].cells].every(cell => cell.tagName === 'TH')) headers.push(rows.shift()!);
    const shell = () => {
      const next = table.cloneNode(false) as HTMLTableElement;
      table.querySelectorAll(':scope > colgroup').forEach(el => next.append(el.cloneNode(true)));
      const body = document.createElement('tbody'); next.append(body);
      headers.forEach(row => body.append(row.cloneNode(true)));
      content.append(next); return { table: next, body };
    };
    let part = shell();
    if (!rows.length) { if (!fits() && content.children.length > 1) { part.table.remove(); newPage(); part = shell(); } scaleOversized(part.table); return; }
    let bodyRows = 0;
    for (let i = 0; i < rows.length;) {
      let end = i + 1;
      for (let j = i; j < end; j++) for (const cell of rows[j].cells) end = Math.min(rows.length, Math.max(end, j + cell.rowSpan));
      const group = rows.slice(i, end).map(row => row.cloneNode(true) as HTMLTableRowElement);
      group.forEach(row => part.body.append(row));
      const split = () => {
        if (group.length !== 1 || group[0].cells.length !== 1 || group[0].cells[0].rowSpan !== 1) return;
        const tail = splitElement(group[0], content.getBoundingClientRect().bottom);
        if (tail) rows.splice(end, 0, tail as HTMLTableRowElement);
      };
      if (!fits()) split();
      if (!fits() && (bodyRows > 0 || content.children.length > 1)) {
        group.forEach(row => row.remove());
        if (!bodyRows) part.table.remove();
        newPage(); part = shell(); bodyRows = 0;
        group.forEach(row => part.body.append(row));
        if (!fits()) split();
      }
      if (!fits()) scaleOversized(part.table);
      bodyRows += group.length; i = end;
    }
  }
  function append(element: HTMLElement) {
    if (element.hasAttribute('data-section-page')) {
      settings = JSON.parse(element.getAttribute('data-section-page')!); sectionNumber++; newPage(); return;
    }
    if (element.hasAttribute('data-page-break')) { newPage(); return; }
    const table = element.matches('table') ? element : element.querySelector(':scope > table');
    if (table) { appendTable(table as HTMLTableElement); return; }
    if (element.matches('ul,ol,blockquote')) {
      // Keep list numbering when a list continues on another page.
      const children = [...element.children] as HTMLElement[];
      let shell = element.cloneNode(false) as HTMLElement; content.append(shell);
      let index = Number(element.getAttribute('start')) || 1;
      for (const child of children) {
        shell.append(child);
        if (!fits() && (shell.children.length > 1 || content.children.length > 1)) {
          child.remove(); if (!shell.children.length) shell.remove();
          newPage(); shell = element.cloneNode(false) as HTMLElement;
          if (shell.tagName === 'OL') shell.setAttribute('start', String(index));
          content.append(shell); shell.append(child);
        }
        if (!fits()) scaleOversized(shell);
        index++;
      }
      return;
    }
    content.append(element);
    if (fits()) return;
    if (element.matches('p,h1,h2,h3,h4,h5,h6,pre') && (element.textContent?.length ?? 0) > 1) {
      const text = element.textContent!;
      // Grapheme boundaries keep Korean, emoji and combining characters intact.
      const segmenter = new Intl.Segmenter(undefined, { granularity: 'grapheme' });
      const offsets = [...segmenter.segment(text)].map(s => s.index).slice(1);
      let low = 0, high = offsets.length - 1, best = 0;
      element.remove();
      while (low <= high) {
        const mid = (low + high) >> 1;
        const [a] = splitText(element, offsets[mid]); content.append(a);
        if (fits()) { best = offsets[mid]; low = mid + 1; } else high = mid - 1;
        a.remove();
      }
      if (best) {
        const [a, b] = splitText(element, best);
        content.append(a); newPage(); append(b); return;
      }
      if (content.children.length) { newPage(); append(element); return; }
      content.append(element); scaleOversized(element); return;
    }
    if (content.children.length > 1) { element.remove(); newPage(); content.append(element); }
    scaleOversized(element);
  }
  try {
    await document.fonts.ready;
    await Promise.all([...source.querySelectorAll('img')].map(img => img.decode().catch(() => {})));
    newPage();
    for (const child of [...source.children]) append(child as HTMLElement);
    return { element: host, pages: pageNumber, warnings: [...warnings], dispose: () => host.remove() };
  } catch (error) { host.remove(); throw error; }
}
