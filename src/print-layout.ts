import { pageMargins, pageFurniture, paperDimensions, type PageSettings } from './types';
import { layoutTabs } from './document-tabs';
import { layoutWordLines } from './word-line-metrics';
import { sanitizeImportedHtml } from './document-html';
import { layoutColumns } from './column-layout';
import { applyFloatingLayout } from './floating-layout';
import { imageFloating } from './image-floating';
import { layoutFurnitureImages } from './furniture-layout';
import { layoutTableAnchors, tableAnchor } from './table-anchor';
import { paragraphFitBottom } from './paragraph-fit';

export type PrintLayout = { element: HTMLElement; pages: number; pageSettings: WeakMap<Element, PageSettings>; warnings: string[]; dispose(): void };
export type PrintLayoutProgress = { label: string; completed?: number; total?: number; unit?: string };
const PX = 96 / 25.4;
const dimensions = paperDimensions;
/** Use existing visual lines: truncating at arbitrary characters can pull part
 * of the next word onto the preceding line, unlike the editable paragraph. */
function paragraphLineOffsets(element: HTMLElement) {
  const breaks = new Set<number>(), walker = document.createTreeWalker(element, NodeFilter.SHOW_TEXT);
  let offset = 0, previousTop: number | undefined;
  while (walker.nextNode()) {
    const node = walker.currentNode as Text, range = document.createRange();
    if (node.parentElement?.closest('[data-docx-tab]')) { offset += node.length; continue; }
    range.selectNodeContents(node);
    const offsets = [...new Intl.Segmenter(undefined, { granularity: 'grapheme' }).segment(node.data)].map(s => s.index);
    for (const rect of range.getClientRects()) {
      if (rect.width < .1 || rect.height < 1) continue;
      if (previousTop === undefined || Math.abs(rect.top - previousTop) > rect.height * .5) {
        let low = 0, high = offsets.length - 1, first = 0;
        while (low <= high) {
          const mid = (low + high) >> 1;
          range.setStart(node, offsets[mid]); range.setEnd(node, offsets[mid + 1] ?? node.length);
          if (range.getBoundingClientRect().top < rect.top - .5) { first = mid + 1; low = mid + 1; } else high = mid - 1;
        }
        const position = offset + (offsets[Math.min(first, offsets.length - 1)] ?? 0);
        if (position) breaks.add(position);
        previousTop = rect.top;
      }
    }
    offset += node.length;
  }
  return [...breaks].sort((a, b) => a - b);
}
function splitText(element: HTMLElement, offset: number): [HTMLElement, HTMLElement] {
  const walker = document.createTreeWalker(element, NodeFilter.SHOW_TEXT);
  let remaining = offset, node: Node | null = walker.nextNode();
  while (node && remaining > (node.textContent?.length ?? 0)) {
    remaining -= node.textContent!.length; node = walker.nextNode();
  }
  const before = document.createRange(), after = document.createRange();
  before.selectNodeContents(element); after.selectNodeContents(element);
  if (node) {
    const atomic = node.parentElement?.closest('[data-note-reference],[data-symbol-font],[data-page-number],[data-docx-tab],[data-docx-bookmark]');
    if (atomic && remaining < (node.textContent?.length || 0)) { before.setEndBefore(atomic); after.setStartBefore(atomic); }
    else { before.setEnd(node, remaining); after.setStart(node, remaining); }
  }
  else after.collapse(false);
  const a = element.cloneNode(false) as HTMLElement, b = element.cloneNode(false) as HTMLElement;
  a.append(before.cloneContents()); b.append(after.cloneContents());
  // Fragment boundaries retain the same line budget as editor continuation.
  a.dataset.printParagraphFragment = 'true'; b.dataset.printParagraphFragment = 'true';
  a.style.marginBottom = '0'; b.style.marginTop = '0'; b.style.textIndent = '0';
  return [a, b];
}

/** Build one measured page tree used unchanged by preview, PDF and native printing. */
export async function createPrintLayout(html: string, initialPage: PageSettings, onProgress?: (value: PrintLayoutProgress) => void, signal?: AbortSignal): Promise<PrintLayout> {
  signal?.throwIfAborted();
  if (onProgress) { onProgress({ label: '쪽 배치를 준비하는 중…' }); await new Promise(resolve => setTimeout(resolve, 24)); signal?.throwIfAborted(); }
  const host = document.createElement('div');
  host.className = 'print-layout measuring';
  host.setAttribute('aria-label', '인쇄 미리보기 문서');
  document.body.append(host);
  const style = document.createElement('style'); host.append(style);
  const source = document.createElement('div'); source.innerHTML = html;
  // The editable table NodeView fixes the width from colgroup; static HTML
  // must do the same or nested tables stretch to the whole printable width.
  for (const table of source.querySelectorAll('table')) {
    const columns = [...table.querySelectorAll<HTMLTableColElement>(':scope > colgroup > col')];
    if (columns.length && columns.every(c => /^\d+(?:\.\d+)?px$/.test(c.style.width)))
      table.style.width = `${columns.reduce((sum, c) => sum + parseFloat(c.style.width), 0)}px`;
  }
  const warnings = new Set<string>();
  const pageSettings = new WeakMap<Element, PageSettings>();
  const noteBodies = new Map<string, HTMLElement>();
  for (const note of source.querySelectorAll<HTMLElement>('[data-document-note][data-note-kind="footnote"]')) {
    noteBodies.set(note.dataset.documentNote!, note); note.remove();
  }
  let settings = initialPage, pageNumber = 0, content: HTMLElement, footnotes: HTMLElement;
  let deferredNotes = new Set<string>();
  let sectionNumber = 0;
  let sectionPage = 0;
  let sectionJustStarted = false;
  const pageStyles: string[] = [];
  function refreshNotes() {
    if (!content?.isConnected) return;
    const current = content.closest('.print-page')!;
    const fullBottom = content.getBoundingClientRect().top + Number(content.dataset.bodyHeight);
    const references = [...content.querySelectorAll<HTMLElement>('[data-note-reference]')];
    let ids = [...new Set(references
      .filter(el => el.getBoundingClientRect().top < fullBottom - .5).map(el => el.dataset.noteReference!))]
      .filter(id => !deferredNotes.has(id) && noteBodies.has(id) && ![...host.querySelectorAll<HTMLElement>('.print-footnotes [data-document-note]')].some(el => el.dataset.documentNote === id && el.closest('.print-page') !== current));
    if (footnotes.dataset.notes !== ids.join('|')) {
      footnotes.replaceChildren(...ids.map(id => noteBodies.get(id)!.cloneNode(true))); footnotes.dataset.notes = ids.join('|');
    }
    const noteBoundary = fullBottom - (ids.length ? footnotes.getBoundingClientRect().height + 12 : 0);
    // A later reference in a continuing row must move with its note. Keep
    // that decision while filling the other columns up to the page bottom;
    // otherwise a late note reserves blank space on the preceding page.
    for (const ref of references) if (ids.includes(ref.dataset.noteReference!) && ref.closest('tr[data-row-split="allow"]') &&
      ref.getBoundingClientRect().bottom > noteBoundary + .5) deferredNotes.add(ref.dataset.noteReference!);
    if (ids.some(id => deferredNotes.has(id))) {
      ids = ids.filter(id => !deferredNotes.has(id));
      footnotes.replaceChildren(...ids.map(id => noteBodies.get(id)!.cloneNode(true))); footnotes.dataset.notes = ids.join('|');
    }
    const reserve = ids.length ? footnotes.getBoundingClientRect().height + 12 : 0;
    content.style.height = `${Math.max(1, Number(content.dataset.bodyHeight) - reserve)}px`;
  }
  function newPage(keepBlank = false) {
    refreshNotes();
    deferredNotes = new Set();
    const [width, height] = dimensions(settings), margins = pageMargins(settings);
    const pageName = `studioSection${sectionNumber}`;
    pageStyles.push(`@page ${pageName} { size:${width}mm ${height}mm; margin:0; }`);
    style.textContent = [...new Set(pageStyles)].join('\n');
    const paper = document.createElement('section'); paper.className = 'print-page';
    pageSettings.set(paper, settings);
    paper.dataset.page = String(++pageNumber);
    if (keepBlank) paper.dataset.keepBlank = 'true';
    const first = ++sectionPage === 1;
    paper.dataset.sectionFirst = String(first);
    const furniture = pageFurniture(settings, pageNumber, first);
    Object.assign(paper.style, { width: `${width * PX}px`, height: `${height * PX}px`, page: pageName });
    content = document.createElement('div'); content.className = 'tiptap print-content';
    Object.assign(content.style, { left: `${margins.left * PX}px`, top: `${margins.top * PX}px`,
      width: `${(width - margins.left - margins.right) * PX}px`, height: `${(height - margins.top - margins.bottom) * PX}px` });
    const header = document.createElement('div'); header.className = 'print-header document-furniture';
    if (furniture.headerHtml) header.innerHTML = sanitizeImportedHtml(furniture.headerHtml).html; else header.textContent = furniture.header;
    const footer = document.createElement('div'); footer.className = 'print-footer';
    const footerText = document.createElement('div'); footerText.className = 'document-furniture';
    if (furniture.footerHtml) footerText.innerHTML = sanitizeImportedHtml(furniture.footerHtml).html; else footerText.textContent = furniture.footer;
    footer.append(footerText);
    for (const field of [...header.querySelectorAll('[data-page-number]'), ...footer.querySelectorAll('[data-page-number]')]) field.textContent = String(pageNumber);
    if (furniture.pageNumbers && !header.querySelector('[data-page-number]') && !footer.querySelector('[data-page-number]')) { const number = document.createElement('span'); number.className = 'print-automatic-number'; number.textContent = String(pageNumber); footer.append(number); }
    for (const el of [header, footer]) { el.style.left = `${margins.left * PX}px`; el.style.right = `${margins.right * PX}px`; }
    header.style.top = `${(settings.headerDistance ?? Math.max(0, margins.top / 2 - 3)) * PX}px`;
    footer.style.bottom = `${(settings.footerDistance ?? Math.max(0, margins.bottom / 2 - 3)) * PX}px`;
    paper.append(header, content, footer); host.append(paper);
    layoutFurnitureImages(paper);
    const meaningful = (element: HTMLElement) => element.textContent?.trim() || element.querySelector('img,table');
    const top = Math.max(margins.top * PX, meaningful(header) ? header.offsetTop + header.getBoundingClientRect().height + 2 : 0);
    const bottom = Math.max(margins.bottom * PX, meaningful(footer) ? parseFloat(footer.style.bottom) + footer.getBoundingClientRect().height + 2 : 0);
    content.style.top = `${top}px`; content.style.height = `${Math.max(1, height * PX - top - bottom)}px`;
    content.dataset.bodyHeight = String(Math.max(1, height * PX - top - bottom));
    footnotes = document.createElement('div'); footnotes.className = 'tiptap print-footnotes';
    Object.assign(footnotes.style, { left: `${margins.left * PX}px`, right: `${margins.right * PX}px`, bottom: `${bottom}px` });
    paper.append(footnotes);
  }
  function fits() {
    refreshNotes();
    layoutTableAnchors(content);
    layoutWordLines(content);
    layoutTabs(content);
    layoutColumns(content, content.getBoundingClientRect().bottom);
    applyFloatingLayout(content);
    const last = content.lastElementChild as HTMLElement | null;
    return !hasDeferredNote(content) && (!last || paragraphFitBottom(last) <= content.getBoundingClientRect().bottom + 0.5)
      // Behind/in-front objects do not consume text flow. A page decoration
      // extending into the margin must not push its anchor onto another page.
      && [...content.querySelectorAll<HTMLElement>('img[data-image-floating]')].every(img =>
        imageFloating(img.dataset.imageFloating)?.wrap?.type === 0 || img.getBoundingClientRect().bottom <= content.getBoundingClientRect().bottom + .5)
      // Page-relative floating pictures can extend past a column box without
      // overflowing its text. scrollWidth includes those absolute descendants
      // and would move otherwise fitting columns onto another sheet.
      && [...content.querySelectorAll<HTMLElement>('[data-column-count]')].every(group =>
        [...group.children].every(child => [...child.getClientRects()].every(rect => rect.right <= group.getBoundingClientRect().right + .5)));
  }
  function hasDeferredNote(element: HTMLElement) {
    return [...element.querySelectorAll<HTMLElement>('[data-note-reference]')].some(ref => deferredNotes.has(ref.dataset.noteReference!));
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
      let offsets = paragraphLineOffsets(element);
      if (hasDeferredNote(element)) {
        const walker = document.createTreeWalker(element, NodeFilter.SHOW_TEXT); let offset = 0;
        while (walker.nextNode()) {
          const text = walker.currentNode as Text, ref = text.parentElement?.closest<HTMLElement>('[data-note-reference]');
          if (ref && deferredNotes.has(ref.dataset.noteReference!)) break;
          offset += text.length;
        }
        const limit = Math.max(0, ...offsets.filter(n => n <= offset));
        offsets = offsets.filter(n => n <= limit);
      }
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
    if (element.matches('tr')) {
      if (element.dataset.rowSplit === 'avoid' || element.dataset.wordRowRule === 'exact') return null;
      const row = element as HTMLTableRowElement;
      if (row.cells.length > 1) {
        if (row.dataset.rowSplit !== 'allow' || [...row.cells].some(c => c.rowSpan > 1)) return null;
        const original = row.cloneNode(true) as HTMLTableRowElement, height = row.getBoundingClientRect().height;
        const available = boundary - row.getBoundingClientRect().top;
        if (available <= .5) return null;
        const tail = row.cloneNode(false) as HTMLTableRowElement;
        row.style.removeProperty('height');
        let moved = false, reservedSpace = false;
        // Split every column against the same physical boundary. Even a
        // finished column needs its empty continuation cell to keep the grid.
        for (const cell of [...row.cells]) {
          const next = cell.cloneNode(false) as HTMLTableCellElement; tail.append(next);
          const rest = splitElement(cell, boundary, depth + 1);
          if (rest) { next.append(...rest.childNodes); moved = true; }
          else if (hasDeferredNote(cell) || [...cell.children].some(child => child.getBoundingClientRect().bottom > boundary - (parseFloat(getComputedStyle(cell).paddingBottom) || 0) + .5)) {
            next.append(...cell.childNodes); moved = true;
          }
        }
        // A stored minimum height can reserve space after the text. Divide
        // that reserve too; repeating the full height on the continuation
        // would enlarge the row, while dropping it would shorten the form.
        if (!moved && height > available + .5) {
          row.style.height = `${available}px`; moved = reservedSpace = true;
        }
        const progress = reservedSpace || [...row.cells].some(c => c.textContent?.trim() || c.querySelector('img,table'));
        if (!moved || !progress || row.getBoundingClientRect().bottom > boundary + .5) {
          row.replaceChildren(...original.childNodes); row.style.cssText = original.style.cssText; return null;
        }
        tail.style.height = `${Math.max(0, height - row.getBoundingClientRect().height)}px`;
        return tail;
      }
    }
    if (element.matches('table')) {
      const table = element as HTMLTableElement;
      // A nested form table that fits a sheet moves as a unit. Splitting its
      // cell paragraphs merely to fill the preceding sheet changes the form.
      if (table.rows.length === 1 && table.getBoundingClientRect().height <= content.getBoundingClientRect().height + .5) return null;
      if ([...table.rows].some(row => [...row.cells].some(cell => cell.rowSpan > 1))) return null;
    }
    const children = [...element.children] as HTMLElement[];
    const style = getComputedStyle(element);
    const innerBoundary = boundary - (parseFloat(style.paddingBottom) || 0) - (parseFloat(style.borderBottomWidth) || 0);
    const index = children.findIndex(child => !child.matches('colgroup') && (hasDeferredNote(child) || child.getBoundingClientRect().bottom + (parseFloat(getComputedStyle(child).marginBottom) || 0) > innerBoundary + .5));
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
  // Only explicitly permitted continuing merges may cross a row boundary. Keep
  // the grid and every cell's content exactly once; this is print DOM only.
  function splitMergedGroup(group: HTMLTableRowElement[], boundary: number): HTMLTableRowElement[] | null {
    if (group.length < 2 || group.some(r => r.dataset.rowSplit === 'avoid' || r.dataset.wordRowRule === 'exact')) return null;
    const entries: { cell: HTMLTableCellElement; row: number; col: number }[] = [];
    const occupied: number[] = [];
    for (let row = 0; row < group.length; row++) {
      let col = 0;
      for (const cell of [...group[row].cells]) {
        while ((occupied[col] || 0) > row) col++;
        entries.push({ cell, row, col });
        for (let j = col; j < col + cell.colSpan; j++) occupied[j] = row + cell.rowSpan;
        col += cell.colSpan;
      }
    }
    for (let cut = group.length - 1; cut > 0; cut--) {
      if (group[cut - 1].getBoundingClientRect().bottom > boundary + .5) continue;
      const crossings = entries.filter(e => e.row < cut && e.row + e.cell.rowSpan > cut);
      if (!crossings.length || crossings.some(e => e.cell.dataset.cellSplit !== 'allow')) continue;
      const tails = group.slice(cut);
      // The continuation starts at a source row boundary, not at the bottom
      // of the sheet. Filling the unused area below that row stretches the
      // retained rows and makes editor and print disagree about the grid.
      const rowBoundary = group[cut - 1].getBoundingClientRect().bottom;
      for (const { cell, row, col } of crossings) {
        const next = cell.cloneNode(false) as HTMLTableCellElement;
        next.dataset.printMergeContinuation = 'true';
        // Partition rich content using the available row height. Applying a
        // whole-cell middle/bottom offset before partitioning can leave most
        // of the first fragment empty and overflow the last one. Each printed
        // fragment then applies the original vertical alignment independently.
        const align = cell.style.verticalAlign;
        const cellStyle = getComputedStyle(cell), child = cell.firstElementChild as HTMLElement | null;
        const childStyle = child && getComputedStyle(child);
        const rich = cell.dataset.printMergeFlow === 'rich' || cell.children.length > 1 ||
          (child?.getBoundingClientRect().height ?? 0) + (parseFloat(childStyle?.marginTop || '') || 0) + (parseFloat(childStyle?.marginBottom || '') || 0) >
          rowBoundary - cell.getBoundingClientRect().top - (parseFloat(cellStyle.paddingTop) || 0) - (parseFloat(cellStyle.paddingBottom) || 0);
        if (rich) { cell.style.verticalAlign = 'top'; next.dataset.printMergeFlow = 'rich'; }
        const remainder = splitElement(cell, rowBoundary);
        if (remainder) next.append(...remainder.childNodes);
        else if ([...cell.children].some(c => c.getBoundingClientRect().bottom > rowBoundary + .5)) next.append(...cell.childNodes);
        cell.style.verticalAlign = align;
        next.rowSpan = row + cell.rowSpan - cut; cell.rowSpan = cut - row;
        cell.style.removeProperty('height'); next.style.removeProperty('height');
        const following = entries.find(e => e.row === cut && e.col > col)?.cell;
        tails[0].insertBefore(next, following || null);
      }
      tails.forEach(r => r.remove()); group.splice(cut);
      return tails;
    }
    return null;
  }
  function appendTable(table: HTMLTableElement) {
    const rows = Array.from(table.rows);
    const headers: HTMLTableRowElement[] = [];
    while (rows.length && rows[0].cells.length > 0 && [...rows[0].cells].every(cell => cell.tagName === 'TH')) headers.push(rows.shift()!);
    const shell = (continuation = false) => {
      const next = table.cloneNode(false) as HTMLTableElement;
      if (continuation && next.dataset.tableAnchor) {
        const anchor = tableAnchor(next.dataset.tableAnchor);
        if (anchor?.reference === 'paragraph') { next.dataset.tableAnchor = JSON.stringify({ ...anchor, y: 0 }); next.style.marginTop = '0px'; }
      }
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
        const mergedTail = splitMergedGroup(group, content.getBoundingClientRect().bottom);
        if (mergedTail) { rows.splice(end, 0, ...mergedTail); return; }
        if (group.length !== 1 || [...group[0].cells].some(c => c.rowSpan !== 1)) return;
        // DOCX enables row splitting to preserve continuing vertical merges.
        // A final one-row fragment is still part of that merge: keep its row
        // boundary instead of moving only its empty height to the next sheet.
        if (group[0].querySelector(':scope > [data-print-merge-continuation]') && group[0].getBoundingClientRect().height <= content.getBoundingClientRect().height + .5) return;
        if (group[0].dataset.rowSplit !== 'allow' && !group[0].querySelector('table') && group[0].getBoundingClientRect().height <= content.getBoundingClientRect().height + .5) return;
        const tail = splitElement(group[0], content.getBoundingClientRect().bottom);
        if (tail) rows.splice(end, 0, tail as HTMLTableRowElement);
      };
      if (!fits()) split();
      if (!fits() && (bodyRows > 0 || content.children.length > 1)) {
        group.forEach(row => row.remove());
        if (!bodyRows) part.table.remove();
        newPage(); part = shell(i > 0); bodyRows = 0;
        group.forEach(row => part.body.append(row));
        if (!fits()) split();
      }
      if (!fits()) scaleOversized(part.table);
      bodyRows += group.length; i = end;
    }
  }
  function append(element: HTMLElement) {
    if (element.hasAttribute('data-section-page')) {
      const next = JSON.parse(element.getAttribute('data-section-page')!) as PageSettings;
      const sameSize = dimensions(settings).join(',') === dimensions(next).join(',');
      settings = next; sectionNumber++; sectionPage = 0;
      if (next.breakType === 'continuous' && sameSize) {
        sectionPage = 1;
        const paper = content.closest<HTMLElement>('.print-page')!, furniture = pageFurniture(settings, pageNumber, true);
        const header = paper.querySelector<HTMLElement>('.print-header')!, footer = paper.querySelector<HTMLElement>('.print-footer .document-furniture')!;
        if (furniture.headerHtml) header.innerHTML = sanitizeImportedHtml(furniture.headerHtml).html; else header.textContent = furniture.header;
        if (furniture.footerHtml) footer.innerHTML = sanitizeImportedHtml(furniture.footerHtml).html; else footer.textContent = furniture.footer;
        paper.querySelector('.print-automatic-number')?.remove();
        if (furniture.pageNumbers && !header.querySelector('[data-page-number]') && !footer.querySelector('[data-page-number]')) {
          const number = document.createElement('span'); number.className = 'print-automatic-number'; number.textContent = String(pageNumber); footer.parentElement!.append(number);
        }
        layoutFurnitureImages(paper);
        return;
      }
      const meaningful = content.textContent?.trim() || content.querySelector('img,table,[data-note-reference]');
      if (!meaningful) { content.closest('.print-page')!.remove(); pageNumber--; }
      if (next.breakType === 'oddPage' || next.breakType === 'evenPage') {
        const target = next.breakType === 'oddPage' ? 1 : 0;
        if ((pageNumber + 1) % 2 !== target) { newPage(true); sectionPage = 0; }
      }
      newPage(true); sectionJustStarted = true; return;
    }
    // A break outside a multi-column group advances to the next sheet,
    // including breaks unwrapped from a one-column section.
    if (element.hasAttribute('data-page-break') || element.hasAttribute('data-column-break')) {
      if (sectionJustStarted && !content.textContent?.trim() && !content.querySelector('img,table')) { sectionJustStarted = false; return; }
      newPage(true); return;
    }
    if (element.hasAttribute('data-section-content')) { for (const child of [...element.children]) append(child as HTMLElement); return; }
    if (element.hasAttribute('data-column-count')) {
      if ([...element.children].some(child => child.hasAttribute('data-page-break') || child.hasAttribute('data-section-page'))) {
        let part = element.cloneNode(false) as HTMLElement;
        for (const child of [...element.children]) {
          if (child.hasAttribute('data-page-break') || child.hasAttribute('data-section-page')) {
            if (part.children.length) append(part); append(child as HTMLElement); part = element.cloneNode(false) as HTMLElement;
          } else part.append(child);
        }
        if (part.children.length) append(part);
        return;
      }
      if (Number(element.dataset.columnCount) === 1) { for (const child of [...element.children]) append(child as HTMLElement); return; }
      content.append(element); if (fits()) return;
      // Split at block boundaries first. Character-only ranges lose the
      // position of empty paragraphs, pictures and forced column breaks.
      const children = [...element.children] as HTMLElement[];
      const fragment = (count: number) => {
        const part = element.cloneNode(false) as HTMLElement;
        part.replaceChildren(...children.slice(0, count).map(child => child.cloneNode(true)));
        return part;
      };
      element.remove(); let low = 1, high = children.length, best = 0;
      while (low <= high) {
        const mid = (low + high) >> 1, head = fragment(mid); content.append(head);
        if (fits()) { best = mid; low = mid + 1; } else high = mid - 1;
        head.remove();
      }
      const head = fragment(best), tail = element.cloneNode(false) as HTMLElement;
      tail.replaceChildren(...children.slice(best).map(child => child.cloneNode(true)));
      content.append(head);
      const next = tail.firstElementChild as HTMLElement | null;
      if (next?.matches('p,h1,h2,h3,h4,h5,h6') && (next.textContent?.length ?? 0) > 1) {
        const offsets = [...new Intl.Segmenter(undefined, { granularity: 'grapheme' }).segment(next.textContent!)].map(s => s.index).slice(1);
        let left = 0, right = offsets.length - 1, cut = 0;
        while (left <= right) {
          const mid = (left + right) >> 1, [part] = splitText(next, offsets[mid]); head.append(part);
          if (fits()) { cut = offsets[mid]; left = mid + 1; } else right = mid - 1;
          part.remove();
        }
        if (cut) { const [part, rest] = splitText(next, cut); head.append(part); next.replaceWith(rest); }
      }
      if (head.children.length) {
        fits(); newPage();
        // The break that overflowed the last column is fulfilled by this new
        // sheet. Applying it again at its start would skip the first column.
        if (tail.firstElementChild?.hasAttribute('data-column-break')) tail.firstElementChild.remove();
        if (tail.children.length) append(tail);
        return;
      }
      head.remove();
      if (content.children.length) { newPage(); append(element); return; }
      content.append(element); fits(); scaleOversized(element); return;
    }
    if (element.textContent?.trim() || element.querySelector('img,table')) sectionJustStarted = false;
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
      const offsets = paragraphLineOffsets(element);
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
    onProgress?.({ label: '문서의 글꼴을 준비하는 중…' });
    const sample = source.textContent || '한글';
    for (const family of ['Noto Serif KR', 'Noto Sans KR', 'Studio Text', 'Studio Math', 'Studio Symbols', 'Studio Symbols 2', 'Studio Devanagari', 'Studio Emoji']) if (html.includes(family))
      await Promise.all([400, 700].map(weight => document.fonts.load(`${weight} 12px "${family}"`, sample)));
    if (source.querySelector('[data-symbol-font]')) await Promise.all(['Studio Symbols 2','Studio Symbols','Studio Math','Studio Text','Studio Devanagari','Studio Emoji'].map(font => document.fonts.load(`12px "${font}"`)));
    await document.fonts.ready;
    signal?.throwIfAborted();
    const images = [...source.querySelectorAll('img')]; let decoded = 0;
    if (images.length) onProgress?.({ label: '쪽 배치에 사용할 그림을 준비하는 중…', completed: 0, total: images.length, unit: '개' });
    await Promise.all(images.map(async img => {
      await img.decode().catch(() => {});
      if (!signal?.aborted) onProgress?.({ label: '쪽 배치에 사용할 그림을 준비하는 중…', completed: ++decoded, total: images.length, unit: '개' });
    }));
    signal?.throwIfAborted();
    newPage();
    const children = [...source.children]; let processed = 0, yieldedAt = performance.now();
    for (const child of children) {
      if (onProgress && (child.tagName === 'TABLE' || performance.now() - yieldedAt >= 32)) {
        onProgress(child.tagName === 'TABLE' ? { label: '표의 쪽 배치를 계산하는 중…' } : { label: 'Word 쪽 배치를 계산하는 중…', completed: processed, total: children.length, unit: '항목' });
        await new Promise(resolve => setTimeout(resolve, 0)); signal?.throwIfAborted(); yieldedAt = performance.now();
      }
      append(child as HTMLElement); processed++;
      onProgress?.({ label: 'Word 쪽 배치를 계산하는 중…', completed: processed, total: children.length, unit: '항목' });
    }
    refreshNotes();
    // Empty trailing paragraphs can spill onto an automatically created sheet.
    // Keep intentional page/section breaks and parity sheets, but do not print
    // an extra blank sheet produced only by those trailing line boxes.
    let last = host.querySelector<HTMLElement>('.print-page:last-child');
    while (pageNumber > 1 && last && !last.dataset.keepBlank && !last.querySelector('.print-content')?.textContent?.trim()
      && !last.querySelector('.print-content img,.print-content table,.print-content hr,.print-footnotes [data-document-note]')) {
      last.remove(); pageNumber--; last = host.querySelector<HTMLElement>('.print-page:last-child');
    }
    for (const field of host.querySelectorAll<HTMLElement>('[data-page-number="NUMPAGES"]')) field.textContent = String(pageNumber);
    for (const field of host.querySelectorAll<HTMLElement>('[data-page-number="PAGE"]')) field.textContent = field.closest<HTMLElement>('.print-page')!.dataset.page!;
    return { element: host, pages: pageNumber, pageSettings, warnings: [...warnings], dispose: () => host.remove() };
  } catch (error) { host.remove(); throw error; }
}
