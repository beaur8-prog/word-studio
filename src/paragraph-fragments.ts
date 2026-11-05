import {imageFloating} from './image-floating';
// A hard break belongs to the preceding printed line. Give it an offset so
// cutting after that line does not leave a spurious blank line on the tail.
export function paragraphTextWalker(element: HTMLElement) {
  return document.createTreeWalker(element, NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT, {
    acceptNode: node => node instanceof Text || node instanceof HTMLBRElement ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_SKIP,
  });
}
export const textUnitLength = (node: Node) => node instanceof Text ? node.length : 1;
/** HTML ignores the empty line after a final BR unless it has an inline box.
 * This view-only pseudo-content reserves that genuine document line. */
export function prepareParagraphEnds(root: HTMLElement) {
  const paragraphs = [...(root.matches('p,h1,h2,h3,h4,h5,h6') ? [root] : []), ...root.querySelectorAll<HTMLElement>('p,h1,h2,h3,h4,h5,h6')];
  for (const p of paragraphs) {
    // An anchor-only Word paragraph still has a paragraph-mark line. The
    // editor supplies a trailing BR; its static HTML needs the same strut.
    const images=[...p.querySelectorAll<HTMLImageElement>('img')];
    if(images.some(img=>imageFloating(img.dataset.imageFloating)?.drawingGroup) && !p.textContent?.trim() && !p.querySelector('br') && images.every(img=>imageFloating(img.dataset.imageFloating)))p.dataset.printDrawingAnchor='true';
    else delete p.dataset.printDrawingAnchor;
    const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT); let last: Node | undefined;
    while (walker.nextNode()) {
      const node = walker.currentNode;
      if (node instanceof HTMLBRElement && !node.classList.contains('ProseMirror-trailingBreak') || node instanceof HTMLImageElement || node instanceof Text && node.data.trim()) last = node;
    }
    const continuedEmptyEnd = !last && p.dataset.printParagraphFragment && p.dataset.printTrailingBreak;
    if (!p.dataset.printSoftEnd && (last instanceof HTMLBRElement || continuedEmptyEnd)) {
      p.dataset.printTrailingBreak = 'true';
      // Plain <p><br></p> is also used as an empty HTML placeholder. Keep
      // authored multi-line flow and empty tails derived from text, while
      // allowing the existing automatic trailing-placeholder trim.
      if (p.textContent?.trim() || p.querySelectorAll('br').length > 1 || p.dataset.printKeepTrailing)
        p.dataset.printKeepTrailing = 'true';
    } else { delete p.dataset.printTrailingBreak; delete p.dataset.printKeepTrailing; }
  }
}
/** Use existing visual lines: truncating at arbitrary characters can pull part
 * of the next word onto the preceding line, unlike the editable paragraph. */
export function paragraphLineOffsets(element: HTMLElement) {
  const breaks = new Set<number>(), walker = paragraphTextWalker(element);
  let offset = 0, previousTop: number | undefined;
  while (walker.nextNode()) {
    if (walker.currentNode instanceof HTMLBRElement) { breaks.add(++offset); previousTop = undefined; continue; }
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
  return [...breaks].filter(n => n < offset || n === offset && !!element.dataset.printTrailingBreak).sort((a, b) => a - b);
}
export function splitText(element: HTMLElement, offset: number): [HTMLElement, HTMLElement] {
  const walker = paragraphTextWalker(element);
  let remaining = offset, node: Node | null = walker.nextNode();
  while (node && remaining > textUnitLength(node)) {
    remaining -= textUnitLength(node); node = walker.nextNode();
  }
  const before = document.createRange(), after = document.createRange();
  before.selectNodeContents(element); after.selectNodeContents(element);
  if (node) {
    const atomic = node.parentElement?.closest('[data-note-reference],[data-symbol-font],[data-page-number],[data-docx-tab],[data-docx-bookmark]');
    if (node instanceof HTMLBRElement) {
      if (remaining) { before.setEndAfter(node); after.setStartAfter(node); }
      else { before.setEndBefore(node); after.setStartBefore(node); }
    }
    else if (atomic && remaining < (node.textContent?.length || 0)) { before.setEndBefore(atomic); after.setStartBefore(atomic); }
    else { before.setEnd(node, remaining); after.setStart(node, remaining); }
  }
  else after.collapse(false);
  const a = element.cloneNode(false) as HTMLElement, b = element.cloneNode(false) as HTMLElement;
  a.append(before.cloneContents()); b.append(after.cloneContents());
  a.dataset.printSoftEnd = 'true'; delete a.dataset.printTrailingBreak;
  if (element.dataset.fragmentOffset !== undefined) b.dataset.fragmentOffset = String(Number(element.dataset.fragmentOffset) + offset);
  // Fragment boundaries retain the same line budget as editor continuation.
  a.dataset.printParagraphFragment = 'true'; b.dataset.printParagraphFragment = 'true';
  if (element.dataset.hwpCellLine) {
    a.dataset.hwpCellContinuation = 'true'; delete b.dataset.hwpCellContinuation;
    // A paragraph ending in a real blank line uses a full strut throughout.
    // Splitting must not switch its earlier fragments to compact ink insets.
    if (element.dataset.hwpCellFullLine || !element.dataset.hwpCellTerminal) a.dataset.hwpCellFullLine = b.dataset.hwpCellFullLine = 'true';
  }
  a.style.marginBottom = '0'; b.style.marginTop = '0'; b.style.textIndent = '0';
  return [a, b];
}


/** Text/BR offset of a DOM boundary, used by both detached print fragments and
 * live editor line positions. Inline view widgets do not contribute text. */
export function paragraphTextOffset(paragraph: HTMLElement, node: Node, offset: number) {
  const range = document.createRange(); range.selectNodeContents(paragraph); range.setEnd(node, offset);
  const fragment = document.createElement('div'); fragment.append(range.cloneContents());
  fragment.querySelectorAll('.pagination-gap,.pagination-cell-gap,[data-flow-widget]').forEach(widget => widget.remove());
  const walker = paragraphTextWalker(fragment); let length = 0;
  while (walker.nextNode()) length += textUnitLength(walker.currentNode);
  return length;
}

/** A superscript on a later line can distort half-leading inferred from the
 * gap between glyph boxes. Measure the actual retained prefix near a page
 * edge, just as print does, without touching ProseMirror's live DOM. */
export function paragraphPrefixHeight(paragraph: HTMLElement, offset: number, scale: number) {
  const [head] = splitText(paragraph, offset), host = document.createElement('div');
  host.className = 'tiptap'; host.setAttribute('aria-hidden', 'true');
  const css = getComputedStyle(paragraph.parentElement!);
  host.style.cssText = `position:fixed;visibility:hidden;pointer-events:none;left:-100000px;top:0;padding:0;min-height:0;width:${paragraph.getBoundingClientRect().width / scale}px`;
  for (const key of ['fontFamily', 'fontSize', 'fontWeight', 'fontStyle', 'lineHeight', 'letterSpacing'] as const) host.style[key] = css[key];
  head.querySelectorAll('.pagination-gap,.pagination-cell-gap,[data-flow-widget],.ProseMirror-separator').forEach(widget => widget.remove());
  Object.assign(head.style, { margin: '0', width: '100%', boxSizing: 'border-box' }); host.append(head); document.body.append(host);
  try { return head.getBoundingClientRect().height; } finally { host.remove(); }
}

/** Word keep-lines protects any fitting ordinary paragraph. A short body
 * paragraph also cannot satisfy both sides of a widow-controlled break
 * when its line count is below orphans + widows. Move these intact. Oversized paragraphs and other flow contexts retain
 * their own continuation rules, so protection never forces text to shrink.
 * Tab stops also keep their existing flow until their line budget is shared. */
export function keepShortWordParagraphTogether(paragraph: HTMLElement, lineCount: number, bodyHeight: number, scale = 1) {
  if (!paragraph.matches('p[data-word-paragraph-style],h1[data-word-paragraph-style],h2[data-word-paragraph-style],h3[data-word-paragraph-style],h4[data-word-paragraph-style],h5[data-word-paragraph-style],h6[data-word-paragraph-style]') || lineCount < 2 ||
      paragraph.closest('li,td,th,[data-document-textbox],[data-textbox-content],[data-column-count]:not([data-column-count="1"])') ||
      paragraph.textContent?.includes('\t') ||
      paragraph.querySelector('img:not(.ProseMirror-separator),svg,table,[data-note-reference],[data-page-number],[data-flow-widget],[data-docx-tab]')) return false;
  const orphans = Math.max(1, Number(paragraph.style.orphans) || 1), widows = Math.max(1, Number(paragraph.style.widows) || 1);
  const before = Math.max(0, parseFloat(getComputedStyle(paragraph).marginTop) || 0);
  return (paragraph.dataset.wordKeepLines === 'true' || paragraph.tagName === 'P' && lineCount < orphans + widows) && paragraph.getBoundingClientRect().height / scale + before <= bodyHeight + .5;
}
