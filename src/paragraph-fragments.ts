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

