import type { EditorView } from '@tiptap/pm/view';

type Line = { top: number; bottom: number; pos: () => number; before?: number; after?: number; blank?: boolean };

export function measureParagraphLines(view: EditorView, block: HTMLElement, naturalY: (value: number) => number) {
  const firstPos = view.posAtDOM(block, 0) - 1;
  const blockLines: Line[] = [];
  const realBreaks = [...block.querySelectorAll<HTMLBRElement>('br:not(.ProseMirror-trailingBreak)')];
  if (!block.textContent?.trim() && !block.querySelector('img:not(.ProseMirror-separator),svg,[data-note-reference],[data-page-number],[data-word-overlap]')) {
    const count = Math.max(1, realBreaks.length + (block.dataset.printSoftEnd ? 0 : 1));
    const box = block.getBoundingClientRect(), top = naturalY(box.top), pitch = (naturalY(box.bottom) - top) / count;
    return Array.from({ length: count }, (_, i) => ({ top: top + i * pitch, bottom: top + (i + 1) * pitch, before: 0, after: 0, blank: true,
      pos: () => i === 0 ? firstPos : i < realBreaks.length
        ? view.posAtDOM(realBreaks[i].parentNode!, [...realBreaks[i].parentNode!.childNodes].indexOf(realBreaks[i])) : view.posAtDOM(block, block.childNodes.length) }));
  }
  const add = (top: number, bottom: number, pos: () => number, blank = false) => {
    const previous = blockLines.at(-1);
    if (previous && Math.min(previous.bottom, bottom) - Math.max(previous.top, top) > Math.min(previous.bottom - previous.top, bottom - top) * .4) {
      previous.top = Math.min(previous.top, top); previous.bottom = Math.max(previous.bottom, bottom); previous.blank = !!previous.blank && blank; return;
    }
    blockLines.push({ top, bottom, pos, blank });
  };
  const walker = document.createTreeWalker(block, NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT);
  while (walker.nextNode()) {
    const node = walker.currentNode;
    if (node instanceof HTMLElement) {
      if (node instanceof HTMLBRElement) {
        if (node.classList.contains('ProseMirror-trailingBreak') && !realBreaks.length) continue;
        const box = node.getBoundingClientRect();
        if (box.height > 1) add(naturalY(box.top), naturalY(box.bottom), () => view.posAtDOM(node.parentNode!, [...node.parentNode!.childNodes].indexOf(node)), true);
        continue;
      }
      if (node.tagName !== 'IMG') continue;
      const box = node.getBoundingClientRect();
      // ProseMirror puts an empty IMG before a non-editable inline atom. It
      // is a caret separator, not a line; counting it corrupts the first
      // line's inset when a paragraph starts with a footnote reference.
      if (box.width < .1 || box.height < 1) continue;
      add(naturalY(box.top), naturalY(box.bottom), () => view.posAtDOM(node.parentNode!, [...node.parentNode!.childNodes].indexOf(node)));
    } else if (node instanceof Text && node.data && !node.parentElement?.closest('.pagination-gap,[data-flow-widget],[data-docx-tab]')) {
      const range = document.createRange(); range.selectNodeContents(node);
      for (const box of range.getClientRects()) {
        if (box.width < .1 || box.height < 1) continue;
        const top = naturalY(box.top), bottom = naturalY(box.bottom), target = box.top;
        add(top, bottom, () => {
          // Find the first character on this visual line only if a page
          // gap is needed. Keep surrogate pairs and combining clusters.
          const offsets = [...new Intl.Segmenter(undefined, { granularity: 'grapheme' }).segment(node.data)].map(s => s.index);
          let low = 0, high = offsets.length - 1, first = 0;
          while (low <= high) {
            const mid = (low + high) >> 1, end = offsets[mid + 1] ?? node.length;
            range.setStart(node, offsets[mid]); range.setEnd(node, end);
            if (range.getBoundingClientRect().top < target - .5) { first = mid + 1; low = mid + 1; } else high = mid - 1;
          }
          return view.posAtDOM(node, offsets[Math.min(first, offsets.length - 1)]);
        });
      }
    }
  }
  if (block.dataset.printTrailingBreak) {
    // Probe only the detached/print view. Live ProseMirror already supplies
    // its trailing caret BR; mutating that editor DOM would affect input.
    const probe = document.createElement('span'); probe.textContent = '\u200b'; block.append(probe);
    try { const range = document.createRange(); range.selectNodeContents(probe); const box = range.getBoundingClientRect();
      add(naturalY(box.top), naturalY(box.bottom), () => view.posAtDOM(block, block.childNodes.length), true);
    } finally { probe.remove(); }
  }
  if (!blockLines.length) { const box = block.getBoundingClientRect(); blockLines.push({ top: naturalY(box.top), bottom: naturalY(box.bottom), pos: () => firstPos }); }
  else blockLines[0].pos = () => firstPos;
  const box = block.getBoundingClientRect();
  const blankSlots = (top: number, pitch: number, end = pitch) => blockLines.forEach((line, i) => {
    if (!line.blank) return;
    line.top = top + i * pitch; line.bottom = line.top + (i === blockLines.length - 1 ? end : pitch); line.before = line.after = 0;
  });
  for (let i = 0; i < blockLines.length; i++) {
    const line = blockLines[i], previous = blockLines[i - 1], next = blockLines[i + 1];
    // Superscripts and large runs can make only one line taller. Derive the
    // leading from adjacent measured lines instead of applying the first
    // line's inset to every line in the paragraph.
    line.before = Math.max(0, previous ? (line.top - previous.bottom) / 2 : line.top - naturalY(box.top));
    line.after = Math.max(0, next ? (next.top - line.bottom) / 2 : naturalY(box.bottom) - line.bottom);
  }
  if (block.dataset.hwpCellTerminal && blockLines.length > 1) {
    const advance = Number(block.dataset.hwpCellAdvance), end = Number(block.dataset.hwpCellEnd), top = naturalY(box.top);
    const first = blockLines[0], pitch = (blockLines.at(-1)!.top - first.top) / (blockLines.length - 1);
    const expected = naturalY(box.bottom) - top, terminal = expected - (blockLines.length - 1) * pitch;
    // CSS rounds each line to a layout unit. Accumulating the unrounded
    // imported pitch can otherwise reject a fitting line on a long cell.
    if (advance > 0 && end > 0 && pitch > 0 && Math.abs(pitch - advance) <= .5 && Math.abs(terminal - end) <= .6 &&
      blockLines.every((line, i) => Math.abs(line.top - first.top - i * pitch) <= .6)) {
      for (let i = 0; i < blockLines.length; i++) {
        blockLines[i].before = blockLines[i].top - top - i * pitch;
        blockLines[i].after = top + (i === blockLines.length - 1 ? expected : (i + 1) * pitch) - blockLines[i].bottom;
      }
      blankSlots(top, pitch, terminal);
      return blockLines;
    }
  }
  // A font's glyph box need not sit symmetrically inside an exact line strut.
  // Halving the gap between glyph boxes can undercount the actual line by a
  // fraction of a pixel and keep a line which print moves to the next sheet.
  // Use the real paragraph grid only when uniform measured lines prove it.
  // Mixed-height lines, overlapping glyphs, objects and AUTO/minimum leading
  // keep the individual geometry above.
  if (blockLines.length > 1 && /(?:px|pt)$/.test(block.style.lineHeight) && !['auto', 'atLeast'].includes(block.dataset.wordLineRule || '') &&
    !block.querySelector('img:not(.ProseMirror-separator),svg,table,[data-note-reference],[data-page-number]')) {
    const css = getComputedStyle(block), top = naturalY(box.top), pitch = (naturalY(box.bottom) - top) / blockLines.length;
    const first = blockLines[0], firstHeight = first.bottom - first.top;
    if (pitch > 0 && Math.abs(pitch - parseFloat(css.lineHeight)) <= .5 && firstHeight <= pitch + .5 &&
      ['paddingTop', 'paddingBottom', 'borderTopWidth', 'borderBottomWidth'].every(key => !(parseFloat(css[key as 'paddingTop']) || 0)) &&
      blockLines.every((line, i) => Math.abs(line.top - first.top - i * pitch) <= .5 && Math.abs(line.bottom - line.top - firstHeight) <= .5)) {
      for (let i = 0; i < blockLines.length; i++) {
        blockLines[i].before = Math.max(0, blockLines[i].top - top - i * pitch);
        blockLines[i].after = Math.max(0, top + (i + 1) * pitch - blockLines[i].bottom);
      }
      blankSlots(top, pitch);
    }
  }
  return blockLines;
}
