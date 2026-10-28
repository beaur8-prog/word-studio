import type { EditorView } from '@tiptap/pm/view';

type Line = { top: number; bottom: number; pos: () => number; before?: number; after?: number };

export function measureParagraphLines(view: EditorView, block: HTMLElement, naturalY: (value: number) => number) {
  const firstPos = view.posAtDOM(block, 0) - 1;
  const blockLines: Line[] = [];
  const add = (top: number, bottom: number, pos: () => number) => {
    const previous = blockLines.at(-1);
    if (previous && Math.min(previous.bottom, bottom) - Math.max(previous.top, top) > Math.min(previous.bottom - previous.top, bottom - top) * .4) {
      previous.top = Math.min(previous.top, top); previous.bottom = Math.max(previous.bottom, bottom); return;
    }
    blockLines.push({ top, bottom, pos });
  };
  const walker = document.createTreeWalker(block, NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT);
  while (walker.nextNode()) {
    const node = walker.currentNode;
    if (node instanceof HTMLElement) {
      if (node.tagName !== 'IMG') continue;
      const box = node.getBoundingClientRect();
      // ProseMirror puts an empty IMG before a non-editable inline atom. It
      // is a caret separator, not a line; counting it corrupts the first
      // line's inset when a paragraph starts with a footnote reference.
      if (box.width < .1 || box.height < 1) continue;
      add(naturalY(box.top), naturalY(box.bottom), () => view.posAtDOM(node.parentNode!, [...node.parentNode!.childNodes].indexOf(node)));
    } else if (node instanceof Text && node.data && !node.parentElement?.closest('.pagination-gap,[data-flow-widget]')) {
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
  if (!blockLines.length) { const box = block.getBoundingClientRect(); blockLines.push({ top: naturalY(box.top), bottom: naturalY(box.bottom), pos: () => firstPos }); }
  else blockLines[0].pos = () => firstPos;
  const box = block.getBoundingClientRect();
  for (let i = 0; i < blockLines.length; i++) {
    const line = blockLines[i], previous = blockLines[i - 1], next = blockLines[i + 1];
    // Superscripts and large runs can make only one line taller. Derive the
    // leading from adjacent measured lines instead of applying the first
    // line's inset to every line in the paragraph.
    line.before = Math.max(0, previous ? (line.top - previous.bottom) / 2 : line.top - naturalY(box.top));
    line.after = Math.max(0, next ? (next.top - line.bottom) / 2 : naturalY(box.bottom) - line.bottom);
  }
  return blockLines;
}
