import { measureParagraphLines } from './editor-lines';
import { paragraphTextOffset, splitText } from './paragraph-fragments';

/** A fitting item moves with its marker. Only items taller than a full body
 * need internal continuation; editor and print use this same threshold. */
export const wholeListItem = (height: number, bodyHeight: number) => height <= bodyHeight + .5;

/** Resolve the print cut from the editor's measured visual lines. Positions in
 * detached HTML count text/BR units instead of ProseMirror node boundaries. */
function listLineOffsets(paragraph: HTMLElement) {
  const positions = { posAtDOM: (node: Node, offset: number) => 1 + paragraphTextOffset(paragraph, node, offset) };
  return measureParagraphLines(positions, paragraph, value => value).slice(1).map(line => line.pos() - 1);
}

/** Partition a too-tall list item without changing font size, indentation or
 * document content. Only the returned continuation suppresses its own marker;
 * nested lists keep their counters and following items keep advancing. */
export function splitListItem(item: HTMLElement, context: {
  fits(): boolean; bodyHeight(element: HTMLElement): number;
}): HTMLElement | null {
  const split = (element: HTMLElement, depth: number): HTMLElement | null => {
    if (depth > 30) return null;
    if (element.matches('p,h1,h2,h3,h4,h5,h6')) {
      if (element.querySelector('img:not(.ProseMirror-separator),table,[data-page-number]')) return null;
      const offsets = listLineOffsets(element), marker = document.createComment('list-line');
      element.replaceWith(marker); let low = 0, high = offsets.length - 1, best = 0;
      while (low <= high) {
        const mid = (low + high) >> 1, [head] = splitText(element, offsets[mid]); marker.before(head);
        // A reference on the tail reserves space on its own page. Recompute
        // the body with just this prefix, rather than keeping the full item's
        // footnotes while searching for an earlier line boundary.
        if (context.fits()) { best = offsets[mid]; low = mid + 1; }
        else high = mid - 1;
        head.remove();
      }
      marker.replaceWith(element);
      if (!best) return null;
      const [head, tail] = splitText(element, best); element.replaceWith(head); return tail;
    }
    if (!element.matches('li,ol,ul')) return null;
    const children = [...element.children] as HTMLElement[];
    let low = 0, high = children.length, best = 0;
    while (low <= high) {
      const mid = (low + high) >> 1; element.replaceChildren(...children.slice(0, mid));
      if (context.fits()) { best = mid; low = mid + 1; } else high = mid - 1;
    }
    element.replaceChildren(...children);
    if (best === children.length) return null;
    const child = children[best];
    // Keep later blocks out of the measurement too: their notes must not
    // shorten an earlier prefix that will stay on this sheet.
    element.replaceChildren(...children.slice(0, best + 1));
    const rest = child.matches('li') && wholeListItem(child.getBoundingClientRect().height, context.bodyHeight(child))
      ? null : split(child, depth + 1);
    if (!rest && best === 0) { element.replaceChildren(...children); return null; }
    const tail = element.cloneNode(false) as HTMLElement;
    if (rest) tail.append(rest); else tail.append(child);
    children.slice(best + 1).forEach(next => tail.append(next));
    if (element.matches('ol')) tail.setAttribute('start', String((Number(element.getAttribute('start')) || 1) + best));
    if (element.matches('li')) tail.dataset.printListContinuation = 'true';
    return tail;
  };
  return split(item, 0);
}
