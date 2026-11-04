import { measureParagraphLines } from './editor-lines';
import { paragraphTextWalker, splitText, textUnitLength } from './paragraph-fragments';

/** A fitting item moves with its marker. Only items taller than a full body
 * need internal continuation; editor and print use this same threshold. */
export const wholeListItem = (height: number, bodyHeight: number) => height <= bodyHeight + .5;

/** Resolve the print cut from the editor's measured visual lines. Positions in
 * detached HTML count text/BR units instead of ProseMirror node boundaries. */
function listLineOffsets(paragraph: HTMLElement) {
  const positions = { posAtDOM(node: Node, offset: number) {
    const range = document.createRange(); range.selectNodeContents(paragraph); range.setEnd(node, offset);
    const fragment = document.createElement('div'); fragment.append(range.cloneContents());
    const walker = paragraphTextWalker(fragment); let length = 1;
    while (walker.nextNode()) length += textUnitLength(walker.currentNode);
    return length;
  } };
  return measureParagraphLines(positions, paragraph, value => value).slice(1).map(line => line.pos() - 1);
}

/** Partition a too-tall list item without changing font size, indentation or
 * document content. Only the returned continuation suppresses its own marker;
 * nested lists keep their counters and following items keep advancing. */
export function splitListItem(item: HTMLElement, boundary: number, bodyHeight: number): HTMLElement | null {
  const split = (element: HTMLElement, depth: number): HTMLElement | null => {
    if (depth > 30) return null;
    if (element.matches('p,h1,h2,h3,h4,h5,h6')) {
      if (element.querySelector('img:not(.ProseMirror-separator),table,[data-page-number]')) return null;
      const offsets = listLineOffsets(element), marker = document.createComment('list-line');
      element.replaceWith(marker); let low = 0, high = offsets.length - 1, best = 0;
      while (low <= high) {
        const mid = (low + high) >> 1, [head] = splitText(element, offsets[mid]); marker.before(head);
        if (head.getBoundingClientRect().bottom <= boundary + .5) { best = offsets[mid]; low = mid + 1; }
        else high = mid - 1;
        head.remove();
      }
      marker.replaceWith(element);
      if (!best) return null;
      const [head, tail] = splitText(element, best); element.replaceWith(head); return tail;
    }
    if (!element.matches('li,ol,ul')) return null;
    const children = [...element.children] as HTMLElement[];
    const index = children.findIndex(child => child.getBoundingClientRect().bottom > boundary + .5);
    if (index < 0) return null;
    const child = children[index];
    const rest = child.matches('li') && wholeListItem(child.getBoundingClientRect().height, bodyHeight)
      ? null : split(child, depth + 1);
    // No progress in the first child: let the enclosing item move as a whole.
    if (!rest && index === 0) return null;
    const tail = element.cloneNode(false) as HTMLElement;
    if (rest) tail.append(rest); else tail.append(child);
    children.slice(index + 1).forEach(next => tail.append(next));
    if (element.matches('ol')) tail.setAttribute('start', String((Number(element.getAttribute('start')) || 1) + index));
    if (element.matches('li')) tail.dataset.printListContinuation = 'true';
    return tail;
  };
  return split(item, 0);
}
