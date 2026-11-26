import { layoutHwpCellLines } from './hwp-cell-line-metrics';
import { prepareParagraphEnds } from './paragraph-fragments';
import { precedingSourceRow, tableFragmenter } from './table-fragments';
import { layoutWordLines } from './word-line-metrics';

function copyPrintRow(row: HTMLTableRowElement) {
  const copy = row.cloneNode(true) as HTMLTableRowElement;
  for (const widget of copy.querySelectorAll('.ProseMirror-widget,.ProseMirror-trailingBreak,.ProseMirror-separator')) widget.remove();
  for (const block of [copy, ...copy.querySelectorAll<HTMLElement>('*')]) {
    block.classList.remove('pagination-split-row', 'pagination-merge-row', 'pagination-merge-padding', 'pagination-merge-flow');
    if (block.classList.contains('tableWrapper') && block.firstElementChild instanceof HTMLTableElement) {
      const nested = block.firstElementChild;
      nested.style.marginTop = block.style.marginTop;
      nested.style.marginBottom = block.style.marginBottom;
      block.replaceWith(nested);
    }
  }
  return copy;
}

/** Measure the space after a single continued row on a detached print copy.
 * The editable cut in every cell must already match that print fragment. */
export function measuredNestedHeadSpace(row: HTMLTableRowElement, portions: { head: string; tail: string }[],
  available: number, bodyHeight: number, scale: number) {
  const original = row.closest('table'), previous = precedingSourceRow(row);
  if (!original || !previous || !(available > .5) || !portions.some(part => part.tail)) return;
  const width = original.getBoundingClientRect().width / scale;
  const host = document.createElement('section'); host.className = 'print-page'; host.setAttribute('aria-hidden', 'true');
  host.style.cssText = 'position:fixed;visibility:hidden;pointer-events:none;left:-100000px;top:0;padding:0';
  const content = document.createElement('div'); content.className = 'print-content tiptap';
  content.style.cssText = `width:${width}px;padding:0`;
  const table = original.cloneNode(false) as HTMLTableElement;
  table.style.width = `${width}px`; table.style.margin = '0';
  const columns = original.querySelector(':scope > colgroup');
  if (columns) table.append(columns.cloneNode(true));
  const body = document.createElement('tbody'), head = copyPrintRow(row);
  // The previous row supplies the authored collapsed border at the top.
  body.append(copyPrintRow(previous), head); table.append(body); content.append(table); host.append(content);
  document.body.append(host);
  try {
    prepareParagraphEnds(head); layoutWordLines(head); layoutHwpCellLines(head);
    const start = head.getBoundingClientRect(), rowSize = row.getBoundingClientRect();
    if (Math.abs(start.height - rowSize.height / scale) > .5 || Math.abs(start.width - rowSize.width / scale) > .5 ||
      [...head.cells].some((cell, i) => Math.abs(cell.getBoundingClientRect().width - row.cells[i].getBoundingClientRect().width / scale) > .5)) return;
    const boundary = start.top + available;
    const fragments = tableFragmenter({ bodyHeight: () => bodyHeight, hasDeferredNote: () => false, deferredNotes: () => new Set() });
    const tail = fragments.splitElement(head, boundary) as HTMLTableRowElement | null;
    if (!tail || head.cells.length !== portions.length || tail.cells.length !== portions.length ||
      portions.some((part, i) => head.cells[i].textContent !== part.head || tail.cells[i].textContent !== part.tail)) return;
    const end = head.getBoundingClientRect().bottom;
    if (end > boundary + .5) return;
    const range = document.createRange(); range.selectNodeContents(head);
    if ([...range.getClientRects()].some(box => box.width > 0 && box.height > 0 && box.bottom > boundary + .5)) return;
    return Math.max(0, boundary - end);
  } finally { host.remove(); }
}
