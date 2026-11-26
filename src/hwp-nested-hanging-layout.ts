import { characterScaleProperty, parseCharacterScale } from './character-scale';
import { lineOpportunities, sourceLineBoundaries } from './hwp-source-cell-lines';
import { hwpSourceWrapAttribute, hwpWrapFingerprint, normalizeHwpWrapText, parseHwpSourceWrap } from './hwp-text-advances';

function matchesAuthoredScale(paragraph: HTMLParagraphElement, expected: number, size: number) {
  const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT);
  while (walker.nextNode()) {
    const text = walker.currentNode as Text;
    if (!text.length) continue;
    const style = getComputedStyle(text.parentElement!);
    if (Math.abs(parseFloat(style.fontSize) - size) > .06) return false;
    let mark = text.parentElement, scale: number | null = null;
    while (mark && mark !== paragraph) {
      scale = parseCharacterScale(mark.style.getPropertyValue(characterScaleProperty));
      if (scale !== null) break;
      mark = mark.parentElement;
    }
    if (scale !== expected) return false;
  }
  return true;
}

/** A nested table may have several independently formatted paragraphs in a
 * cell. Its measured print copy can reproduce a verified hanging paragraph
 * without changing the editing model or applying a cell-wide transform. */
export function layoutNestedHwpHangingScale(ownerTable: HTMLTableElement) {
  const range = document.createRange();
  for (const paragraph of ownerTable.querySelectorAll<HTMLParagraphElement>(`p[${hwpSourceWrapAttribute}]`)) {
    const source = parseHwpSourceWrap(paragraph.getAttribute(hwpSourceWrapAttribute));
    const fit = source?.nestedHanging, width = source?.cellWidth;
    if (!fit || !width) continue;
    const cell = paragraph.parentElement, table = cell?.closest('table'), owner = table?.parentElement?.closest('td,th');
    if (!(cell instanceof HTMLTableCellElement) || !(owner instanceof HTMLTableCellElement) ||
      cell.tagName !== 'TD' || owner.tagName !== 'TD' || cell.colSpan !== 1 || cell.rowSpan !== 1 ||
      owner.colSpan !== 1 || owner.rowSpan !== 1 || owner.closest('table') !== ownerTable ||
      table?.dataset.wordTableWidthAuto === 'true' || table?.dataset.cellTablePosition ||
      !table?.querySelector(':scope > colgroup > col') ||
      paragraph.querySelector('img,svg,br,a,sup,sub,[id],[data-docx-tab],[data-docx-bookmark],[data-word-overlap],[data-equation],mark[data-comment-ids]')) continue;
    const literal = paragraph.textContent || '', style = getComputedStyle(paragraph);
    const bounds = cell.getBoundingClientRect(), hanging = fit.hanging;
    const tolerance = (parseFloat(getComputedStyle(cell).borderLeftWidth) || 0) +
      (parseFloat(getComputedStyle(cell).borderRightWidth) || 0) + .75;
    if (source.fingerprint !== hwpWrapFingerprint(literal) ||
        source.ends.at(-1) !== normalizeHwpWrapText(literal).length ||
        style.transform !== 'none' || style.direction !== 'ltr' || style.writingMode !== 'horizontal-tb' ||
        Math.abs(parseFloat(style.marginLeft) - hanging) > .08 || Math.abs(parseFloat(style.textIndent) + hanging) > .08 ||
        parseFloat(style.marginRight) !== 0 || Math.abs(bounds.width - width.outer) > tolerance ||
        !matchesAuthoredScale(paragraph, fit.scale, fit.size)) continue;
    const points = sourceLineBoundaries(paragraph, source.ends.slice(0, -1));
    if (!points) continue;

    const oldChildren = [...paragraph.childNodes], oldStyle = paragraph.style.cssText;
    const continuation = lineOpportunities(paragraph, points, range);
    for (const opportunity of continuation.querySelectorAll('wbr')) opportunity.replaceWith(document.createElement('br'));
    const ratio = fit.scale / 100;
    paragraph.style.width = `calc(${100 / ratio}% - ${hanging / ratio}px)`;
    paragraph.style.textIndent = `${-hanging / ratio}px`;
    paragraph.style.transform = `scaleX(${ratio})`;
    paragraph.style.transformOrigin = 'left center';
    paragraph.replaceChildren(continuation);
    const used = cell.getBoundingClientRect();
    const bands = [...paragraph.children].filter(line => line.tagName !== 'BR');
    const safe = bands.length === source.ends.length && bands.every((line, index) => {
      if (normalizeHwpWrapText(line.textContent || '').length !== source.ends[index] - (source.ends[index - 1] ?? 0))
        return false;
      range.selectNodeContents(line);
      const rects = [...range.getClientRects()].filter(rect => rect.width > 0 && rect.height > 0);
      return rects.length > 0 && rects.every(rect => rect.left >= used.left && rect.right <= used.right) &&
        Math.max(...rects.map(rect => rect.top)) < Math.min(...rects.map(rect => rect.bottom));
    });
    if (!safe || paragraph.textContent !== literal || Math.abs(used.width - bounds.width) > 1 / 64) {
      paragraph.replaceChildren(...oldChildren);
      paragraph.style.cssText = oldStyle;
    }
  }
}
