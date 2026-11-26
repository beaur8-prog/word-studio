import { paperInset } from './paper-inset';
import { imageFloating, imageOutsidePaper, recoverableCellImage } from './image-floating';

const paragraph = 'p,h1,h2,h3,h4,h5,h6';
const originalMinimum = 'data-floating-cell-original-min-height';

/** View-only reservations must shrink again when text or a drawing changes. */
export function restoreCellPictureReservations(root: HTMLElement) {
  for (const block of root.querySelectorAll<HTMLElement>(`[${originalMinimum}]`)) {
    block.style.minHeight = block.getAttribute(originalMinimum) || '';
    block.removeAttribute(originalMinimum);
  }
}

/** Word layoutInCell positions page-relative wrapping artwork in the cell's
 * content group. Include the drawing in that group's natural height, rather
 * than measuring its offset from the physical paper or allocating it once per
 * paragraph. Restrict this path to single drawings and ordinary paragraph
 * groups; exact rows, complex nested blocks and intentionally off-paper art
 * retain their existing layout policy. */
export function cellRelativePicture(img: HTMLImageElement, root: HTMLElement) {
  const geometry = imageFloating(img.dataset.imageFloating);
  const cell = img.closest<HTMLElement>('td,th');
  const owner = img.closest<HTMLElement>(paragraph);
  if (!geometry || geometry.layoutInCell !== true || !geometry.wrap?.type ||
      geometry.behindDocument || geometry.drawingGroup || !cell || !owner ||
      imageOutsidePaper(img) || recoverableCellImage(img) ||
      geometry.verticalPosition.relative !== 'page' || geometry.verticalPosition.align ||
      cell.closest('tr')?.dataset.wordRowRule === 'exact' || cell.closest('.document-textbox,[data-column-count]')) return;
  const blocks = [...cell.children].filter((node): node is HTMLElement =>
    node instanceof HTMLElement && !node.classList.contains('ProseMirror-widget'));
  if (!blocks.length || blocks.some(block => !block.matches(paragraph)) ||
      cell.querySelectorAll('img[data-image-floating]').length !== 1) return;
  const width = Number(img.getAttribute('width')), height = Number(img.getAttribute('height'));
  const offset = (geometry.verticalPosition.offset || 0) / 9525;
  const sheet = img.closest<HTMLElement>('[data-page-height],.print-page') || root;
  const style = getComputedStyle(sheet);
  const bodyHeight = (Number(sheet.dataset.pageHeight) || parseFloat(style.height) || 1122.52) -
    paperInset(sheet.dataset.margin, 96) - paperInset(sheet.dataset.bottomMargin, 96);
  if (!(width > 0 && height > 0 && offset >= 0 && offset + height <= bodyHeight + .5 &&
      width <= owner.clientWidth + .5)) return;
  return { cell, first: blocks[0], last: blocks.at(-1)!, offset, height };
}

export function reserveCellPicture(frame: NonNullable<ReturnType<typeof cellRelativePicture>>) {
  const { first, last, offset, height } = frame;
  if (!last.hasAttribute(originalMinimum)) last.setAttribute(originalMinimum, last.style.minHeight);
  last.style.minHeight = last.getAttribute(originalMinimum) || '';
  // Re-measure after restoring the authored minimum: previously wrapped text
  // already occupies space and must not be added to an old drawing reserve.
  const start = first.getBoundingClientRect(), end = last.getBoundingClientRect();
  const extra = Math.max(0, offset + height - (end.bottom - start.top));
  if (extra > .01) last.style.minHeight = `${end.height + extra}px`;
  return first.getBoundingClientRect().top + offset;
}

export function cellPictureReservation(block: HTMLElement) {
  return block.hasAttribute(originalMinimum) ? {
    minimumHeight: parseFloat(block.style.minHeight) || undefined,
    originalMinimum: block.getAttribute(originalMinimum) || '',
  } : undefined;
}
