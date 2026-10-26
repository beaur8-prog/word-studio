import { imageFloating } from './image-floating';
import { documentScale } from './document-scale';

/** Header/footer drawings are anchored to their paper, independently of the
 * story's text flow. Keep margin-spanning rules from enlarging the body inset. */
export function layoutFurnitureImages(paper: HTMLElement) {
  const scale = documentScale(paper), sheet = paper.getBoundingClientRect();
  for (const image of paper.querySelectorAll<HTMLImageElement>('.document-furniture img[data-image-floating]')) {
    const geometry = imageFloating(image.dataset.imageFloating); if (!geometry) continue;
    const x = geometry.horizontalPosition, y = geometry.verticalPosition;
    image.style.position = 'absolute'; image.style.margin = '0'; image.style.maxWidth = 'none';
    const parent = image.offsetParent as HTMLElement | null; if (!parent) continue;
    const rect = parent.getBoundingClientRect();
    if (x.relative === 'page' && x.offset !== undefined) image.style.left = `${x.offset / 9525 - (rect.left - sheet.left) / scale}px`;
    if (y.relative === 'page' && y.offset !== undefined) image.style.top = `${y.offset / 9525 - (rect.top - sheet.top) / scale}px`;
    else if (y.relative === 'paragraph' && y.offset !== undefined) image.style.top = `${y.offset / 9525}px`;
  }
}
