import {inlineDrawingAnchor,prepareDrawingGroupAnchors} from './drawing-group';
import {prepareStaticTextBoxes,layoutTextBoxes} from './textbox-frame';
import { imageFloating } from './image-floating';
import { documentScale } from './document-scale';
import { layoutTabs } from './document-tabs';

/** Header/footer drawings are anchored to their paper, independently of the
 * story's text flow. Keep margin-spanning rules from enlarging the body inset. */
export function layoutFurnitureImages(paper: HTMLElement) {
  prepareDrawingGroupAnchors(paper);
  prepareStaticTextBoxes(paper);
  const scale = documentScale(paper), sheet = paper.getBoundingClientRect();
  for (const image of paper.querySelectorAll<HTMLImageElement>('.document-furniture img[data-image-floating]')) {
    const geometry = imageFloating(image.dataset.imageFloating); if (!geometry) continue;
    const x = geometry.horizontalPosition, y = geometry.verticalPosition;
    image.style.position = 'absolute'; image.style.margin = '0'; image.style.maxWidth = 'none';
    const parent = image.offsetParent as HTMLElement | null; if (!parent) continue;
    const rect = parent.getBoundingClientRect();
    const inline=geometry.drawingGroup?.position==='inline' && inlineDrawingAnchor(paper,geometry.drawingGroup.id)?.getBoundingClientRect();
    if(inline){image.style.left=`${(inline.left-rect.left)/scale+(x.offset||0)/9525-geometry.drawingGroup!.x}px`;image.style.top=`${(inline.top-rect.top)/scale+(y.offset||0)/9525-geometry.drawingGroup!.y}px`;continue;}
    if (x.relative === 'page' && x.offset !== undefined) image.style.left = `${x.offset / 9525 - (rect.left - sheet.left) / scale}px`;
    if (y.relative === 'page' && y.offset !== undefined) image.style.top = `${y.offset / 9525 - (rect.top - sheet.top) / scale}px`;
    else if (y.relative === 'paragraph' && y.offset !== undefined) image.style.top = `${y.offset / 9525}px`;
  }
  // Source tabs can separate a floating logo from its title or page fields.
  // Resolve them after drawings leave the text flow, in every paper/story.
  for (const story of paper.querySelectorAll<HTMLElement>('.document-furniture')) layoutTabs(story);
  layoutTextBoxes(paper);
}
