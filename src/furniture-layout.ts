import {inlineDrawingAnchor,prepareDrawingGroupAnchors} from './drawing-group';
import {prepareStaticTextBoxes,layoutTextBoxes} from './textbox-frame';
import { imageFloating } from './image-floating';
import { documentScale } from './document-scale';
import { layoutTabs } from './document-tabs';

/** Materialize an unambiguous source grid before table-layout:fixed runs.
 * Raw rich stories have colwidth attributes rather than editor TableViews;
 * otherwise a framed header gets equal columns and can enlarge the body inset.
 * The source HTML and authored relative/ambiguous widths remain unchanged. */
function materializeFurnitureTableGrids(paper: HTMLElement) {
  for (const table of paper.querySelectorAll<HTMLTableElement>('.document-furniture table[data-word-table-layout="fixed"]')) {
    if (table.querySelector(':scope > colgroup')) continue;
    const first = table.rows[0]; if (!first) continue;
    const columns = [...first.cells].map(cell => {
      const widths = (cell.getAttribute('colwidth') || '').split(',').map(Number);
      return widths.length === cell.colSpan && widths.every(width => Number.isFinite(width) && width > 0 && width < 10000) ? widths : null;
    });
    if (!columns.length || columns.some(widths => !widths)) continue;
    const widths = columns.flat() as number[], total = widths.reduce((sum, width) => sum + width, 0);
    if (widths.length > 1000 || total > 50000 || table.style.width &&
        (!/^\d+(?:\.\d+)?px$/.test(table.style.width) || Math.abs(parseFloat(table.style.width) - total) > .5)) continue;
    const group = document.createElement('colgroup');
    for (const width of widths) { const col = document.createElement('col'); col.style.width = `${width}px`; group.append(col); }
    table.prepend(group);
    if (!table.style.width) table.style.width = `${total}px`;
  }
}

/** Header/footer drawings are anchored to their paper, independently of the
 * story's text flow. Keep margin-spanning rules from enlarging the body inset. */
export function layoutFurnitureImages(paper: HTMLElement) {
  materializeFurnitureTableGrids(paper);
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
