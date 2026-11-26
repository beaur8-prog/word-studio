import { documentScale } from './document-scale';

// Legacy glyph metrics remain separate from paragraph font metrics. Keep
// their native strut/geometry while rendering the resolved Unicode label.
export const unresolvedWordListMarker = (label: unknown) => typeof label === 'string' && /[\uE000-\uF8FF]/.test(label);

let markerContext: CanvasRenderingContext2D | null | undefined;
/** Word's hanging indent positions the label, not the first body text.
 * Keep the stored paragraph indent and cancel only the HTML list gutter.
 * Wide labels retain the native marker until numbering tabs are represented. */
export function wordListGeometry(item: HTMLElement, paragraph: HTMLElement, originalIndent: unknown) {
  if (!item.dataset.listLabel || unresolvedWordListMarker(item.dataset.listLabel)) return null;
  if (item.dataset.listSuffix && item.dataset.listSuffix !== 'tab') return null;
  if (typeof originalIndent !== 'string' || !/^-\d+(?:\.\d+)?(?:px|pt)$/.test(originalIndent)) return null;
  const hanging = parseFloat(originalIndent) * (originalIndent.endsWith('pt') ? 4 / 3 : 1);
  if (!(hanging < 0) || !item.parentElement?.matches('ol,ul')) return null;
  const css = getComputedStyle(paragraph);
  markerContext ??= document.createElement('canvas').getContext('2d');
  if (!markerContext) return null;
  markerContext.font = `${css.fontStyle} ${css.fontWeight} ${css.fontSize} ${getComputedStyle(item, '::marker').fontFamily}`;
  markerContext.fontKerning = 'none';
  if (markerContext.measureText(item.dataset.listGlyph || item.dataset.listLabel || '').width > -hanging - .5) return null;
  let top = 0;
  // Ordinary first paragraph margins collapse through the unpadded LI. Page
  // gap widgets can precede it, so measure only that exceptional first line.
  if (item.firstElementChild !== paragraph || paragraph.firstElementChild?.matches('.pagination-gap')) {
    const scale = documentScale(item);
    top = (paragraph.getBoundingClientRect().top - item.getBoundingClientRect().top) / scale;
    if (paragraph.firstElementChild?.matches('.pagination-gap'))
      top += paragraph.firstElementChild.getBoundingClientRect().height / scale;
  }
  const gutter = parseFloat(getComputedStyle(item.parentElement).paddingLeft) || 0;
  return `margin-left:${-gutter}px;--word-list-marker-left:${parseFloat(css.marginLeft) + hanging}px;--word-list-marker-top:${top}px`;
}
