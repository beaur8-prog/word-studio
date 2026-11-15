import { documentScale } from './document-scale';
import { parseTabStops, parseWordNumberingTabs } from './document-tabs';
import { cssPixels } from './hwp-cell-line';

// Legacy glyph metrics remain separate from paragraph font metrics. Keep
// their native strut/geometry while rendering the resolved Unicode label.
export const unresolvedWordListMarker = (label: unknown) => typeof label === 'string' && /[\uE000-\uF8FF]/.test(label);

/** A source list tab at the body margin removes the browser's extra gutter
 * from continued lines. Keep the native marker and the first glyph in place. */
export function nativeWordListIndents(item: HTMLElement, paragraph: HTMLElement) {
  if (!paragraph.dataset.wordNumberingTabs || !item.dataset.listGlyph || item.dataset.listFormat !== 'bullet' ||
      (item.dataset.listSuffix && item.dataset.listSuffix !== 'tab') ||
      paragraph.tagName !== 'P' ||
      item.closest('li li,td,th,.document-furniture,[data-document-note],[data-document-textbox],[data-column-count]:not([data-column-count="1"])') ||
      paragraph.querySelector('[data-docx-tab],img:not(.ProseMirror-separator),[data-note-reference]') ||
      parseTabStops(paragraph.dataset.tabStops).length) return null;
  let first = item.firstElementChild;
  while (first?.matches('.pagination-gap,.pagination-cell-gap,[data-flow-widget]')) first = first.nextElementSibling;
  if (paragraph !== first) return null;
  // A split can put an explicit tab or other blocker in the other fragment.
  // A fragment keeps the full paragraph's eligibility during fitting probes.
  if (paragraph.hasAttribute('data-print-paragraph-fragment') &&
      !paragraph.hasAttribute('data-word-native-list-indent')) return null;
  const stops = parseWordNumberingTabs(paragraph.dataset.wordNumberingTabs);
  const left = cssPixels(paragraph.style.marginLeft);
  if (stops.length !== 1 || left === null || Math.round(left * 15) !== stops[0]) return null;
  const continuation = item.hasAttribute('data-print-list-continuation');
  const originalIndent = paragraph.style.textIndent;
  const hanging = continuation ? 0 : originalIndent.startsWith('-') ? cssPixels(originalIndent.slice(1)) : null;
  const list = item.parentElement;
  if (!list?.matches('ol,ul') || hanging === null) return null;
  const paragraphStyle = getComputedStyle(paragraph);
  const markerStyle = getComputedStyle(item);
  if (paragraphStyle.direction !== 'ltr' || paragraphStyle.writingMode !== 'horizontal-tb' ||
      !['left', 'start'].includes(paragraphStyle.textAlign) || markerStyle.listStylePosition !== 'outside') return null;
  const gutter = parseFloat(getComputedStyle(list).paddingLeft) || 0;
  if (gutter <= 0 || left < gutter) return null;
  return { left: left - gutter, first: continuation ? 0 : gutter - hanging };
}

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
