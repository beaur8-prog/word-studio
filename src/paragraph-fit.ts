import { documentScale } from './document-scale';
import { parseHwpBodyLine } from './hwp-body-line';
import { nativeEmptyLineContentHeight } from './page-empty-lines';

let context: CanvasRenderingContext2D | null | undefined;

const graphemes = new Intl.Segmenter(undefined, { granularity: 'grapheme' });
const inkMeasureScale = 16;

// Obtain the already laid-out text of each visual line. Keep the original
// whole-run bound when a cluster crosses boxes or the order is ambiguous.
function visualLineText(node: Text, rects: readonly DOMRect[]) {
  if (rects.length < 2) return;
  const ranges = rects.map(() => ({ start: -1, end: 0 }));
  const range = document.createRange();
  let line = 0;
  for (const { segment, index } of graphemes.segment(node.data)) {
    if (!segment.trim()) continue;
    range.setStart(node, index);
    range.setEnd(node, index + segment.length);
    const pieces = [...range.getClientRects()].filter(rect => rect.width > 0 && rect.height > 0);
    if (pieces.length !== 1) return;
    const [piece] = pieces;
    while (line < rects.length && piece.top > rects[line].top + .05) line++;
    const box = rects[line];
    if (!box || Math.abs(piece.top - box.top) > .05 || Math.abs(piece.bottom - box.bottom) > .05 ||
        piece.left < box.left - .05 || piece.right > box.right + .05) return;
    const bounds = ranges[line];
    if (bounds.start < 0) bounds.start = index;
    bounds.end = index + segment.length;
  }
  if (ranges.every(range => range.start >= 0)) return ranges.map(({ start, end }) => node.data.slice(start, end));
}


/** The final line's empty leading can cross the body edge while every painted
 * glyph still fits. Keep its full flow box for following paragraphs, but use
 * the text's paint bounds when deciding whether this paragraph must move. */
export function paragraphFitBottom(paragraph: HTMLElement, suppressBottomSpacing = true, nativeEmptyLine = false): number {
  const box = paragraph.getBoundingClientRect(), fallback = box.bottom;
  const hwpRunAdvances = paragraph.dataset.hwpBodyLayout === 'true' || paragraph.dataset.hwpBodyLayout === 'mixed-top';
  const hwpLineTop = paragraph.dataset.hwpBodyLayout === 'top' || paragraph.dataset.hwpBodyLayout === 'mixed-top';
  // Standard Word spacing reserves the entire exact line. Trimming its empty
  // leading is an opt-in document compatibility policy, not a universal fit.
  if (!suppressBottomSpacing || paragraph.dataset.wordLineRule && paragraph.dataset.wordLineRule !== 'exact') return fallback;
  if (!paragraph.textContent) {
    const source = parseHwpBodyLine(paragraph.dataset.hwpBodyLine);
    const measured = source?.empty && !paragraph.closest('td,th,li,[data-column-count]:not([data-column-count="1"])');
    if (nativeEmptyLine || measured) {
      const scale = documentScale(paragraph);
      const height = nativeEmptyLineContentHeight(paragraph, scale);
      if (height !== null && (nativeEmptyLine || source &&
        Math.abs(height - source.font) < .06 && Math.abs(box.height / scale - source.base) < .07))
        return Math.min(fallback, box.top + height * scale);
    }
  }
  if (!paragraph.matches('p,h1,h2,h3,h4,h5,h6') || paragraph.hasAttribute('data-print-paragraph-fragment') ||
    (!hwpRunAdvances && !/(?:px|pt)$/.test(paragraph.style.lineHeight)) ||
    !paragraph.textContent?.trim() || paragraph.querySelector('img,svg,table,br,mark,u,s,del,sub,sup,[data-note-reference],[style*="background"],[style*="text-decoration"]') ||
    paragraph.closest('[data-column-count]:not([data-column-count="1"])')) return fallback;
  const css = getComputedStyle(paragraph), scale = documentScale(paragraph), lineHeight = parseFloat(css.lineHeight);
  if (css.textShadow !== 'none' || css.backgroundColor !== 'rgba(0, 0, 0, 0)' ||
    ['paddingBottom', 'borderBottomWidth'].some(p => (parseFloat(css[p as 'paddingBottom']) || 0) > 0)) return fallback;
  if (context === undefined) context = document.createElement('canvas').getContext('2d');
  if (!context || (!hwpRunAdvances && !(lineHeight > 0)) || !(scale > 0)) return fallback;
  const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT);
  let bottom = -Infinity;
  while (walker.nextNode()) {
    const node = walker.currentNode as Text;
    if (!node.data.trim() || node.parentElement?.closest('.pagination-gap,[data-flow-widget]')) continue;
    const style = getComputedStyle(node.parentElement!);
    if (style.textShadow !== 'none' || style.backgroundColor !== 'rgba(0, 0, 0, 0)') return fallback;
    context.font = `${style.fontStyle} ${style.fontWeight} ${style.fontSize} ${style.fontFamily}`;
    const transform = (text: string) => style.textTransform === 'uppercase' ? text.toUpperCase() :
      style.textTransform === 'lowercase' ? text.toLowerCase() : text;
    const metrics = context.measureText(transform(node.data)), fontHeight = metrics.fontBoundingBoxAscent + metrics.fontBoundingBoxDescent;
    if (!(fontHeight > 0) || !Number.isFinite(metrics.actualBoundingBoxDescent)) return fallback;
    const range = document.createRange(); range.selectNodeContents(node);
    const rects = [...range.getClientRects()];
    const lines = hwpLineTop && style.direction === 'ltr' && style.writingMode === 'horizontal-tb' &&
      style.transform === 'none' && style.fontVariationSettings === 'normal' ? visualLineText(node, rects) : undefined;
    const font = context.font, fontSize = parseFloat(style.fontSize);
    for (const [index, rect] of rects.entries()) {
      if (rect.width < .1 || rect.height < 1) continue;
      // Smaller exact leading can intentionally overlap lines. It does not
      // authorize ignoring an overflowing glyph's own height.
      // Very large exact leading is an explicit spacer. Preserve its full
      // pagination budget even when a small label is its only visible ink.
      // Mixed-size HWP paragraphs disable the common strut and let their
      // text runs supply the line advance.
      const advance = hwpRunAdvances ? parseFloat(style.lineHeight) : lineHeight;
      if (!(advance > 0) || advance + .5 < rect.height / scale || advance > rect.height / scale * 2) return fallback;
      const ratio = rect.height / fontHeight, baseline = rect.top + metrics.fontBoundingBoxAscent * ratio;
      const decoration = style.textDecorationLine.includes('underline') ? parseFloat(style.fontSize) * .15 : 0;
      let runDescent = metrics.actualBoundingBoxDescent;
      const line = lines?.[index];
      if (line) {
        // Chromium rounds canvas ink metrics at the requested CSS size.
        // Measure with subpixel precision without resizing the laid-out text.
        context.font = `${style.fontStyle} ${style.fontWeight} ${fontSize * inkMeasureScale}px ${style.fontFamily}`;
        const measured = context.measureText(transform(line)).actualBoundingBoxDescent;
        context.font = font;
        if (Number.isFinite(measured)) runDescent = Math.min(runDescent, (measured + 1) / inkMeasureScale);
      }
      const descent = Math.max(runDescent, decoration);
      // DOM and canvas round font box edges differently. At a verified HWP
      // line top use the lower baseline implied by either edge of a closely
      // matching box. Scale the ink by the document zoom, not by that rounding.
      const matchingFontBox = Math.round(Math.abs(rect.height / scale - fontHeight) * 64) <= 64;
      const inkBottom = hwpLineTop && matchingFontBox
        ? Math.max(rect.top + metrics.fontBoundingBoxAscent * scale, rect.bottom - metrics.fontBoundingBoxDescent * scale) + descent * scale
        : baseline + descent * ratio + scale;
      bottom = Math.max(bottom, inkBottom);
    }
  }
  return Number.isFinite(bottom) ? Math.min(fallback, bottom) : fallback;
}
