import { documentScale } from './document-scale';

let context: CanvasRenderingContext2D | null | undefined;

/** The final line's empty leading can cross the body edge while every painted
 * glyph still fits. Keep its full flow box for following paragraphs, but use
 * the text's paint bounds when deciding whether this paragraph must move. */
export function paragraphFitBottom(paragraph: HTMLElement): number {
  const box = paragraph.getBoundingClientRect(), fallback = box.bottom;
  if (!paragraph.matches('p,h1,h2,h3,h4,h5,h6') || paragraph.hasAttribute('data-print-paragraph-fragment') || !/(?:px|pt)$/.test(paragraph.style.lineHeight) ||
    !paragraph.textContent?.trim() || paragraph.querySelector('img,svg,table,br,mark,u,s,del,sub,sup,[data-note-reference],[style*="background"],[style*="text-decoration"]') ||
    paragraph.closest('[data-column-count]:not([data-column-count="1"])')) return fallback;
  const css = getComputedStyle(paragraph), scale = documentScale(paragraph), lineHeight = parseFloat(css.lineHeight);
  if (css.textShadow !== 'none' || css.backgroundColor !== 'rgba(0, 0, 0, 0)' ||
    ['paddingBottom', 'borderBottomWidth'].some(p => (parseFloat(css[p as 'paddingBottom']) || 0) > 0)) return fallback;
  if (context === undefined) context = document.createElement('canvas').getContext('2d');
  if (!context || !(lineHeight > 0) || !(scale > 0)) return fallback;
  const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT);
  let bottom = -Infinity;
  while (walker.nextNode()) {
    const node = walker.currentNode as Text;
    if (!node.data.trim() || node.parentElement?.closest('.pagination-gap,[data-flow-widget]')) continue;
    const style = getComputedStyle(node.parentElement!);
    if (style.textShadow !== 'none' || style.backgroundColor !== 'rgba(0, 0, 0, 0)') return fallback;
    context.font = `${style.fontStyle} ${style.fontWeight} ${style.fontSize} ${style.fontFamily}`;
    const text = style.textTransform === 'uppercase' ? node.data.toUpperCase() : style.textTransform === 'lowercase' ? node.data.toLowerCase() : node.data;
    const metrics = context.measureText(text), fontHeight = metrics.fontBoundingBoxAscent + metrics.fontBoundingBoxDescent;
    if (!(fontHeight > 0) || !Number.isFinite(metrics.actualBoundingBoxDescent)) return fallback;
    const range = document.createRange(); range.selectNodeContents(node);
    for (const rect of range.getClientRects()) {
      if (rect.width < .1 || rect.height < 1) continue;
      // Smaller exact leading can intentionally overlap lines. It does not
      // authorize ignoring an overflowing glyph's own height.
      // Very large exact leading is an explicit spacer. Preserve its full
      // pagination budget even when a small label is its only visible ink.
      if (lineHeight + .5 < rect.height / scale || lineHeight > rect.height / scale * 2) return fallback;
      const ratio = rect.height / fontHeight, baseline = rect.top + metrics.fontBoundingBoxAscent * ratio;
      const decoration = style.textDecorationLine.includes('underline') ? parseFloat(style.fontSize) * .15 : 0;
      bottom = Math.max(bottom, baseline + Math.max(metrics.actualBoundingBoxDescent, decoration) * ratio + scale);
    }
  }
  return Number.isFinite(bottom) ? Math.min(fallback, bottom) : fallback;
}
