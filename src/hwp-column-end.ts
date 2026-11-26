import { type HwpBodyLine } from './hwp-body-line';
import { paragraphFitBottom } from './paragraph-fit';

let canvas: CanvasRenderingContext2D | null | undefined;

/** An authored break can be the following child or close the column group. */
export function followsWithHwpColumnBreak(paragraph: HTMLElement): boolean {
  const group = paragraph.parentElement;
  if (paragraph.tagName !== 'P' || !group?.matches('[data-column-count]:not([data-column-count="1"])')) return false;
  const next = paragraph.nextElementSibling ?? group.nextElementSibling;
  return !!next?.matches('[data-column-break],[data-page-break]');
}

/** Measure the painted tail of a native single line before a forced break.
 * The detached paragraph leaves the editor's DOM and selection untouched. */
export function measureHwpColumnEnd(p: HTMLElement, source: HwpBodyLine, scale: number): number | null {
  if (p.hasAttribute('data-print-paragraph-fragment')) return null;
  const width = (p.getClientRects()[0]?.width ?? 0) / scale;
  if (!(width > 0)) return null;
  canvas ??= document.createElement('canvas').getContext('2d');
  if (!canvas) return null;

  const host = document.createElement('div'), line = p.cloneNode(true) as HTMLElement;
  host.className = 'tiptap'; host.setAttribute('aria-hidden', 'true');
  host.style.cssText = `position:fixed;visibility:hidden;pointer-events:none;left:-100000px;top:0;min-height:0;padding:0;width:${width}px`;
  const style = getComputedStyle(p);
  Object.assign(line.style, { margin: '0', fontFamily: style.fontFamily, fontSize: style.fontSize, fontWeight: style.fontWeight,
    fontStyle: style.fontStyle, letterSpacing: style.letterSpacing, wordSpacing: style.wordSpacing });
  delete line.dataset.hwpColumnEnd;
  line.style.removeProperty('--hwp-column-em'); line.style.removeProperty('--hwp-column-height');
  host.append(line); document.body.append(host);
  try {
    const natural = line.getBoundingClientRect();
    if (Math.abs(natural.height - source.base) > .5 || paragraphFitBottom(line) >= natural.bottom) return null;
    const probe = document.createElement('span');
    probe.style.cssText = 'display:inline-block;vertical-align:baseline;height:0;width:0;line-height:0;font-size:0;padding:0;border:0;margin:0';
    line.dataset.hwpColumnEnd = 'true'; line.style.setProperty('--hwp-column-em', `${source.font}px`);
    line.append(probe);
    const baseline = probe.getBoundingClientRect().bottom - line.getBoundingClientRect().top;
    const walker = document.createTreeWalker(line, NodeFilter.SHOW_TEXT);
    let descent = -Infinity;
    while (walker.nextNode()) {
      const text = walker.currentNode.textContent;
      if (!text?.trim()) continue;
      const css = getComputedStyle(walker.currentNode.parentElement!);
      const size = parseFloat(css.fontSize);
      if (size > (source.maxFont ?? source.font) + .1 || css.textDecorationLine !== 'none' || css.textShadow !== 'none') return null;
      canvas.font = `${css.fontStyle} ${css.fontWeight} ${css.fontSize} ${css.fontFamily}`;
      const metrics = canvas.measureText(text);
      descent = Math.max(descent, metrics.actualBoundingBoxDescent);
    }
    const height = baseline + descent + .25;
    return height > 0 && height < source.base ? height : null;
  } finally { host.remove(); }
}
