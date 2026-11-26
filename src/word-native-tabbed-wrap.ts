import { paragraphLineOffsets } from './paragraph-fragments';

const sameOffsets = (a: number[], b: number[]) => a.length === b.length && a.every((value, index) => value === b[index]);

function regularTabParagraph(p: HTMLParagraphElement, style: CSSStyleDeclaration) {
  const text = p.textContent ?? '';
  if (text.length < 120 || text.length > 420 || !/^[\x20-\x7e\t\u00a0\u2019]+$/u.test(text) ||
      p.currentCSSZoom !== 1 || p.hasAttribute('data-print-paragraph-fragment') ||
      p.querySelectorAll('[data-docx-tab]').length !== 1 || p.querySelector('br,img,svg,a,sup,sub,[data-note-reference],[data-document-field],[data-page-number],[data-symbol-font],[data-word-overlap]') ||
      style.direction !== 'ltr' || style.writingMode !== 'horizontal-tb' || style.boxSizing !== 'border-box' ||
      style.textAlign !== 'left' && style.textAlign !== 'start' ||
      style.fontWeight !== '400' || style.fontStyle !== 'normal' || parseFloat(style.textIndent) >= 0 ||
      parseFloat(style.marginLeft) <= 0 || (parseFloat(style.wordSpacing) || 0) !== 0) return false;
  const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
  while (walker.nextNode()) {
    const node = walker.currentNode, parent = node.parentElement;
    if (!parent || !node.textContent?.trim() || parent.closest('[data-docx-tab]')) continue;
    const run = getComputedStyle(parent);
    if (run.fontFamily !== style.fontFamily || run.fontSize !== style.fontSize || run.fontWeight !== style.fontWeight ||
        run.fontStyle !== style.fontStyle || run.letterSpacing !== style.letterSpacing || run.wordSpacing !== style.wordSpacing) return false;
  }
  return true;
}

/** An unchanged Word tabbed paragraph can have a different wrapped word
 * because DirectWrite's fallback advance rounds at the display scale. Use
 * the observed 96 DPI break offsets, then try only the smallest inter-word
 * correction that reproduces those offsets at the native output scale. */
export function restoreWordNativeTabbedWrap(root: HTMLElement) {
  const zoom = 1 / devicePixelRatio, range = document.createRange();
  for (const paragraph of root.querySelectorAll<HTMLParagraphElement>('.print-content > p[data-word-line-rule="auto"]')) {
    const style = getComputedStyle(paragraph);
    if (!regularTabParagraph(paragraph, style)) continue;
    const offsets = paragraphLineOffsets(paragraph);
    if (offsets.length < 2 || offsets.length > 7) continue;
    const box = paragraph.getBoundingClientRect();
    if (!(box.width > 0 && box.height > 0)) continue;
    const holder = document.createElement('div'), copy = paragraph.cloneNode(true) as HTMLParagraphElement;
    holder.className = 'tiptap print-content';
    holder.lang = paragraph.closest('[lang]')?.getAttribute('lang') ?? 'en-US';
    holder.style.cssText = `position:fixed;left:-100000px;top:0;width:${box.width}px;visibility:hidden;margin:0;padding:0;height:auto`;
    Object.assign(copy.style, { width: `${box.width}px`, margin: '0', fontFamily: style.fontFamily,
      fontSize: style.fontSize, fontStyle: style.fontStyle, fontWeight: style.fontWeight,
      lineHeight: style.lineHeight, letterSpacing: style.letterSpacing, zoom: String(zoom) });
    holder.append(copy);
    document.body.append(holder);
    let reference: number[];
    try {
      if (Math.abs(copy.currentCSSZoom - zoom) > .000001) continue;
      reference = paragraphLineOffsets(copy);
    } finally { holder.remove(); }
    if (reference.length !== offsets.length || reference[0] !== offsets[0] || sameOffsets(reference, offsets) ||
        reference.some((value, index) => value < offsets[index] || value > offsets[index] + 12)) continue;
    const previous = paragraph.style.wordSpacing;
    let accepted = false;
    try {
      for (const adjustment of [1 / 64, 2 / 64]) {
        paragraph.style.wordSpacing = `${-adjustment}px`;
        if (!sameOffsets(reference, paragraphLineOffsets(paragraph)) ||
            Math.abs(paragraph.getBoundingClientRect().height - box.height) > .02) continue;
        range.selectNodeContents(paragraph);
        if ([...range.getClientRects()].every(bounds => !bounds.width || bounds.right <= box.right + .02)) {
          accepted = true;
          break;
        }
      }
    } finally { if (!accepted) paragraph.style.wordSpacing = previous; }
  }
}
