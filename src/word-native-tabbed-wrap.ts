const graphemes = new Intl.Segmenter(undefined, { granularity: 'grapheme' });

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

function visibleLineGeometry(paragraph: HTMLParagraphElement, range: Range) {
  const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT);
  const box = paragraph.getBoundingClientRect(), starts: number[] = [];
  let glyphs = 0, top = -Infinity, right = -Infinity;
  while (walker.nextNode()) {
    const text = walker.currentNode as Text;
    if (text.parentElement?.closest('[data-docx-tab]')) continue;
    for (const { index, segment } of graphemes.segment(text.data)) {
      if (!segment.trim()) continue;
      range.setStart(text, index);
      range.setEnd(text, index + segment.length);
      const rect = range.getBoundingClientRect();
      if (!(rect.width > .1 && rect.height > 1)) return;
      if (Math.abs(rect.top - top) > .5) { starts.push(glyphs); top = rect.top; }
      right = Math.max(right, rect.right);
      glyphs++;
    }
  }
  return { starts, glyphs, right: right - box.right, height: box.height };
}

/** An unchanged Word tabbed paragraph can have a different wrapped word
 * because DirectWrite's fallback advance rounds at the display scale. Use
 * the observed 96 DPI break offsets, then try only the smallest inter-word
 * correction that reproduces those offsets at the native output scale. */
export async function restoreWordNativeTabbedWrap(root: HTMLElement, checkpoint?: () => void | Promise<void>) {
  const zoom = 1 / devicePixelRatio, range = document.createRange();
  for (const paragraph of root.querySelectorAll<HTMLParagraphElement>('.print-content > p[data-word-line-rule="auto"]')) {
    const style = getComputedStyle(paragraph);
    if (!regularTabParagraph(paragraph, style)) continue;
    await checkpoint?.();
    const native = visibleLineGeometry(paragraph, range);
    if (!native || native.starts.length < 2 || native.starts.length > 7) continue;
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
    let reference: ReturnType<typeof visibleLineGeometry>;
    try {
      if (Math.abs(copy.currentCSSZoom - zoom) > .000001 ||
          Math.abs(copy.getBoundingClientRect().width / zoom - box.width) > .05) continue;
      reference = visibleLineGeometry(copy, range);
    } finally { holder.remove(); }
    if (!reference || reference.right > .02 || reference.glyphs !== native.glyphs ||
        reference.starts.length !== native.starts.length || sameOffsets(reference.starts, native.starts) ||
        reference.starts.some((value, index) => value < native.starts[index] || value > native.starts[index] + 12)) continue;
    const previous = paragraph.style.wordSpacing;
    let accepted = false;
    try {
      for (const adjustment of [1 / 64, 2 / 64]) {
        paragraph.style.wordSpacing = `${-adjustment}px`;
        const measured = visibleLineGeometry(paragraph, range);
        if (measured && measured.glyphs === native.glyphs && sameOffsets(reference.starts, measured.starts) &&
            Math.abs(measured.height - box.height) <= .02 && measured.right <= .02) {
          accepted = true;
          break;
        }
      }
    } finally { if (!accepted) paragraph.style.wordSpacing = previous; }
  }
}
