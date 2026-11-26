/** At display scaling above 100%, the browser can round a regular parent
 * and bold inline text to the same baseline even though their line box needs
 * more room at 96 DPI. Measure that box without replacing the printed runs. */
export function reserveWordTabbedStruts(body: HTMLElement, visited: WeakSet<HTMLElement>) {
  const zoom = 1 / devicePixelRatio;
  const range = document.createRange();
  for (const p of body.querySelectorAll<HTMLElement>(':scope > p[data-word-line-rule="auto"]')) {
    if (visited.has(p)) continue;
    visited.add(p);
    if (!p.querySelector('u,[data-word-underline]') || p.querySelectorAll('[data-docx-tab]').length !== 1 ||
        p.hasAttribute('data-print-paragraph-fragment') || p.style.height || p.style.minHeight ||
        p.querySelector('br,img,svg,sup,sub,[data-note-reference],[data-document-field],[data-symbol-font],[data-word-overlap]')) continue;
    const text = p.textContent ?? '';
    if (text.length > 120 || !/^[\x20-\x7e\t\u00a0\u2019]+$/u.test(text)) continue;
    const style = getComputedStyle(p);
    if (style.direction !== 'ltr' || style.writingMode !== 'horizontal-tb' || style.fontWeight !== '400' ||
        p.currentCSSZoom !== 1 || style.boxSizing !== 'border-box') continue;
    const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
    let count = 0, hasOnlyBoldText = true;
    while (walker.nextNode()) {
      const node = walker.currentNode, mark = node.parentElement;
      if (!mark || !node.textContent?.trim() || mark.closest('[data-docx-tab]')) continue;
      const run = getComputedStyle(mark);
      if (!mark.closest('u,[data-word-underline]') || run.fontWeight !== '700' || run.fontSize !== style.fontSize ||
          run.fontFamily !== style.fontFamily || run.fontStyle !== style.fontStyle) {
        hasOnlyBoldText = false; break;
      }
      count++;
    }
    if (!hasOnlyBoldText || !count) continue;
    const box = p.getBoundingClientRect(), line = parseFloat(style.lineHeight);
    if (!(box.width > 0 && line > 0) || Math.abs(box.height - line) > .12) continue;
    const singleBand = (node: HTMLElement) => {
      range.selectNodeContents(node);
      const rects = [...range.getClientRects()].filter(rect => rect.width > .1 && rect.height > .1);
      return rects.length > 0 && Math.max(...rects.map(rect => rect.top)) < Math.min(...rects.map(rect => rect.bottom));
    };
    if (!singleBand(p)) continue;
    const holder = document.createElement('div'), copy = p.cloneNode(true) as HTMLElement;
    holder.className = 'tiptap print-content';
    holder.style.cssText = 'position:fixed;left:-100000px;top:0;width:' + box.width + 'px;visibility:hidden;margin:0;padding:0;height:auto';
    holder.lang = p.closest('[lang]')?.getAttribute('lang') ?? 'en-US';
    Object.assign(copy.style, { width:box.width + 'px', margin:'0', fontSize:style.fontSize,
      fontFamily:style.fontFamily, fontWeight:style.fontWeight, fontStyle:style.fontStyle, lineHeight:style.lineHeight, zoom:String(zoom) });
    holder.append(copy);
    document.body.append(holder);
    try {
      if (Math.abs(copy.currentCSSZoom - zoom) > .000001 || !singleBand(copy)) continue;
      const height = copy.getBoundingClientRect().height / zoom;
      const gap = height - box.height;
      // At the normal font size the witnessed additional reserve is one
      // physical reference pixel, within Chromium's 1/64px layout rounding.
      if (gap >= .75 && gap <= 1.25) p.style.minHeight = height + 'px';
    } finally { holder.remove(); }
  }
}
