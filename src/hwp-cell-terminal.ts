/** View-only line boxes for a simple HWP cell's terminal paragraph. Its
 * preceding lines retain full advance; the final line reserves only ink/em.
 * No text, hard break, stored margin, or row minimum is removed. */
export interface CellTerminalLayout { node: Node; offset: number; shift: number }
const profiles = new Map<string, number | null>();
const excluded = 'img,svg,table,sup,sub,mark,u,s,del,[data-note-reference],[data-word-overlap],[data-page-number],[data-docx-tab],[data-docx-bookmark],[data-symbol-font]';
const widgets = '.pagination-gap,.pagination-cell-gap,[data-flow-widget]';

export function unwrapHwpCellPrefixes(root: HTMLElement) {
  for (const prefix of root.querySelectorAll<HTMLElement>('[data-hwp-cell-prefix]')) {
    // ProseMirror decorates a hard-break leaf directly rather than wrapping
    // it in a span. It remains document content when a view is measured.
    if (prefix instanceof HTMLBRElement) {
      delete prefix.dataset.hwpCellPrefix;
      for (const property of ['line-height', 'position', 'top']) prefix.style.removeProperty(property);
    } else prefix.replaceWith(...prefix.childNodes);
  }
}
function fontKey(style: CSSStyleDeclaration) {
  return `${style.fontSize}|${style.fontFamily}|${style.fontWeight}|${style.fontStyle}|${style.fontStretch}|${style.fontVariantCaps}`;
}
function prefixShift(p: HTMLElement, run: HTMLElement, font: number, advance: number, end: number, text: string): number | null {
  const paragraphStyle = getComputedStyle(p), style = getComputedStyle(run);
  const key = `${fontKey(paragraphStyle)}|${fontKey(style)}|${font}|${advance}|${end}|${[...text][0]}`;
  if (profiles.has(key)) return profiles.get(key)!;
  const probe = document.createElement('p'), prefix = document.createElement('span'), glyph = document.createElement('span');
  probe.style.cssText = 'position:fixed;left:-100000px;top:0;width:500px;visibility:hidden;pointer-events:none;margin:0;padding:0;border:0;white-space:pre;';
  for (const property of ['font-family','font-weight','font-style','font-stretch','font-variant-caps']) probe.style.setProperty(property, paragraphStyle.getPropertyValue(property));
  probe.style.fontSize = `${font}px`; probe.style.lineHeight = `${end}px`;
  for (const property of ['font-family','font-weight','font-style','font-stretch','font-variant-caps']) prefix.style.setProperty(property, style.getPropertyValue(property));
  prefix.style.fontSize = `${font}px`; prefix.style.lineHeight = `${end}px`;
  glyph.style.lineHeight = `${Math.min(font, end)}px`; glyph.textContent = [...text].slice(0, 8).join('') || 'Hg';
  prefix.append(glyph); probe.append(prefix); document.body.append(probe);
  let shift: number | null = null;
  try {
    const range = document.createRange(); range.selectNodeContents(glyph);
    const compact = range.getBoundingClientRect().top - probe.getBoundingClientRect().top;
    const compactHeight = probe.getBoundingClientRect().height;
    prefix.style.lineHeight = `${advance}px`;
    const full = range.getBoundingClientRect().top - probe.getBoundingClientRect().top;
    if (Math.abs(compactHeight - end) <= .5 && Math.abs(probe.getBoundingClientRect().height - advance) <= .5) shift = compact - full;
  } finally { probe.remove(); }
  if (profiles.size > 512) profiles.clear(); profiles.set(key, shift); return shift;
}
const invalidate = () => profiles.clear();
if (typeof document !== 'undefined') document.fonts.addEventListener('loadingdone', invalidate);

function previousLeaf(node: Node, root: Node): Node | null {
  while (node !== root && !node.previousSibling) node = node.parentNode!;
  if (node === root) return null;
  node = node.previousSibling!;
  while (node.lastChild) node = node.lastChild;
  return node;
}
export function hwpCellTerminalLayout(p: HTMLElement, font: number, advance: number, end: number): CellTerminalLayout | null {
  if (!(advance > end + .05) || !p.textContent?.trim() || p.querySelector(excluded)) return null;
  const texts: Text[] = [], walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT);
  let last: Node | undefined;
  while (walker.nextNode()) {
    const node = walker.currentNode;
    if (node.parentElement?.closest(widgets)) continue;
    if (node instanceof HTMLBRElement && !node.classList.contains('ProseMirror-trailingBreak')) last = node;
    else if (node instanceof Text && node.data.length) { texts.push(node); if (node.data.trim()) last = node; }
  }
  // A final explicit blank line is real document content, not terminal leading.
  const continuation = !!p.dataset.hwpCellContinuation;
  if (!texts.length || !continuation && last instanceof HTMLBRElement) return null;
  const firstStyle = getComputedStyle(texts[0].parentElement!), firstFont = fontKey(firstStyle);
  if (Math.abs(parseFloat(firstStyle.fontSize) - font) > .05 || texts.some(t => fontKey(getComputedStyle(t.parentElement!)) !== firstFont)) return null;
  const shift = prefixShift(p, texts[0].parentElement!, font, advance, end, texts[0].data);
  if (shift === null) return null;
  // A continued fragment reserves a full last line, but its ink still uses
  // the same inset as the original paragraph's compact terminal strut.
  if (continuation) return { node: p, offset: p.childNodes.length, shift };
  let target: Text | null = null, lastTop = -Infinity;
  const range = document.createRange();
  for (const text of texts) {
    range.selectNodeContents(text);
    for (const rect of range.getClientRects()) if (rect.width > .1 && rect.height > 1 && rect.top > lastTop + .5) { lastTop = rect.top; target = text; }
  }
  if (!target) return null;
  const offsets = [...new Intl.Segmenter(undefined, { granularity: 'grapheme' }).segment(target.data)].map(s => s.index);
  let low = 0, high = offsets.length - 1, first = 0;
  while (low <= high) {
    const mid = (low + high) >> 1;
    range.setStart(target, offsets[mid]); range.setEnd(target, offsets[mid + 1] ?? target.length);
    if (range.getBoundingClientRect().top < lastTop - .5) { first = mid + 1; low = mid + 1; } else high = mid - 1;
  }
  let node: Node = target, offset = offsets[Math.min(first, offsets.length - 1)] ?? 0;
  if (!offset) {
    const previous = previousLeaf(node, p);
    if (previous instanceof HTMLBRElement) { node = previous.parentNode!; offset = [...node.childNodes].indexOf(previous); }
  }
  range.selectNodeContents(p); range.setEnd(node, offset);
  if (!range.toString() && !range.cloneContents().querySelector('br')) return null;
  return { node, offset, shift };
}
export const hwpCellPrefixStyle = (advance: number, shift: number) => `line-height:${advance}px!important;position:relative;top:${shift}px`;
export function applyHwpCellTerminal(p: HTMLElement, layout: CellTerminalLayout, advance: number, end: number) {
  const range = document.createRange(); range.selectNodeContents(p); range.setEnd(layout.node, layout.offset);
  const prefix = document.createElement('span'); prefix.dataset.hwpCellPrefix = 'true';
  prefix.style.cssText = hwpCellPrefixStyle(advance, layout.shift); prefix.append(range.extractContents()); range.insertNode(prefix);
  const lastHeight = p.dataset.hwpCellContinuation ? advance : end;
  p.style.lineHeight = `${lastHeight}px`; p.style.setProperty('--exact-line-height', `${end}px`);
  p.dataset.hwpCellTerminal = 'true'; p.dataset.hwpCellAdvance = String(advance); p.dataset.hwpCellEnd = String(lastHeight);
}
