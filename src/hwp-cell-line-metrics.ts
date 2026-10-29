import { Extension } from '@tiptap/core';
import { Plugin, PluginKey } from '@tiptap/pm/state';
import { Decoration, DecorationSet } from '@tiptap/pm/view';
import { activeHwpCellLine, encodeHwpCellLine, hwpCellAdvance, hwpCellFont, parseHwpCellLine, type HwpCellLine } from './hwp-cell-line';

function multipleLines(p: HTMLElement): boolean {
  if (p.querySelector('br:not(.ProseMirror-trailingBreak)')) return true;
  const seen = new Map<string, number>();
  let top: number | undefined, bottom = 0;
  const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
  while (walker.nextNode()) {
    if (!walker.currentNode.textContent?.trim()) continue;
    const style = getComputedStyle(walker.currentNode.parentElement!);
    const range = document.createRange(); range.selectNodeContents(walker.currentNode);
    for (const rect of range.getClientRects()) if (rect.width > .1 && rect.height > 1) {
      const font = `${style.font}:${rect.height}`;
      const previous = seen.get(font);
      if (previous !== undefined && Math.abs(rect.top - previous) > .75) return true;
      seen.set(font, rect.top);
      // Small and large runs share a baseline but have different glyph tops.
      // Their overlapping boxes are one line; a wrapped run of the same font
      // is identified above even before a previously compact strut expands.
      if (top !== undefined && Math.min(bottom, rect.bottom) - Math.max(top, rect.top) <= Math.min(bottom - top, rect.height) * .4) return true;
      top = Math.min(top ?? rect.top, rect.top); bottom = Math.max(bottom, rect.bottom);
    }
  }
  return false;
}
function viewHeight(p: HTMLElement, line: HwpCellLine, font: number): number {
  const advance = hwpCellAdvance(line, font);
  // Inline objects must be allowed their full line budget. Mixed-size text
  // runs may still share a single baseline and retain its terminal height.
  const simple = !p.querySelector('img,sub,sup,[data-note-reference],[data-word-overlap],[data-page-number]');
  return simple && !multipleLines(p) ? Math.min(advance, line.end * font / line.font) : Math.max(advance, font);
}
export function layoutHwpCellLines(root: HTMLElement) {
  for (const p of root.querySelectorAll<HTMLElement>('[data-hwp-cell-line]')) {
    const line = parseHwpCellLine(p.dataset.hwpCellLine);
    if (!line || !p.closest('td,th')) continue;
    const inherited = p.dataset.hwpCellSourceSize || p.style.fontSize;
    if (!p.dataset.hwpCellSourceSize) p.dataset.hwpCellSourceSize = inherited;
    const content = [], walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
    while (walker.nextNode()) if (walker.currentNode.textContent?.trim()) {
      let parent = walker.currentNode.parentElement, size = inherited;
      while (parent && parent !== p) { if (parent.style.fontSize) { size = parent.style.fontSize; break; } parent = parent.parentElement; }
      content.push({ type: 'text', text: walker.currentNode.textContent, marks: [{ type: 'textStyle', attrs: { fontSize: size } }] });
    }
    const font = hwpCellFont({ attrs: { paragraphSize: inherited }, content }, line);
    p.style.fontSize = `${font}px`;
    const end = Math.min(hwpCellAdvance(line, font), line.end * font / line.font);
    p.style.lineHeight = `${end}px`;
    p.style.setProperty('--exact-line-height', `${end}px`);
    const height = viewHeight(p, line, font);
    p.style.lineHeight = `${height}px`;
    p.style.setProperty('--exact-line-height', `${height}px`);
  }
}
const key = new PluginKey('hwpCellLineMetrics');
export const HwpCellLineMetrics = Extension.create({
  name: 'hwpCellLineMetrics',
  addGlobalAttributes: () => [{ types: ['paragraph', 'heading'], attributes: {
    hwpCellLine: { default: null,
      parseHTML: el => { const line = parseHwpCellLine(el.dataset.hwpCellLine || el.style.getPropertyValue('--hwp-cell-line')); return line ? encodeHwpCellLine(line) : null; },
      renderHTML: attrs => { const line = activeHwpCellLine(attrs); return line ? { 'data-hwp-cell-line': encodeHwpCellLine(line) } : {}; },
    },
  } }],
  addProseMirrorPlugins: () => [new Plugin({
    key,
    state: { init: () => DecorationSet.empty, apply: (tr, old) => tr.getMeta(key) ?? old.map(tr.mapping, tr.doc) },
    props: { decorations: state => key.getState(state) },
    view(view) {
      let frame = 0, signature = '', lastDocument = view.state.doc;
      const schedule = () => { cancelAnimationFrame(frame); frame = requestAnimationFrame(() => {
        if (!view.dom.isConnected || view.composing) return;
        const decorations: Decoration[] = [], values: string[] = [];
        view.state.doc.descendants((node, pos) => {
          const line = activeHwpCellLine(node.attrs);
          if (!line) return;
          const dom = view.nodeDOM(pos);
          if (!(dom instanceof HTMLElement) || !dom.closest('td,th')) return;
          const font = hwpCellFont(node.toJSON(), line), height = viewHeight(dom, line, font);
          decorations.push(Decoration.node(pos, pos + node.nodeSize, { style: `font-size:${font}px;line-height:${height}px;--exact-line-height:${height}px` }));
          values.push(`${pos}:${font}:${height}`);
        });
        const next = values.join(',');
        if (next !== signature || lastDocument !== view.state.doc) {
          signature = next; lastDocument = view.state.doc;
          view.dispatch(view.state.tr.setMeta(key, DecorationSet.create(view.state.doc, decorations)).setMeta('addToHistory', false));
        }
      }); };
      const resize = new ResizeObserver(schedule); resize.observe(view.dom);
      document.fonts.addEventListener('loadingdone', schedule); schedule();
      return { update: schedule, destroy() { cancelAnimationFrame(frame); resize.disconnect(); document.fonts.removeEventListener('loadingdone', schedule); } };
    },
  })],
});
