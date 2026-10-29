import { Extension } from '@tiptap/core';
import { Plugin, PluginKey } from '@tiptap/pm/state';
import { Decoration, DecorationSet } from '@tiptap/pm/view';
import { activeHwpCellLine, encodeHwpCellLine, hwpCellAdvance, hwpCellFontProfile, parseHwpCellLine, type HwpCellLine } from './hwp-cell-line';
import { applyHwpCellTerminal, hwpCellPrefixStyle, hwpCellTerminalLayout, unwrapHwpCellPrefixes } from './hwp-cell-terminal';

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
function viewHeight(p: HTMLElement, line: HwpCellLine, font: number, mixedGrowth = false): number {
  const advance = hwpCellAdvance(line, font);
  // Inline objects must be allowed their full line budget. Mixed-size text
  // runs may still share a single baseline and retain its terminal height.
  const simple = !p.querySelector('img:not(.ProseMirror-separator),sub,sup,[data-note-reference],[data-word-overlap],[data-page-number]');
  return simple && !mixedGrowth && !multipleLines(p) ? Math.min(advance, line.end * font / line.font)
    : line.rule === 'Fixed' ? advance : Math.max(advance, font);
}
export function layoutHwpCellLines(root: HTMLElement) {
  const paragraphs = [...(root.matches('[data-hwp-cell-line]') ? [root] : []), ...root.querySelectorAll<HTMLElement>('[data-hwp-cell-line]')];
  for (const p of paragraphs) {
    const line = parseHwpCellLine(p.dataset.hwpCellLine);
    if (!line || !p.closest('td,th')) continue;
    unwrapHwpCellPrefixes(p);
    delete p.dataset.hwpCellTerminal; delete p.dataset.hwpCellAdvance; delete p.dataset.hwpCellEnd;
    const inherited = p.dataset.hwpCellSourceSize || p.style.fontSize;
    if (!p.dataset.hwpCellSourceSize) p.dataset.hwpCellSourceSize = inherited;
    const content = [], walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
    while (walker.nextNode()) if (walker.currentNode.textContent?.length && !walker.currentNode.parentElement?.closest('.pagination-gap,.pagination-cell-gap,[data-flow-widget]')) {
      let parent = walker.currentNode.parentElement, size = inherited;
      while (parent && parent !== p) { if (parent.style.fontSize) { size = parent.style.fontSize; break; } parent = parent.parentElement; }
      content.push({ type: 'text', text: walker.currentNode.textContent, marks: [{ type: 'textStyle', attrs: { fontSize: size } }] });
    }
    const profile = hwpCellFontProfile({ attrs: { paragraphSize: inherited }, content }, line), { font } = profile;
    p.style.fontSize = `${profile.uniform ? font : profile.inherited}px`;
    const advance = hwpCellAdvance(line, font), end = Math.min(advance, line.end * font / line.font);
    p.style.lineHeight = `${end}px`;
    p.style.setProperty('--exact-line-height', `${end}px`);
    const height = p.dataset.hwpCellContinuation || p.dataset.hwpCellFullLine ? advance : viewHeight(p, line, font, !profile.uniform && font > line.font + .05);
    p.style.lineHeight = `${height}px`;
    p.style.setProperty('--exact-line-height', `${height}px`);
    const terminal = !p.dataset.hwpCellFullLine && (profile.uniform || font <= line.font + .05) && hwpCellTerminalLayout(p, font, advance, end);
    if (terminal) applyHwpCellTerminal(p, terminal, advance, end);
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
          const profile = hwpCellFontProfile(node.toJSON(), line), { font } = profile;
          const advance = hwpCellAdvance(line, font), end = Math.min(advance, line.end * font / line.font);
          const terminal = (profile.uniform || font <= line.font + .05) && hwpCellTerminalLayout(dom, font, advance, end);
          const height = terminal ? end : viewHeight(dom, line, font, !profile.uniform && font > line.font + .05);
          const attributes: Record<string, string> = { style: `font-size:${profile.uniform ? font : profile.inherited}px;line-height:${height}px;--exact-line-height:${height}px` };
          if (terminal) Object.assign(attributes, { 'data-hwp-cell-terminal': 'true', 'data-hwp-cell-advance': String(advance), 'data-hwp-cell-end': String(end) });
          decorations.push(Decoration.node(pos, pos + node.nodeSize, attributes));
          const prefixEnd = terminal ? view.posAtDOM(terminal.node, terminal.offset) : pos + 1;
          if (terminal && prefixEnd > pos + 1) decorations.push(Decoration.inline(pos + 1, prefixEnd,
            { 'data-hwp-cell-prefix': 'true', style: hwpCellPrefixStyle(advance, end) }));
          values.push(`${pos}:${font}:${height}:${prefixEnd}:${terminal ? terminal.shift : ''}`);
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
