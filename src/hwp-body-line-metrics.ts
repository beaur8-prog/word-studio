import { Extension } from '@tiptap/core';
import { Plugin, PluginKey } from '@tiptap/pm/state';
import { Decoration, DecorationSet } from '@tiptap/pm/view';
import { activeHwpBodyLine, encodeHwpBodyLine, hwpBodyTextLine, parseHwpBodyLine } from './hwp-body-line';

const excluded = 'td,th';
const objects = 'br,img,svg,table,sup,sub,[data-note-reference],[data-page-number],[data-docx-tab],[data-docx-bookmark],[data-word-overlap],[data-symbol-font]';
function eligible(p: HTMLElement) {
  return !!p.textContent?.trim() && !p.closest(excluded) && !p.querySelector(objects) &&
    ![...p.querySelectorAll<HTMLElement>('[style]')].some(span => !!span.style.lineHeight);
}
/** Static print/measurement DOM only. The editor uses decorations below, so
 * neither wrapping spans nor measured line boundaries enter the document. */
export function layoutHwpBodyLines(root: HTMLElement) {
  for (const p of root.querySelectorAll<HTMLElement>('[data-hwp-body-line]')) {
    const line = parseHwpBodyLine(p.dataset.hwpBodyLine); if (!line) continue;
    for (const span of p.querySelectorAll<HTMLElement>('[data-hwp-body-run]')) {
      delete span.dataset.hwpBodyRun; span.style.removeProperty('line-height');
      if (![...span.attributes].some(a => a.name !== 'style') && !span.style.length) span.replaceWith(...span.childNodes);
    }
    delete p.dataset.hwpBodyLayout; p.style.removeProperty('--hwp-body-ratio');
    p.style.lineHeight = `${line.base}px`;
    if (!eligible(p)) continue;
    const texts: Text[] = [], walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
    while (walker.nextNode()) if (walker.currentNode.textContent?.length) texts.push(walker.currentNode as Text);
    for (const text of texts) {
      const span = document.createElement('span'); span.dataset.hwpBodyRun = 'true';
      text.replaceWith(span); span.append(text);
    }
    p.dataset.hwpBodyLayout = 'true';
    p.style.setProperty('--hwp-body-ratio', String(line.percent / 100));
    // A paragraph-wide strut based on the large first run otherwise keeps
    // every later, smaller line at the large font's height. All real text
    // retains its font size and supplies its own percentage line box.
    p.style.lineHeight = '0';
  }
}

const key = new PluginKey('hwpBodyLineMetrics');
export const HwpBodyLineMetrics = Extension.create({
  name: 'hwpBodyLineMetrics',
  addGlobalAttributes: () => [{ types: ['paragraph', 'heading'], attributes: {
    hwpBodyLine: { default: null,
      parseHTML: el => { const line = parseHwpBodyLine(el.dataset.hwpBodyLine || el.style.getPropertyValue('--hwp-body-line')); return line ? encodeHwpBodyLine(line) : null; },
      renderHTML: attrs => { const line = activeHwpBodyLine(attrs); return line ? { 'data-hwp-body-line': encodeHwpBodyLine(line) } : {}; },
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
          const line = hwpBodyTextLine(node.toJSON()); if (!line) return;
          const dom = view.nodeDOM(pos);
          if (!(dom instanceof HTMLElement) || !eligible(dom)) return;
          decorations.push(Decoration.node(pos, pos + node.nodeSize,
            { 'data-hwp-body-layout': 'true', style: `line-height:0;--hwp-body-ratio:${line.percent / 100}` }));
          decorations.push(Decoration.inline(pos + 1, pos + node.nodeSize - 1, { 'data-hwp-body-run': 'true' }));
          values.push(`${pos}:${line.percent}:${node.nodeSize}`);
        });
        const next = values.join(',');
        if (signature !== next || lastDocument !== view.state.doc) {
          signature = next; lastDocument = view.state.doc;
          view.dispatch(view.state.tr.setMeta(key, DecorationSet.create(view.state.doc, decorations)).setMeta('addToHistory', false));
        }
      }); };
      document.fonts.addEventListener('loadingdone', schedule); schedule();
      return { update: schedule, destroy() { cancelAnimationFrame(frame); document.fonts.removeEventListener('loadingdone', schedule); } };
    },
  })],
});
