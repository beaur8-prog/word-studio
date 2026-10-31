import { Extension } from '@tiptap/core';
import { Plugin, PluginKey } from '@tiptap/pm/state';
import { Decoration, DecorationSet } from '@tiptap/pm/view';

const metrics = new Map<string, number>();
const legacyKorean = /^(돋움체?|굴림체?|Dotum(?:Che)?|Gulim(?:Che)?)$/i;
function singleLine(style: CSSStyleDeclaration, text = 'Ag') {
  const font = `${style.fontStyle} ${style.fontWeight} ${style.fontSize} ${style.fontFamily}`;
  // Measure only glyphs that this run actually uses. An unconditional Hangul
  // sample selects a taller fallback even for entirely Latin Word paragraphs.
  // Collapse explicit breaks: this is a single-line metric, not paragraph height.
  const sample = text.replace(/[\t\r\n\u2028\u2029]/g, ' ');
  const key = `${font}\n${sample}`;
  const cached = metrics.get(key);
  if (cached !== undefined) {
    metrics.delete(key); metrics.set(key, cached); return cached;
  }
  // Windows Word's single-line box for these legacy Korean faces is about
  // 1.30 em (10 pt / 280-auto produces a measured 20.16–20.32 px baseline
  // pitch). Chromium's bitmap/outline normal box is only about 1.15 em.
  // Apply this measured font profile only to imported OOXML auto/minimum
  // spacing; exact spacing and ordinary authored HTML remain unchanged.
  const family = style.fontFamily.split(',')[0].replace(/["']/g, '').trim();
  let height: number;
  if (legacyKorean.test(family)) {
    // This actual character selects Word's taller Noto Sans KR fallback.
    height = parseFloat(style.fontSize) * (sample.includes('\uff65') ? 1.55 : 1.3);
  } else {
    const probe = document.createElement('span');
    Object.assign(probe.style, { position: 'fixed', left: '-10000px', visibility: 'hidden',
      font, lineHeight: 'normal', whiteSpace: 'pre', padding: '0', border: '0' });
    probe.textContent = sample; document.body.append(probe);
    height = probe.getBoundingClientRect().height; probe.remove();
  }
  // Text is part of the key so changing scripts cannot reuse another fallback's
  // height. Bound both the number and size of retained samples in large files.
  if (sample.length <= 4096) {
    if (metrics.size >= 2048) metrics.delete(metrics.keys().next().value!);
    metrics.set(key, height);
  }
  return height;
}

/** OOXML auto spacing multiplies a single font line, not the font's em size.
 * Keep the source multiplier in the document; measured pixels are view-only. */
export function wordLineHeight(paragraph: HTMLElement) {
  const rule = paragraph.dataset.wordLineRule, source = paragraph.dataset.wordLine;
  if (!rule || !source || rule === 'exact') return null;
  const value = parseFloat(source);
  if (!Number.isFinite(value) || value <= 0) return null;
  // The paragraph contributes its font strut, including for empty paragraphs.
  // Then measure text with its own resolved run font, never with an ancestor's
  // font that may not support those glyphs (e.g. a smaller CJK run in Latin text).
  let normal = singleLine(getComputedStyle(paragraph));
  const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT);
  while (walker.nextNode()) {
    const node = walker.currentNode as Text;
    if (node.data && node.parentElement)
      normal = Math.max(normal, singleLine(getComputedStyle(node.parentElement), node.data));
  }
  return rule === 'auto' ? normal * value : Math.max(normal, value * (source.endsWith('pt') ? 4 / 3 : 1));
}
export function layoutWordLines(root: HTMLElement) {
  for (const p of root.querySelectorAll<HTMLElement>('[data-word-line-rule]')) {
    const height = wordLineHeight(p);
    if (height !== null) p.style.lineHeight = `${height}px`;
  }
}

const key = new PluginKey('wordLineMetrics');
export const WordLineMetrics = Extension.create({
  name: 'wordLineMetrics',
  addGlobalAttributes: () => [{ types: ['paragraph', 'heading'], attributes: {
    wordLineRule: { default: null, parseHTML: el => el.getAttribute('data-word-line-rule'),
      renderHTML: a => ['auto', 'atLeast', 'exact'].includes(a.wordLineRule)
        ? { 'data-word-line-rule': a.wordLineRule, 'data-word-line': a.lineSpacing } : {} },
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
          if (!node.attrs.wordLineRule) return;
          const dom = view.nodeDOM(pos);
          if (!(dom instanceof HTMLElement)) return;
          const height = wordLineHeight(dom);
          if (height === null) return;
          decorations.push(Decoration.node(pos, pos + node.nodeSize, { style: `line-height:${height}px` }));
          values.push(`${pos}:${height}`);
        });
        const next = values.join(',');
        if (next !== signature || lastDocument !== view.state.doc) {
          signature = next; lastDocument = view.state.doc;
          view.dispatch(view.state.tr.setMeta(key, DecorationSet.create(view.state.doc, decorations)).setMeta('addToHistory', false));
        }
      }); };
      const fontsChanged = () => { metrics.clear(); schedule(); };
      document.fonts.addEventListener('loadingdone', fontsChanged); schedule();
      return { update: schedule, destroy() { cancelAnimationFrame(frame); document.fonts.removeEventListener('loadingdone', fontsChanged); } };
    },
  })],
});
