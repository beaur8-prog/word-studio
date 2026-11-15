import { parseTabStops } from './document-tabs';
import { Extension } from '@tiptap/core';
import { Plugin, PluginKey } from '@tiptap/pm/state';
import { documentMetadata } from './document-history';
import { isOrdinaryWordBodyParagraph, isWordAlignedBulletParagraph, isWordGroupedTextBoxParagraph } from './word-trailing-spaces';
import { wordFlowHasWrappingAnchors } from './word-keep-policy';
import { Decoration, DecorationSet } from '@tiptap/pm/view';
import { wordFallbackLineHeight, wordFallbackSpaceAdjustment, wordFontMetricsChanged } from './word-font-fallbacks';
import { nativeWordListIndents, wordListGeometry, unresolvedWordListMarker } from './word-list-layout';

const metrics = new Map<string, number>();
// The shared LRU bounds retained samples, but documents can contain more unique
// runs than it holds. Keep one measurement per live DOM node too, so unchanged
// selection updates cannot evict and remeasure the entire long document.
let nodeMetrics = new WeakMap<Node, { font: string; text: string; height: number }>();
const legacyKorean = /^(돋움체?|굴림체?|Dotum(?:Che)?|Gulim(?:Che)?)$/i;
function singleLine(style: CSSStyleDeclaration, text = 'Ag', owner?: Node) {
  const font = `${style.fontStyle} ${style.fontWeight} ${style.fontSize} ${style.fontFamily}`;
  const previous = owner && nodeMetrics.get(owner);
  if (previous && previous.font === font && previous.text === text) return previous.height;
  // Measure only glyphs that this run actually uses. An unconditional Hangul
  // sample selects a taller fallback even for entirely Latin Word paragraphs.
  // Collapse explicit breaks: this is a single-line metric, not paragraph height.
  const sample = text.replace(/[\t\r\n\u2028\u2029]/g, ' ');
  const key = `${font}\n${sample}`;
  const cached = metrics.get(key);
  if (cached !== undefined) {
    metrics.delete(key); metrics.set(key, cached);
    if (owner) nodeMetrics.set(owner, { font, text, height: cached });
    return cached;
  }
  // Windows Word's single-line box for these legacy Korean faces is about
  // 1.30 em (10 pt / 280-auto produces a measured 20.16–20.32 px baseline
  // pitch). Chromium's bitmap/outline normal box is only about 1.15 em.
  // Apply this measured font profile only to imported OOXML auto/minimum
  // spacing; exact spacing and ordinary authored HTML remain unchanged.
  const family = style.fontFamily.split(',')[0].replace(/["']/g, '').trim();
  let height: number;
  const fractional = wordFallbackLineHeight(style, sample);
  if (fractional !== null) height = fractional;
  else if (legacyKorean.test(family)) {
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
  if (owner) nodeMetrics.set(owner, { font, text, height });
  return height;
}

/** OOXML auto spacing multiplies a single font line, not the font's em size.
 * Keep the source multiplier in the document; measured pixels are view-only. */
export function wordLineHeight(paragraph: HTMLElement, plainBody = false) {
  const rule = paragraph.dataset.wordLineRule, source = paragraph.dataset.wordLine;
  if (!rule || !source || rule === 'exact') return null;
  const value = parseFloat(source);
  if (!Number.isFinite(value) || value <= 0) return null;
  // The paragraph contributes its font strut, including for empty paragraphs.
  // Then measure text with its own resolved run font, never with an ancestor's
  // font that may not support those glyphs (e.g. a smaller CJK run in Latin text).
  const style = getComputedStyle(paragraph);
  let normal = singleLine(style, 'Ag', paragraph);
  const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT);
  const outline = paragraph.dataset.wordOutlineLevel;
  const ordinary = plainBody && rule === 'auto' && paragraph.matches('p[data-word-paragraph-style]') &&
    (outline === undefined || outline === '9') && style.direction === 'ltr' && style.writingMode === 'horizontal-tb' &&
    !paragraph.closest('li,table,blockquote,[data-hwp-body-line],[data-hwp-cell-line],.document-furniture,.print-footnotes,.pagination-headers,[data-document-note],[data-document-textbox],[data-textbox-content],[data-column-count]:not([data-column-count="1"])') &&
    !paragraph.querySelector('img:not(.ProseMirror-separator),svg,video,canvas,sup,sub,[data-note-reference],[data-page-number],[data-symbol-font],[data-word-overlap],[data-document-field],[data-drawing-group-anchor]') &&
    parseTabStops(paragraph.dataset.tabStops).every(stop => stop.type === 'left');
  let bodyHeight = ordinary ? wordFallbackLineHeight(style, 'Ag', true) : null;
  let hasText = false;
  while (walker.nextNode()) {
    const node = walker.currentNode as Text;
    if (!node.data || !node.parentElement) continue;
    const run = getComputedStyle(node.parentElement);
    normal = Math.max(normal, singleLine(run, node.data, node));
    if (bodyHeight !== null && node.data.trim()) {
      hasText = true;
      const sample = node.data.replace(/[\t\r\n\u2028\u2029]/g, ' ');
      if (run.fontFamily !== style.fontFamily || run.fontSize !== style.fontSize ||
          run.fontWeight !== style.fontWeight || run.fontStyle !== style.fontStyle ||
          wordFallbackLineHeight(run, sample, true) !== bodyHeight) bodyHeight = null;
    }
  }
  if (hasText && bodyHeight !== null) normal = Math.max(normal, bodyHeight);
  return rule === 'auto' ? normal * value : Math.max(normal, value * (source.endsWith('pt') ? 4 / 3 : 1));
}
/** A Word list item is a numbered paragraph, not another font/line strut.
 * Chromium otherwise adds the editor's larger default marker box above a
 * smaller first paragraph. Keep these inherited marker metrics view-only. */
function listLineProperties(paragraph: HTMLElement, height?: number | null) {
  const css = getComputedStyle(paragraph);
  return { 'font-family': css.fontFamily, 'font-size': css.fontSize, 'font-weight': css.fontWeight,
    'font-style': css.fontStyle, 'line-height': height == null ? css.lineHeight : `${height}px` };
}
const symbolStrutFlag = 'data-word-symbol-body-strut';
const symbolStrutReady = 'data-word-symbol-body-strut-ready';
const symbolStrutVar = '--word-symbol-body-strut';
const fallbackSpaceProperty = '--word-fallback-space-adjustment';

export function measuredBodySpaceAdjustment(paragraph: HTMLElement) {
  if (paragraph.querySelector('br:not(.ProseMirror-trailingBreak),[data-docx-tab],[data-symbol-font],[data-word-overlap],sup,sub')) return null;
  const css = getComputedStyle(paragraph);
  if (css.direction !== 'ltr' || css.writingMode !== 'horizontal-tb' ||
      css.letterSpacing !== 'normal' || paragraph.style.wordSpacing ||
      (css.wordSpacing !== '0px' && css.wordSpacing !== 'normal' && !paragraph.style.getPropertyValue(fallbackSpaceProperty))) return null;
  const delta = wordFallbackSpaceAdjustment(paragraph, css);
  if (delta === null) return null;
  const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT);
  let hasSpace = false;
  while (walker.nextNode()) {
    const node = walker.currentNode as Text;
    if (!node.data || !node.parentElement) continue;
    const run = getComputedStyle(node.parentElement);
    if (!/^[\x20-\x7e]+$/.test(node.data) || run.fontFamily !== css.fontFamily || run.fontSize !== css.fontSize ||
        run.fontWeight !== css.fontWeight || run.fontStyle !== css.fontStyle || run.letterSpacing !== 'normal' ||
        run.wordSpacing !== css.wordSpacing) return null;
    hasSpace ||= node.data.includes(' ');
  }
  return hasSpace ? delta : null;
}
function isMeasuredSymbolItem(item: HTMLElement, paragraph: HTMLElement) {
  return item.dataset.listFont?.toLowerCase() === 'symbol' && item.dataset.listFormat === 'bullet' &&
    item.dataset.listLabel === '\uf0b7' && item.dataset.listGlyph === '•' &&
    paragraph.matches('p[data-word-line-rule="auto"]:not([data-hwp-body-line],[data-hwp-cell-line])') &&
    !item.closest('li li,table,blockquote,.document-furniture,[data-document-note],[data-document-textbox],[data-textbox-content],[data-column-count]') &&
    !paragraph.querySelector('img:not(.ProseMirror-separator),svg,video,canvas,sup,sub,[data-note-reference],[data-page-number],[data-flow-widget],[data-symbol-font],[data-word-overlap],[data-document-field],[data-drawing-group-anchor]');
}
function measuredSymbolLineHeight(paragraph: HTMLElement) {
  const css = getComputedStyle(paragraph);
  if (css.direction !== 'ltr' || css.writingMode !== 'horizontal-tb') return;
  const profile = wordFallbackLineHeight(css, 'Ag', true);
  if (profile === null) return;
  const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT);
  let hasText = false;
  while (walker.nextNode()) {
    const node = walker.currentNode as Text;
    if (!node.data.trim() || !node.parentElement) continue;
    hasText = true;
    const run = getComputedStyle(node.parentElement);
    if (run.fontFamily !== css.fontFamily || run.fontSize !== css.fontSize ||
      run.fontWeight !== css.fontWeight || run.fontStyle !== css.fontStyle ||
      wordFallbackLineHeight(run, node.data.replace(/[\t\r\n\u2028\u2029]/g, ' '), true) !== profile) return;
  }
  return hasText ? wordLineHeight(paragraph) : null;
}
/** Classify from the complete static tree. The page fitter can temporarily
 * truncate a multi-paragraph item, which must not make it eligible mid-fit. */
export function markMeasuredWordSymbolItems(root: HTMLElement) {
  for (const item of root.querySelectorAll<HTMLElement>('li[data-list-id]')) {
    const paragraph = item.firstElementChild;
    if (item.children.length === 1 && paragraph instanceof HTMLElement && isMeasuredSymbolItem(item, paragraph))
      item.setAttribute(symbolStrutFlag, '');
  }
}
export function layoutWordListLines(root: HTMLElement) {
  for (const item of root.querySelectorAll<HTMLElement>('li[data-list-id]')) {
    const paragraph = item.firstElementChild;
    if (!(paragraph instanceof HTMLElement) || !paragraph.matches('p,h1,h2,h3,h4,h5,h6') || !paragraph.dataset.wordLineRule) continue;
    const legacyMarker = unresolvedWordListMarker(item.dataset.listLabel || item.dataset.listPattern);
    const indents = legacyMarker ? nativeWordListIndents(item, paragraph) : null;
    if (indents) {
      paragraph.dataset.wordNativeListIndent = 'true';
      paragraph.style.setProperty('--word-native-list-left', `${indents.left}px`);
      paragraph.style.setProperty('--word-native-list-first', `${indents.first}px`);
    } else if (paragraph.hasAttribute('data-word-native-list-indent')) {
      delete paragraph.dataset.wordNativeListIndent;
      paragraph.style.removeProperty('--word-native-list-left');
      paragraph.style.removeProperty('--word-native-list-first');
    }
    if (legacyMarker) {
      // fits() sees each complete item before probing fragments. Cache its
      // decision (including ineligibility) so clones cannot change it just
      // because a disqualifying run moved onto another page.
      if (item.hasAttribute(symbolStrutFlag) && !item.hasAttribute(symbolStrutReady)) {
        const height = measuredSymbolLineHeight(paragraph);
        if (height != null) item.style.setProperty(symbolStrutVar, `${height}px`);
        item.setAttribute(symbolStrutReady, height == null ? 'no' : 'yes');
      }
      continue;
    }
    for (const [property, value] of Object.entries(listLineProperties(paragraph)))
      if (item.style.getPropertyValue(property) !== value) item.style.setProperty(property, value);
    const original = paragraph.dataset.wordListHanging ?? paragraph.style.textIndent;
    const geometry = wordListGeometry(item, paragraph, original);
    if (geometry) {
      paragraph.dataset.wordListHanging = original;
      item.dataset.wordListLayout = 'true'; paragraph.dataset.wordListParagraph = 'true';
      const properties = document.createElement('span').style; properties.cssText = geometry;
      for (const property of [...properties]) if (item.style.getPropertyValue(property) !== properties.getPropertyValue(property))
        item.style.setProperty(property, properties.getPropertyValue(property));
    } else {
      delete item.dataset.wordListLayout; delete paragraph.dataset.wordListParagraph;
      for (const property of ['margin-left', '--word-list-marker-left', '--word-list-marker-top']) item.style.removeProperty(property);
    }
  }
}
export function layoutWordLines(root: HTMLElement) {
  for (const p of root.querySelectorAll<HTMLElement>('[data-word-body-strut="true"]:not([data-print-word-space],[data-print-paragraph-fragment])')) {
    const delta = measuredBodySpaceAdjustment(p);
    if (delta !== null) p.style.setProperty(fallbackSpaceProperty, `${delta}px`);
    // Freeze eligibility on the full paragraph before the print fitter splits it.
    p.setAttribute('data-print-word-space', '');
  }
  for (const p of root.querySelectorAll<HTMLElement>('[data-word-line-rule]')) {
    // Do not rerun this new, full-paragraph policy on split fragments.
    // Empty means the full paragraph matched no measured body profile.
    let frozen = p.getAttribute('data-print-word-body-line');
    if (p.dataset.wordBodyStrut === 'true' && !p.hasAttribute('data-print-paragraph-fragment') && frozen === null) {
      const normal = wordLineHeight(p);
      const measured = wordLineHeight(p, true);
      frozen = measured !== null && measured !== normal ? String(measured) : '';
      p.setAttribute('data-print-word-body-line', frozen);
    }
    const height = frozen ? Number(frozen) : wordLineHeight(p);
    if (height !== null) p.style.lineHeight = `${height}px`;
  }
  layoutWordListLines(root);
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
        const wordLayout = documentMetadata.getState(view.state);
        const ordinaryFlow = (wordLayout?.wordHangingSpaces || wordLayout?.wordBodyStrut) && !wordFlowHasWrappingAnchors(view.dom);
        const hangWordSpaces = ordinaryFlow && wordLayout?.wordHangingSpaces;
        const useWordStrut = ordinaryFlow && wordLayout?.wordBodyStrut;
        view.state.doc.descendants((node, pos) => {
          if (useWordStrut && node.type.name === 'paragraph') {
            const p = view.nodeDOM(pos);
            if (p instanceof HTMLElement && isOrdinaryWordBodyParagraph(p, view.dom)) {
              const delta = measuredBodySpaceAdjustment(p);
              if (delta !== null) {
                const style = `${fallbackSpaceProperty}:${delta}px`;
                decorations.push(Decoration.node(pos, pos + node.nodeSize, { style }));
                values.push(`${pos}:${style}`);
              }
            }
          }
          if (wordLayout?.wordHangingSpaces && node.type.name === 'paragraph') {
            const paragraph = view.nodeDOM(pos);
            if (paragraph instanceof HTMLElement &&
                ((hangWordSpaces && (isOrdinaryWordBodyParagraph(paragraph, view.dom) || isWordAlignedBulletParagraph(paragraph))) ||
                 isWordGroupedTextBoxParagraph(paragraph))) {
              decorations.push(Decoration.node(pos, pos + node.nodeSize, { 'data-word-hanging-spaces': 'true' }));
              values.push(`${pos}:word-hanging-spaces`);
            }
          }
          if (node.type.name === 'listItem' && node.attrs.wordListId && node.firstChild?.attrs.wordLineRule) {
            const paragraph = view.nodeDOM(pos + 1);
            if (paragraph instanceof HTMLElement && unresolvedWordListMarker(node.attrs.wordListLabel || node.attrs.wordListPattern)) {
              const item = paragraph.parentElement!;
              const height = useWordStrut && node.childCount === 1 && isMeasuredSymbolItem(item, paragraph)
                ? measuredSymbolLineHeight(paragraph) : null;
              if (height != null) {
                const style = `${symbolStrutVar}:${height}px`;
                decorations.push(Decoration.node(pos, pos + node.nodeSize, { [symbolStrutReady]: 'yes', style }));
                values.push(`${pos}:${style}`);
              }
              const indents = nativeWordListIndents(item, paragraph);
              if (indents) {
                const style = `--word-native-list-left:${indents.left}px;--word-native-list-first:${indents.first}px`;
                decorations.push(Decoration.node(pos + 1, pos + 1 + node.firstChild.nodeSize, { 'data-word-native-list-indent': 'true', style }));
                values.push(`${pos + 1}:native-list:${style}`);
              }
            } else if (paragraph instanceof HTMLElement) {
              let style = Object.entries(listLineProperties(paragraph, wordLineHeight(paragraph))).map(([p, v]) => `${p}:${v}`).join(';');
              const geometry = wordListGeometry(paragraph.parentElement!, paragraph, node.firstChild.attrs.firstLineIndent);
              if (geometry) {
                style += ';' + geometry;
                decorations.push(Decoration.node(pos + 1, pos + 1 + node.firstChild.nodeSize, { 'data-word-list-paragraph': 'true' }));
              }
              decorations.push(Decoration.node(pos, pos + node.nodeSize, { style, ...(geometry ? { 'data-word-list-layout': 'true' } : {}) }));
              values.push(`${pos}:${style}`);
            }
          }
          if (!node.attrs.wordLineRule) return;
          const dom = view.nodeDOM(pos);
          if (!(dom instanceof HTMLElement)) return;
          const height = wordLineHeight(dom, useWordStrut && isOrdinaryWordBodyParagraph(dom, view.dom));
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
      const fontsChanged = () => { metrics.clear(); nodeMetrics = new WeakMap(); signature = '\0'; schedule(); };
      document.fonts.addEventListener('loadingdone', fontsChanged);
      document.addEventListener(wordFontMetricsChanged, fontsChanged); schedule();
      return { update: schedule, destroy() { cancelAnimationFrame(frame); document.fonts.removeEventListener('loadingdone', fontsChanged); document.removeEventListener(wordFontMetricsChanged, fontsChanged); } };
    },
  })],
});
