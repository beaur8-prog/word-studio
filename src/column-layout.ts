import { Extension } from '@tiptap/core';
import { Plugin } from '@tiptap/pm/state';
import { layoutTabs } from './document-tabs';
import { layoutWordLines } from './word-line-metrics';
import { documentScale } from './document-scale';

/** A column row is a paper's text area. Its gap includes the bottom margin,
 * the paper gap and the following top margin. A view-only spanning prefix
 * reserves text already above a continuous section on the first sheet. */
export function layoutColumns(root: HTMLElement, bottom?: number) {
  for (const group of root.querySelectorAll<HTMLElement>('[data-column-count]')) {
    if (Number(group.dataset.columnCount) === 1) {
      for (const property of ['column-height', 'column-wrap', 'row-gap', '--column-prefix', 'margin-top', 'height']) group.style.removeProperty(property);
      continue;
    }
    const sheet = group.closest<HTMLElement>('[data-page-height]');
    const scale = documentScale(group);
    const height = Number(sheet?.dataset.pageHeight) || 1122.52;
    const top = Number(sheet?.dataset.margin) || 96, marginBottom = Number(sheet?.dataset.bottomMargin) || top;
    const usable = Math.max(1, height - top - marginBottom), stride = height + 24;
    const oldPrefix = parseFloat(group.style.getPropertyValue('--column-prefix')) || 0;
    const origin = sheet?.getBoundingClientRect().top || 0;
    const naturalTop = (group.getBoundingClientRect().top - origin) / scale + oldPrefix;
    let prefix = bottom === undefined ? Math.max(0, ((naturalTop - top) % stride + stride) % stride) : 0;
    // Browser subpixel rounding at the page start must not turn -0.01px
    // into an entire page's spanning prefix.
    if (prefix < .5 || stride - prefix < .5) prefix = 0;
    const cap = bottom === undefined ? usable : Math.max(1, (bottom - group.getBoundingClientRect().top) / scale);
    if (CSS.supports('column-wrap', 'wrap')) {
      group.style.setProperty('column-height', `${cap}px`);
      group.style.setProperty('column-wrap', bottom === undefined ? 'wrap' : 'nowrap');
      group.style.rowGap = `${top + marginBottom + 24}px`;
      group.style.setProperty('--column-prefix', `${prefix}px`);
      group.style.marginTop = prefix ? `${-prefix}px` : '0px';
      group.style.removeProperty('height');
      // column-height preserves fragment geometry even when the outside box is
      // trimmed to the final row's actual content (balanced or unbalanced).
      const rect = group.getBoundingClientRect();
      const extent = Math.max(prefix, ...[...group.children].flatMap(child => [...child.getClientRects()].map(r => (r.bottom - rect.top) / scale)));
      group.style.height = `${Math.max(1, extent)}px`;
    } else {
      const count = Number(group.dataset.columnCount) || 1, gap = Number(group.dataset.columnSpace) || 0;
      const measure = document.createElement('div'); measure.className = 'tiptap column-measure';
      Object.assign(measure.style, { position: 'absolute', visibility: 'hidden', left: '-100000px', top: '0', width: `${(group.clientWidth - gap * (count - 1)) / count}px`, padding: '0', minHeight: '0' });
      measure.innerHTML = group.innerHTML; document.body.append(measure); layoutWordLines(measure); layoutTabs(measure);
      const natural = Math.ceil(measure.getBoundingClientRect().height); measure.remove();
      group.style.height = `${Math.min(natural || 1, cap)}px`;
    }
  }
}

export const ColumnLayout = Extension.create({
  name: 'columnLayout',
  addProseMirrorPlugins() {
    return [new Plugin({ view(view) {
      let frame = 0;
      const calculate = () => { cancelAnimationFrame(frame); frame = requestAnimationFrame(() => { if (!view.composing) layoutColumns(view.dom); }); };
      const observer = new ResizeObserver(calculate); observer.observe(view.dom); calculate();
      document.fonts.addEventListener('loadingdone', calculate);
      return { update: calculate, destroy() { cancelAnimationFrame(frame); observer.disconnect(); document.fonts.removeEventListener('loadingdone', calculate); } };
    } })];
  },
});
