import { Extension } from '@tiptap/core';
import { Plugin } from '@tiptap/pm/state';
import { layoutTabs } from './document-tabs';
import { layoutWordLines } from './word-line-metrics';

/** Word's noColumnBalance fills the first column to the page boundary before
 * starting the next. CSS needs an explicit height to produce that flow. */
export function layoutColumns(root: HTMLElement, bottom?: number) {
  for (const group of root.querySelectorAll<HTMLElement>('[data-column-count]')) {
    if (group.dataset.columnBalance !== 'false') { group.style.removeProperty('height'); continue; }
    const count = Number(group.dataset.columnCount) || 1, gap = Number(group.dataset.columnSpace) || 0;
    const scale = group.getBoundingClientRect().width / group.offsetWidth || 1;
    const width = (group.clientWidth - gap * (count - 1)) / count;
    if (width <= 0) continue;
    const measure = document.createElement('div'); measure.className = 'tiptap column-measure';
    Object.assign(measure.style, { position: 'absolute', visibility: 'hidden', left: '-100000px', top: '0', width: `${width}px`, padding: '0', minHeight: '0' });
    measure.innerHTML = group.innerHTML; measure.querySelectorAll('.pagination-gap').forEach(el => el.remove());
    document.body.append(measure); layoutWordLines(measure); layoutTabs(measure);
    const natural = Math.ceil(measure.getBoundingClientRect().height); measure.remove();
    const sheet = group.closest<HTMLElement>('[data-page-height]');
    const height = Number(sheet?.dataset.pageHeight) || 1122.52;
    const top = Number(sheet?.dataset.margin) || 96, marginBottom = Number(sheet?.dataset.bottomMargin) || top;
    const origin = sheet?.getBoundingClientRect().top || 0;
    const y = (group.getBoundingClientRect().top - origin) / scale;
    const available = bottom === undefined ? height - marginBottom - (y % (height + 24)) : (bottom - group.getBoundingClientRect().top) / scale;
    const next = `${Math.max(1, Math.min(natural || 1, available > 0 ? available : height - top - marginBottom))}px`;
    if (group.style.height !== next) group.style.height = next;
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
