import { Extension } from '@tiptap/core';
import { Plugin } from '@tiptap/pm/state';
import { documentScale } from './document-scale';

export type TableAnchor = { x: number; y: number; bottom: number; reference?: 'paragraph'; horizontal?: 'margin' };
export const paragraphTableNamespace = 'urn:wordstudio:paragraph-table:1';

/** Full-width tables wrapping above/below the body. Other
 * floating table wraps need their own exclusion layout and are not inferred. */
export function tableAnchor(value: unknown): TableAnchor | undefined {
  try {
    const a = typeof value === 'string' ? JSON.parse(value) : value;
    if (a && [a.x, a.y, a.bottom].every(Number.isFinite) && a.x >= 0 && a.x <= 5000 && a.y >= 0 && a.y <= 5000 && a.bottom >= 0 && a.bottom <= 1000 &&
      (a.reference === undefined || a.reference === 'paragraph') && (a.horizontal === undefined || a.horizontal === 'margin'))
      return { x: a.x, y: a.y, bottom: a.bottom, ...(a.reference ? { reference: a.reference } : {}), ...(a.horizontal ? { horizontal: a.horizontal } : {}) };
  } catch { /* Unknown floating settings remain ordinary flow. */ }
}

export function layoutTableAnchors(root: HTMLElement) {
  for (const table of root.querySelectorAll<HTMLElement>('table[data-table-anchor]')) {
    const anchor = tableAnchor(table.dataset.tableAnchor); if (!anchor) continue;
    const target = table.parentElement?.classList.contains('tableWrapper') ? table.parentElement : table;
    const sheet = target.closest<HTMLElement>('.print-page,.document-section-content,.document-stack'); if (!sheet) continue;
    const scale = documentScale(target), paper = sheet.getBoundingClientRect();
    const oldTop = parseFloat(target.style.marginTop) || 0, oldLeft = parseFloat(target.style.marginLeft) || 0;
    const rect = target.getBoundingClientRect(), naturalTop = (rect.top - paper.top) / scale - oldTop;
    const stride = (Number(sheet.dataset.pageHeight) || sheet.clientHeight) + 24;
    const page = sheet.classList.contains('print-page') ? 0 : Math.max(0, Math.floor((naturalTop + .5) / stride));
    target.style.marginTop = `${anchor.reference === 'paragraph' ? anchor.y : page * stride + anchor.y - naturalTop}px`;
    const parent = target.parentElement!, parentStyle = getComputedStyle(parent);
    const horizontalOrigin = anchor.horizontal === 'margin'
      ? (parent.getBoundingClientRect().left - paper.left) / scale + (parseFloat(parentStyle.paddingLeft) || 0) + (parseFloat(parentStyle.borderLeftWidth) || 0) : 0;
    target.style.marginLeft = `${horizontalOrigin + anchor.x - (rect.left - paper.left) / scale + oldLeft}px`;
    const bodyTop = sheet.classList.contains('print-page') ? (root.getBoundingClientRect().top - paper.top) / scale : Number(sheet.dataset.margin) || 0;
    target.style.marginBottom = `${anchor.reference === 'paragraph' ? anchor.bottom : Math.max(anchor.bottom, bodyTop - anchor.y - table.getBoundingClientRect().height / scale)}px`;
  }
}

export const TableAnchorLayout = Extension.create({
  name: 'tableAnchorLayout',
  addProseMirrorPlugins: () => [new Plugin({ view(view) {
    let frame = 0;
    const update = () => { cancelAnimationFrame(frame); frame = requestAnimationFrame(() => { if (!view.composing) layoutTableAnchors(view.dom); }); };
    const observer = new ResizeObserver(update); observer.observe(view.dom); update();
    document.fonts.addEventListener('loadingdone', update);
    return { update, destroy() { cancelAnimationFrame(frame); observer.disconnect(); document.fonts.removeEventListener('loadingdone', update); } };
  } })],
});
