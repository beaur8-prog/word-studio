import { Extension } from '@tiptap/core';
import { Plugin } from '@tiptap/pm/state';
import { documentScale } from './document-scale';
import { tableAnchor, tableWrapFits } from './table-anchor-data';
export { tableAnchor, tableWrapFits, paragraphTableNamespace, type TableAnchor } from './table-anchor-data';

/** Pagination widgets are view furniture, not a new anchor paragraph. */
export function tableFollowingParagraph(target: HTMLElement): HTMLElement | undefined {
  let next = target.nextElementSibling;
  while (next?.classList.contains('pagination-gap')) next = next.nextElementSibling;
  return next?.matches('p,h1,h2,h3,h4,h5,h6') ? next as HTMLElement : undefined;
}

export function layoutTableAnchors(root: HTMLElement) {
  // Restore natural geometry before resolving again; classes and variables
  // also survive print fragments cloned while searching for a page break.
  for (const frame of root.querySelectorAll<HTMLElement>('.table-overlay-frame')) {
    frame.classList.remove('table-overlay-frame');
    frame.style.removeProperty('--table-overlay-left'); frame.style.removeProperty('--table-overlay-top');
  }
  for (const frame of root.querySelectorAll<HTMLElement>('.table-wrap-frame')) {
    frame.classList.remove('table-wrap-frame');
    for (const name of ['--table-wrap-left', '--table-wrap-right', '--table-wrap-bottom', '--table-wrap-top']) frame.style.removeProperty(name);
  }
  for (const table of root.querySelectorAll('.table-overlay-active')) table.classList.remove('table-overlay-active');
  for (const table of root.querySelectorAll<HTMLElement>('table[data-table-anchor]')) {
    const anchor = tableAnchor(table.dataset.tableAnchor); if (!anchor) continue;
    let target = table.parentElement?.classList.contains('tableWrapper') || table.parentElement?.hasAttribute('data-table-overlay-wrapper') ? table.parentElement! : table;
    const sheet = target.closest<HTMLElement>('.print-page,.document-section-content,.document-stack'); if (!sheet) continue;
    const scale = documentScale(target), paper = sheet.getBoundingClientRect();
    if (anchor.wrap) {
      const parent = target.parentElement!, padding = getComputedStyle(parent);
      const width = (parent.getBoundingClientRect().width / scale) - (parseFloat(padding.paddingLeft) || 0) - (parseFloat(padding.paddingRight) || 0);
      // An enlarged label must become ordinary flow rather than paint beyond
      // the paper or trap the following text in an unusably narrow strip.
      const following = tableFollowingParagraph(target);
      if (!following || !tableWrapFits(anchor, table.getBoundingClientRect().width / scale, width) || table.getBoundingClientRect().height / scale > paper.height / scale) continue;
      target.classList.add('table-wrap-frame');
      target.style.setProperty('--table-wrap-top', `${Math.max(0, parseFloat(getComputedStyle(following).marginTop) || 0)}px`);
      target.style.setProperty('--table-wrap-left', `${anchor.x}px`);
      target.style.setProperty('--table-wrap-right', `${anchor.right}px`);
      target.style.setProperty('--table-wrap-bottom', `${anchor.bottom}px`);
      continue;
    }
    if (anchor.overlay) {
      const height = Number(sheet.dataset.pageHeight) || paper.height / scale;
      const top = sheet.classList.contains('print-page') ? (root.getBoundingClientRect().top - paper.top) / scale : Number(sheet.dataset.margin) || 0;
      const bottom = sheet.classList.contains('print-page') ? (root.getBoundingClientRect().bottom - paper.top) / scale
        : height - (Number(sheet.dataset.bottomMargin) || Number(sheet.dataset.margin) || 0);
      const tableHeight = table.getBoundingClientRect().height / scale;
      const flowHeight = [...table.getClientRects()].reduce((sum, rect) => sum + rect.height / scale, 0);
      // A table enlarged beyond one sheet must remain ordinary, splittable
      // flow rather than disappear below an absolute frame.
      if (!(scale > 0) || Math.max(tableHeight, flowHeight) > bottom - top + .5) continue;
      if (target === table) {
        const frame = document.createElement('div'); frame.dataset.tableOverlayWrapper = 'true';
        table.replaceWith(frame); frame.append(table); target = frame;
      }
      target.classList.add('table-overlay-frame'); table.classList.add('table-overlay-active');
      const rect = target.getBoundingClientRect(), naturalTop = (rect.top - paper.top) / scale;
      const stride = height + 24;
      const page = sheet.classList.contains('print-page') ? 0 : Math.max(0, Math.floor((naturalTop - top + .5) / stride));
      target.style.setProperty('--table-overlay-left', `${anchor.x - (rect.left - paper.left) / scale}px`);
      target.style.setProperty('--table-overlay-top', `${page * stride + (anchor.overlay === 'page-bottom' ? bottom : height) - anchor.y - tableHeight - naturalTop}px`);
      continue;
    }
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
