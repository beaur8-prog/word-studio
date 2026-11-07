import { documentScale } from './document-scale';

/** Paper backgrounds must cover the committed editor content, even when a
 * table reflow makes the pagination event lag one sheet behind. Never use
 * the editor's min-height: it already depends on the previous page count. */
export function visibleEditorPages(root: HTMLElement, pageHeight: number, paginated: number) {
  const requested = Math.max(1, paginated);
  // Sections have their own paper sizes and backgrounds. A single global
  // stride cannot measure their physical page count.
  if (root.querySelector('[data-section-content]')) return requested;
  const range = document.createRange(); range.selectNodeContents(root);
  const bounds = range.getBoundingClientRect();
  if (!bounds.height) return requested;
  const extent = (bounds.bottom - root.getBoundingClientRect().top) / documentScale(root);
  return Math.max(requested, Math.ceil((extent - .5) / (pageHeight + 24)));
}
