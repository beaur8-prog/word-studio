import type { Node as PMNode } from '@tiptap/pm/model';
import { defaultPage, pageMargins, pageFurniture, paperDimensions, type PageSettings } from './types';
import { sanitizeImportedHtml } from './document-html';
import { layoutFurnitureImages } from './furniture-layout';

export function sectionSettings(raw: unknown): PageSettings {
  try {
    const page = JSON.parse(String(raw));
    if (!page || !['A4','Letter','Custom'].includes(page.size) || !['portrait','landscape'].includes(page.orientation) || !Number.isFinite(page.margin)) return defaultPage;
    if (page.size === 'Custom' && (![page.width, page.height].every(n => Number.isFinite(n) && n >= 25 && n <= 1320))) return defaultPage;
    return page;
  } catch { return defaultPage; }
}
export const pageDimensions = (page: PageSettings) => paperDimensions(page).map(n => Math.round(n * 96 / 25.4 * 100) / 100);

/** Give each imported section its own editable text width and page furniture.
 * Backgrounds are outside contentDOM, so they never become document content. */
export class SectionNodeView {
  dom = document.createElement('div');
  contentDOM = document.createElement('div');
  private backgrounds = document.createElement('div');
  private observer: ResizeObserver;
  private frame = 0;
  private signature = '';
  private stack: HTMLElement | null = null;
  private pageCount = () => this.schedule();
  constructor(private node: PMNode) {
    this.dom.className = 'document-section-content'; this.contentDOM.className = 'section-editable';
    this.backgrounds.className = 'section-papers'; this.backgrounds.contentEditable = 'false'; this.backgrounds.setAttribute('aria-hidden', 'true');
    this.dom.append(this.backgrounds, this.contentDOM);
    this.observer = new ResizeObserver(() => this.schedule()); this.observer.observe(this.contentDOM);
    this.schedule();
  }
  private schedule() { if (this.frame) return; this.frame = requestAnimationFrame(() => { this.frame = 0; this.layout(); }); }
  private layout() {
    const page = sectionSettings(this.node.attrs.page), [width, height] = pageDimensions(page), margins = pageMargins(page);
    const px = 96 / 25.4, count = Math.max(1, Number(this.dom.dataset.pages) || 1);
    const stack = this.dom.closest<HTMLElement>('.document-stack');
    if (stack !== this.stack) {
      this.stack?.removeEventListener('pagecount', this.pageCount);
      this.stack = stack; stack?.addEventListener('pagecount', this.pageCount);
    }
    const sections = [...(stack?.querySelectorAll<HTMLElement>('.document-section-content') ?? [])];
    const before = sections.slice(0, sections.indexOf(this.dom)).reduce((sum, section) => sum + (Number(section.dataset.pages) || 1), 0);
    const total = Number(stack?.dataset.pages) || count;
    this.dom.style.width = `${width}px`; this.dom.dataset.sectionContent = this.node.attrs.page;
    this.dom.dataset.pageHeight = String(height);
    this.dom.dataset.hideEmptyLines = String(page.hideEmptyLines === true);
    this.dom.dataset.suppressBottomSpacing = String(page.suppressBottomSpacing !== false);
    const signature = `${this.node.attrs.page}:${count}:${before}:${total}`;
    if (signature !== this.signature) {
      this.signature = signature;
      for (const old of this.backgrounds.querySelectorAll('.section-header,.section-footer')) this.observer.unobserve(old);
      this.backgrounds.replaceChildren();
      for (let i = 0; i < count; i++) {
        const paper = document.createElement('div'); paper.className = 'section-paper';
        Object.assign(paper.style, { top: `${i * (height + 24)}px`, width: `${width}px`, height: `${height}px` });
        const number = before + i + 1;
        const furniture = pageFurniture(page, number, i === 0), header = document.createElement('div'), footer = document.createElement('div');
        header.className = 'section-header document-furniture'; footer.className = 'section-footer document-furniture';
        if (furniture.headerHtml) header.innerHTML = sanitizeImportedHtml(furniture.headerHtml).html; else header.textContent = furniture.header;
        if (furniture.footerHtml) footer.innerHTML = sanitizeImportedHtml(furniture.footerHtml).html; else footer.textContent = furniture.footer;
        if (furniture.pageNumbers && !/data-page-number/.test((furniture.headerHtml || '') + (furniture.footerHtml || ''))) {
          const field = document.createElement('span'); field.dataset.pageNumber = 'PAGE'; footer.append(field);
        }
        for (const element of [header, footer]) Object.assign(element.style, { left: `${margins.left * px}px`, right: `${margins.right * px}px` });
        header.style.top = `${(page.headerDistance ?? Math.max(0, margins.top / 2 - 3)) * px}px`;
        footer.style.bottom = `${(page.footerDistance ?? Math.max(0, margins.bottom / 2 - 3)) * px}px`;
        paper.append(header, footer); this.backgrounds.append(paper);
        layoutFurnitureImages(paper);
        this.observer.observe(header); this.observer.observe(footer);
        for (const field of paper.querySelectorAll<HTMLElement>('[data-page-number]')) field.textContent = String(field.dataset.pageNumber === 'NUMPAGES' ? total : number);
      }
    }
    const occupied = (selector: string, distance: number) => Math.max(0, ...[...this.backgrounds.querySelectorAll<HTMLElement>(selector)].map(el =>
      el.textContent?.trim() || el.querySelector('img,table') ? distance * px + el.offsetHeight + 2 : 0));
    const top = Math.max(margins.top * px, occupied('.section-header', page.headerDistance ?? Math.max(0, margins.top / 2 - 3)));
    const bottom = Math.max(margins.bottom * px, occupied('.section-footer', page.footerDistance ?? Math.max(0, margins.bottom / 2 - 3)));
    Object.assign(this.contentDOM.style, { padding: `${top}px ${margins.right * px}px ${bottom}px ${margins.left * px}px`, minHeight: `${count * (height + 24) - 24}px` });
    this.dom.dataset.margin = String(top); this.dom.dataset.bottomMargin = String(bottom);
    this.dom.dataset.leftMargin = String(margins.left * px); this.dom.dataset.rightMargin = String(margins.right * px);
  }
  update(node: PMNode) { if (node.type !== this.node.type) return false; this.node = node; this.schedule(); return true; }
  ignoreMutation(mutation: { target: globalThis.Node; type?: string }) { return mutation.target === this.dom || (mutation.type === 'attributes' && mutation.target === this.contentDOM) || !this.contentDOM.contains(mutation.target); }
  destroy() { cancelAnimationFrame(this.frame); this.stack?.removeEventListener('pagecount', this.pageCount); this.observer.disconnect(); }
}
