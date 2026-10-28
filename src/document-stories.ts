import { Node, mergeAttributes } from '@tiptap/core';
import { SectionNodeView } from './section-node-view';
import { Plugin } from '@tiptap/pm/state';
import { documentScale } from './document-scale';

const dataAttribute = (attribute: string, fallback = '') => ({ default: fallback,
  parseHTML: (el: HTMLElement) => el.getAttribute(attribute) || fallback,
  renderHTML: (attrs: Record<string, unknown>) => ({ [attribute]: attrs[attribute === 'data-note-kind' ? 'kind' : attribute === 'data-note-label' ? 'label' : 'id'] }),
});
export const NoteReference = Node.create({
  name: 'noteReference', group: 'inline', inline: true, atom: true,
  addAttributes: () => ({ id: dataAttribute('data-note-reference'), kind: dataAttribute('data-note-kind', 'footnote'), label: dataAttribute('data-note-label') }),
  parseHTML: () => [{ tag: 'span[data-note-reference]' }],
  renderHTML: ({ node, HTMLAttributes }) => ['span', { ...HTMLAttributes, class: 'note-reference', title: `${node.attrs.kind === 'endnote' ? '미주' : '각주'} ${node.attrs.label}` }, node.attrs.label],
  addProseMirrorPlugins() {
    return [new Plugin({ props: { handleClick(view, _position, event) {
      const ref = (event.target as Element)?.closest<HTMLElement>('[data-note-reference]'); if (!ref) return false;
      const note = [...view.dom.querySelectorAll<HTMLElement>('[data-document-note]')].find(el => el.dataset.documentNote === ref.dataset.noteReference);
      if (!note) return false; note.scrollIntoView({ block: 'center' }); event.preventDefault(); return true;
    } } })];
  },
});
export const DocumentNote = Node.create({
  name: 'documentNote', group: 'block', content: 'block+', defining: true,
  addAttributes: () => ({ id: dataAttribute('data-document-note'), kind: dataAttribute('data-note-kind', 'footnote'), label: dataAttribute('data-note-label') }),
  parseHTML: () => [{ tag: 'div[data-document-note]' }],
  renderHTML: ({ HTMLAttributes }) => ['div', { ...HTMLAttributes, class: 'document-note' }, 0],
});
export const DocumentField = Node.create({
  name: 'documentField', group: 'inline', inline: true, atom: true,
  addAttributes: () => ({ kind: { default: 'PAGE', parseHTML: el => el.getAttribute('data-page-number'), renderHTML: a => ({ 'data-page-number': a.kind === 'NUMPAGES' ? 'NUMPAGES' : 'PAGE' }) } }),
  parseHTML: () => [{ tag: 'span[data-page-number]' }],
  renderHTML: ({ HTMLAttributes }) => ['span', { ...HTMLAttributes, class: 'document-page-field' }, '1'],
  addNodeView: () => ({ node, editor }) => {
    const dom = document.createElement('span'); dom.className = 'document-page-field'; dom.contentEditable = 'false';
    let current = node, frame = 0;
    const update = () => {
      cancelAnimationFrame(frame); frame = requestAnimationFrame(() => {
        if (!dom.isConnected) return;
        const stack = dom.closest<HTMLElement>('.document-stack'), section = dom.closest<HTMLElement>('.document-section-content');
        const root = section?.querySelector<HTMLElement>('.section-editable') ?? editor.view.dom;
        const stride = Number(section?.dataset.pageHeight ?? stack?.dataset.pageHeight) + 24;
        const before = section ? [...editor.view.dom.querySelectorAll<HTMLElement>('.document-section-content')].slice(0,
          [...editor.view.dom.querySelectorAll('.document-section-content')].indexOf(section)).reduce((sum, s) => sum + (Number(s.dataset.pages) || 1), 0) : 0;
        const y = (dom.getBoundingClientRect().top - root.getBoundingClientRect().top) / documentScale(root);
        // Glyph ink can extend above the first line's body margin. Determine
        // its physical paper, not whether its top clears the body inset.
        const page = Number.isFinite(stride) && stride > 24 ? before + Math.max(0, Math.floor((y + .5) / stride)) + 1 : 1;
        dom.dataset.pageNumber = current.attrs.kind;
        dom.textContent = String(current.attrs.kind === 'NUMPAGES' ? Number(stack?.dataset.pages) || 1 : page);
      });
    };
    const stack = editor.view.dom.closest('.document-stack'); stack?.addEventListener('pagecount', update);
    update();
    return { dom, update(next) { if (next.type !== current.type) return false; current = next; update(); return true; },
      ignoreMutation: () => true, destroy() { cancelAnimationFrame(frame); stack?.removeEventListener('pagecount', update); } };
  },
});
export const ColumnGroup = Node.create({
  name: 'columnGroup', group: 'block', content: 'block+', defining: true,
  addAttributes: () => ({ count: { default: 2, parseHTML: el => Math.max(1, Math.min(12, Number(el.getAttribute('data-column-count')) || 2)) },
    space: { default: 48, parseHTML: el => Number(el.getAttribute('data-column-space')) || 0 },
    balance: { default: true, parseHTML: el => el.getAttribute('data-column-balance') !== 'false' } }),
  parseHTML: () => [{ tag: 'div[data-column-count]' }],
  renderHTML: ({ node, HTMLAttributes }) => ['div', mergeAttributes(HTMLAttributes, { 'data-column-count': node.attrs.count, 'data-column-space': node.attrs.space,
    'data-column-balance': String(node.attrs.balance), class: 'document-columns', style: `column-count:${node.attrs.count === 1 ? 'auto' : Math.max(1, Math.min(12, node.attrs.count))};column-gap:${Math.max(0, Math.min(1000, node.attrs.space))}px;column-fill:${node.attrs.balance ? 'balance' : 'auto'}` }), 0],
  addNodeView: () => ({ node }) => {
    const dom = document.createElement('div');
    const update = (next: typeof node) => {
      if (next.type.name !== 'columnGroup') return false;
      dom.className = 'document-columns'; dom.dataset.columnCount = String(next.attrs.count); dom.dataset.columnSpace = String(next.attrs.space); dom.dataset.columnBalance = String(next.attrs.balance);
      // column-count:1 still creates a multicol formatting context. A forced
      // break would create an off-screen second column instead of normal
      // page flow, even when pagination inserts the correct vertical gap.
      dom.style.columnCount = next.attrs.count === 1 ? 'auto' : String(next.attrs.count); dom.style.columnGap = `${next.attrs.space}px`; dom.style.columnFill = next.attrs.balance ? 'balance' : 'auto'; return true;
    };
    update(node); return { dom, contentDOM: dom, update, ignoreMutation: mutation => mutation.type === 'attributes' && mutation.target === dom };
  },
});
export const SectionContent = Node.create({
  name: 'sectionContent', group: 'block', content: 'block+', defining: true,
  addAttributes: () => ({ page: { default: '', parseHTML: el => el.getAttribute('data-section-content'), renderHTML: a => ({ 'data-section-content': a.page }) } }),
  parseHTML: () => [{ tag: 'div[data-section-content]' }],
  renderHTML: ({ HTMLAttributes }) => ['div', { ...HTMLAttributes, class: 'document-section-content' }, 0],
  addNodeView: () => ({ node }) => new SectionNodeView(node),
});
