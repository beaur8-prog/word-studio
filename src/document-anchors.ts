import { Node } from '@tiptap/core';
import { Plugin } from '@tiptap/pm/state';

export const DocumentBookmark = Node.create({
  name: 'documentBookmark', group: 'inline', inline: true, atom: true, selectable: false,
  addAttributes: () => ({ name: { default: '', parseHTML: el => el.getAttribute('data-docx-bookmark'), renderHTML: a => ({ 'data-docx-bookmark': a.name, id: a.name }) } }),
  parseHTML: () => [{ tag: 'span[data-docx-bookmark]' }],
  renderHTML: ({ HTMLAttributes }) => ['span', { ...HTMLAttributes, class: 'document-bookmark' }],
  addProseMirrorPlugins() {
    return [new Plugin({ props: { handleClick: (view, _position, event) => {
      const anchor = (event.target as Element)?.closest('a[href^="#"]');
      if (!anchor || (view.editable && !event.ctrlKey && !event.metaKey)) return false;
      let id = anchor.getAttribute('href')!.slice(1);
      try { id = decodeURIComponent(id); } catch { /* A literal bookmark can contain percent. */ }
      const target = [...view.dom.querySelectorAll<HTMLElement>('[data-docx-bookmark]')].find(el => el.dataset.docxBookmark === id);
      if (!target) return false;
      target.scrollIntoView({ block: 'center' }); event.preventDefault(); return true;
    } } })];
  },
});
