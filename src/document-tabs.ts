import { documentScale } from './document-scale';
import { Extension, Node, type Editor } from '@tiptap/core';
import { Plugin, TextSelection } from '@tiptap/pm/state';

export type DocumentTabStop = { position: number; type: 'left' | 'right' | 'center' | 'decimal'; leader?: 'dot' | 'hyphen' | 'underscore' };
export function parseTabStops(value: unknown): DocumentTabStop[] {
  if (typeof value !== 'string' || value.length > 10000) return [];
  try {
    const list = JSON.parse(value);
    return Array.isArray(list) ? list.filter(s => s && Number.isFinite(s.position) && Math.abs(s.position) <= 10000 &&
      ['left', 'right', 'center', 'decimal'].includes(s.type)).slice(0, 100).sort((a, b) => a.position - b.position) : [];
  } catch { return []; }
}

/** Tabs are measured from the text margin, not replaced with a fixed number of
 * spaces. This is also called on the shared preview/PDF tree after each append. */
export function layoutTabs(root: HTMLElement) {
  for (const paragraph of root.querySelectorAll<HTMLElement>('p,h1,h2,h3,h4,h5,h6')) {
    const tabs = [...paragraph.querySelectorAll<HTMLElement>('[data-docx-tab]')];
    if (!tabs.length) continue;
    const stops = parseTabStops(paragraph.dataset.tabStops);
    const paragraphStyle = getComputedStyle(paragraph);
    const scale = documentScale(paragraph);
    const origin = paragraph.getBoundingClientRect().left - (parseFloat(paragraphStyle.marginLeft) || 0) * scale;
    // Header/footer HTML is sanitized directly rather than passed through the
    // editor's node view. Give those imported spans the same measurable inline
    // box; an ordinary inline span ignores the computed tab width.
    for (const tab of tabs) { tab.classList.add('document-tab'); tab.style.width = '0px'; }
    for (const tab of tabs) {
      const current = (tab.getBoundingClientRect().left - origin) / scale;
      const nextDefault = (Math.floor((current + .5) / 48) + 1) * 48;
      let stop: DocumentTabStop = stops.find(s => s.position > current + .5) ?? { position: nextDefault, type: 'left' };
      // A hanging first line still stops at the text margin (zero), even if
      // the first positive custom stop suppresses ordinary default stops.
      if (current < -.5 && nextDefault <= 0 && nextDefault < stop.position) stop = { position: nextDefault, type: 'left' };
      let following = 0;
      if (stop.type !== 'left') {
        const boundaries = [...paragraph.querySelectorAll('[data-docx-tab],br')];
        const end = boundaries[boundaries.indexOf(tab) + 1];
        const range = document.createRange(); range.selectNodeContents(paragraph); range.setStartAfter(tab);
        if (end) range.setEndBefore(end);
        const measure = document.createElement('span');
        // Apply the document's font fallback rules to cloned styled runs too,
        // while staying outside ProseMirror's observed editing DOM.
        measure.className = 'tiptap';
        // Chromium can return an empty font shorthand for a mixed computed
        // style. Copy longhands so measurement uses the paragraph's typeface.
        Object.assign(measure.style, { position: 'fixed', left: '-10000px', visibility: 'hidden', whiteSpace: 'pre', width: 'max-content', textIndent: '0', padding: '0', minHeight: '0',
          fontFamily: paragraphStyle.fontFamily, fontSize: paragraphStyle.fontSize,
          fontWeight: paragraphStyle.fontWeight, fontStyle: paragraphStyle.fontStyle,
          fontStretch: paragraphStyle.fontStretch, fontVariant: paragraphStyle.fontVariant,
          fontKerning: paragraphStyle.fontKerning, fontFeatureSettings: paragraphStyle.fontFeatureSettings,
          fontVariationSettings: paragraphStyle.fontVariationSettings,
          letterSpacing: paragraphStyle.letterSpacing, wordSpacing: paragraphStyle.wordSpacing,
          textTransform: paragraphStyle.textTransform });
        measure.append(range.cloneContents());
        if (stop.type === 'decimal') {
          // Keep each run's font/size while measuring up to the separator.
          // Flattening to text measured mixed-size numbers in the P font.
          const walker = document.createTreeWalker(measure, NodeFilter.SHOW_TEXT);
          while (walker.nextNode()) {
            const node = walker.currentNode as Text, at = node.data.search(/[.,]/);
            if (at < 0) continue;
            const tail = document.createRange(); tail.selectNodeContents(measure); tail.setStart(node, at); tail.deleteContents(); break;
          }
        }
        // Measuring inside the editor would trigger ProseMirror's DOM observer
        // and reparse the paragraph, discarding the measured node view widths.
        document.body.append(measure); following = measure.getBoundingClientRect().width; measure.remove();
      }
      tab.style.width = `${Math.max(0, stop.position - current - (stop.type === 'center' ? following / 2 : following))}px`;
      tab.dataset.leader = stop.leader ?? '';
    }
  }
}

/** Native mouse/arrow selectionchange can arrive after the next keydown.
 * Custom insertion commands must use the visible caret instead of a stale
 * text selection. Leave node/cell/gap selections and IME composition alone. */
function synchronizeTabSelection(editor: Editor) {
  const { view, state } = editor;
  if (view.composing || !(state.selection instanceof TextSelection) ||
      view.dom.ownerDocument.activeElement !== view.dom) return;
  const selection = view.dom.ownerDocument.getSelection();
  if (!selection?.anchorNode || !selection.focusNode ||
      !view.dom.contains(selection.anchorNode) || !view.dom.contains(selection.focusNode)) return;
  try {
    const anchor = view.posAtDOM(selection.anchorNode, selection.anchorOffset);
    const head = view.posAtDOM(selection.focusNode, selection.focusOffset);
    if (!state.doc.resolve(anchor).parent.inlineContent || !state.doc.resolve(head).parent.inlineContent) return;
    const next = TextSelection.create(state.doc, anchor, head);
    if (!next.eq(state.selection)) view.dispatch(state.tr.setSelection(next));
  } catch {
    // A native selection outside the editable document has no model position.
  }
}

export const DocumentTab = Node.create({
  // Synchronize before lower-priority table/list/code keymaps decide whether
  // to consume Tab; their specialized navigation still handles those contexts.
  priority: 1100,
  name: 'documentTab', group: 'inline', inline: true, atom: true, selectable: false,
  parseHTML: () => [{ tag: 'span[data-docx-tab]' }],
  renderHTML: () => ['span', { 'data-docx-tab': 'true', class: 'document-tab' }, '\t'],
  renderText: () => '\t',
  addNodeView: () => () => {
    const dom = document.createElement('span'); dom.className = 'document-tab'; dom.dataset.docxTab = 'true';
    dom.textContent = '\t'; dom.contentEditable = 'false';
    return { dom, ignoreMutation: () => true };
  },
  addKeyboardShortcuts() {
    return { Tab: () => {
      synchronizeTabSelection(this.editor);
      return this.editor.isActive('table') || this.editor.isActive('listItem') || this.editor.isActive('codeBlock') ? false
        : this.editor.commands.insertContent({ type: this.name });
    } };
  },
});
export const DocumentTabs = Extension.create({
  name: 'documentTabs',
  addGlobalAttributes: () => [{ types: ['paragraph', 'heading'], attributes: {
    tabStops: { default: null, parseHTML: el => el.getAttribute('data-tab-stops'),
      renderHTML: a => a.tabStops ? { 'data-tab-stops': JSON.stringify(parseTabStops(a.tabStops)) } : {} },
  } }],
  addProseMirrorPlugins: () => [new Plugin({ view(view) {
    let frame = 0;
    const schedule = () => { cancelAnimationFrame(frame); frame = requestAnimationFrame(() => { if (view.dom.isConnected) layoutTabs(view.dom); }); };
    const observer = new ResizeObserver(schedule); observer.observe(view.dom);
    document.fonts.addEventListener('loadingdone', schedule); schedule();
    return { update: schedule, destroy() { cancelAnimationFrame(frame); observer.disconnect(); document.fonts.removeEventListener('loadingdone', schedule); } };
  } })],
});
