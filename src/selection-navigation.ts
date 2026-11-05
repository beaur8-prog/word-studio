import { Extension } from '@tiptap/core';
import { TextSelection } from '@tiptap/pm/state';

// Native range collapse can stop at our noneditable flow widgets. The table
// keymap can instead move a selected cell paragraph into its neighbouring cell.
// Leave bidi text to the browser, whose horizontal navigation is visual.
const simpleLtrText = /^[\p{Script=Latin}\p{Script=Hangul}\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Greek}\p{Script=Cyrillic}\p{Number}\p{Punctuation}\p{Symbol}\p{Mark}\p{Separator}\s]*$/u;

export const SelectionNavigation = Extension.create({
  name: 'selectionNavigation',
  priority: 1200,
  addKeyboardShortcuts() {
    const collapse = (right: boolean) => {
      const { state, view } = this.editor, selection = state.selection;
      if (view.composing || !(selection instanceof TextSelection) || selection.empty ||
          !selection.$from.sameParent(selection.$to) || !selection.$from.parent.isTextblock ||
          !simpleLtrText.test(selection.$from.parent.textContent)) return false;
      const block = view.nodeDOM(selection.$from.before());
      if (!(block instanceof HTMLElement) || getComputedStyle(block).direction !== 'ltr') return false;
      let inCell = false;
      for (let depth = selection.$from.depth - 1; depth > 0; depth--) {
        if (['cell', 'header_cell'].includes(selection.$from.node(depth).type.spec.tableRole)) {
          inCell = true; break;
        }
      }
      if (!inCell && !block.querySelector('[data-flow-widget]')) return false;
      view.dispatch(state.tr.setSelection(TextSelection.create(state.doc, right ? selection.to : selection.from)).scrollIntoView());
      return true;
    };
    return { ArrowLeft: () => collapse(false), ArrowRight: () => collapse(true) };
  },
});
