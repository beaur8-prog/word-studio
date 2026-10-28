import { Extension } from '@tiptap/core';
import { TextSelection } from '@tiptap/pm/state';
import { GapCursor } from '@tiptap/pm/gapcursor';
import { TableMap } from '@tiptap/pm/tables';

/** A trailing table has no synthetic paragraph. PC table arrow navigation
 * otherwise resolves back into the last cell when no following text exists. */
export const TableNavigation = Extension.create({
  name: 'tableNavigation',
  priority: 1100,
  addKeyboardShortcuts() {
    return { ArrowDown: () => {
      const { state, view } = this.editor, selection = state.selection;
      if (!(selection instanceof TextSelection) || !selection.empty || !view.endOfTextblock('down')) return false;
      const pos = selection.$head;
      for (let depth = pos.depth - 1; depth > 1; depth--) {
        if (pos.indexAfter(depth) !== pos.node(depth).childCount) return false;
        if (!['cell', 'header_cell'].includes(pos.node(depth).type.spec.tableRole)) continue;
        const tableDepth = depth - 2, table = pos.node(tableDepth);
        if (table.type.spec.tableRole !== 'table') return false;
        const cell = TableMap.get(table).findCell(pos.before(depth) - pos.start(tableDepth));
        if (cell.bottom !== table.childCount || pos.indexAfter(tableDepth - 1) !== pos.node(tableDepth - 1).childCount) return false;
        const after = state.doc.resolve(pos.after(tableDepth));
        // The table is isolating and this is the end of its parent. Only a
        // parent that accepts a text block can turn this gap into typed text.
        if (after.parent.inlineContent || after.parent.type.spec.allowGapCursor === false ||
            !after.parent.contentMatchAt(after.index()).defaultType?.isTextblock) return false;
        view.dispatch(state.tr.setSelection(new GapCursor(after)).scrollIntoView());
        return true;
      }
      return false;
    } };
  },
});
