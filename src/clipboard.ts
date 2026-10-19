import { getTextBetween, getTextSerializersFromSchema, type Editor } from '@tiptap/core';
import type { Transaction } from '@tiptap/pm/state';
import type { Slice } from '@tiptap/pm/model';
import { TableMap } from '@tiptap/pm/tables';
import { sanitizeImportedHtml } from './document-html';

const MAX_TABLE_HTML_BYTES = 10 * 1024 * 1024;
const MAX_TABLE_CELLS = 100000;
const MAX_SPAN = 1000;

/** Keep ProseMirror's rich HTML, adding the contract's quoted TSV fallback. */
export function tableClipboardText(slice: Slice): string {
  const table = slice.content.childCount === 1 ? slice.content.firstChild : null;
  // An empty result lets ProseMirror continue to its regular text serializer.
  if (table?.type.name !== 'table') return '';
  const map = TableMap.get(table), seen = new Set<number>();
  const textSerializers = getTextSerializersFromSchema(table.type.schema);
  const rows: string[] = [];
  for (let r = 0; r < map.height; r++) {
    const cells: string[] = [];
    for (let c = 0; c < map.width; c++) {
      const pos = map.map[r * map.width + c];
      let text = '';
      if (!seen.has(pos)) {
        const cell = table.nodeAt(pos)!;
        text = getTextBetween(cell, { from: 0, to: cell.content.size }, { blockSeparator: '\n', textSerializers });
        seen.add(pos);
      }
      cells.push(/[\t\r\n"]/.test(text) ? '"' + text.replace(/"/g, '""') + '"' : text);
    }
    rows.push(cells.join('\t'));
  }
  return rows.join('\n');
}

/** Validate before ProseMirror allocates a table map or replaces the selection. */
export function clipboardHtml(source: string): string {
  if (/<table\b/i.test(source) &&
      (source.length > MAX_TABLE_HTML_BYTES || new Blob([source]).size > MAX_TABLE_HTML_BYTES)) {
    throw new Error('붙여넣을 표 HTML은 10 MB 이하여야 합니다.');
  }
  const html = sanitizeImportedHtml(source).html;
  const template = document.createElement('template');
  template.innerHTML = html;
  let total = 0;
  for (const table of template.content.querySelectorAll('table')) {
    const rows = [...table.rows];
    if (rows.length + total > MAX_TABLE_CELLS) {
      throw new Error('붙여넣을 표는 전체 100,000셀 이하여야 합니다.');
    }
    const occupied: Set<number>[] = Array.from({ length: rows.length }, () => new Set());
    let width = 0;
    const span = (cell: HTMLTableCellElement, key: 'rowspan' | 'colspan') => {
      const value = (cell.getAttribute(key) ?? '1').trim();
      if (!/^\d+$/.test(value) || Number(value) < 1 || Number(value) > MAX_SPAN) {
        throw new Error('붙여넣을 표의 병합 범위가 올바르지 않습니다.');
      }
      return Number(value);
    };
    for (let r = 0; r < rows.length; r++) {
      let c = 0;
      for (const cell of rows[r].cells) {
        while (occupied[r].has(c)) c++;
        const rowSpan = span(cell, 'rowspan'), colSpan = span(cell, 'colspan');
        if (r + rowSpan > rows.length) throw new Error('붙여넣을 표의 병합 범위가 올바르지 않습니다.');
        width = Math.max(width, c + colSpan);
        if (rows.length * width + total > MAX_TABLE_CELLS) {
          throw new Error('붙여넣을 표는 전체 100,000셀 이하여야 합니다.');
        }
        for (let y = r; y < r + rowSpan; y++) {
          for (let x = c; x < c + colSpan; x++) {
            if (occupied[y].has(x)) throw new Error('붙여넣을 표의 병합 셀이 겹칩니다.');
            occupied[y].add(x);
          }
        }
        c += colSpan;
      }
    }
    total += rows.length * width;
  }
  return html;
}

/** Home and cell-menu paste use the same clipboard snapshot and editor guard. */
export async function pasteClipboard(editor: Editor, notify: (message: string) => void) {
  if (editor.isDestroyed || !editor.isEditable) return;
  editor.view.focus();
  const shortcut = /Mac/i.test(navigator.platform) ? '⌘V' : 'Ctrl+V';
  if (window.wordDesktop?.paste) {
    try { await window.wordDesktop.paste(); }
    catch { notify(`문서에서 ${shortcut}로 붙여넣으세요.`); }
    return;
  }

  const { doc, selection } = editor.state;
  let cancelled = false;
  const changed = ({ transaction }: { transaction: Transaction }) => {
    if (transaction.docChanged || transaction.selectionSet) cancelled = true;
  };
  const blurred = () => { cancelled = true; };
  editor.on('transaction', changed);
  editor.on('blur', blurred);
  try {
    let html = '', text = '';
    if (navigator.clipboard.read) {
      const items = await navigator.clipboard.read();
      // Never combine HTML and plain text from different ClipboardItems.
      const item = items.find(item => item.types.includes('text/html')) ??
        items.find(item => item.types.includes('text/plain'));
      if (item?.types.includes('text/html')) html = await (await item.getType('text/html')).text();
      if (!html && item?.types.includes('text/plain')) text = await (await item.getType('text/plain')).text();
    } else text = await navigator.clipboard.readText();

    if (editor.isDestroyed) return;
    if (cancelled || !editor.isEditable || editor.state.doc !== doc || !editor.state.selection.eq(selection)) {
      notify('삽입 위치나 편집 상태가 바뀌었습니다. 다시 붙여넣으세요.');
      return;
    }
    if (html) editor.view.pasteHTML(clipboardHtml(html));
    else if (text) editor.view.pasteText(text);
    else notify('클립보드에 붙여넣을 글이나 표가 없습니다.');
  } catch (error) {
    if (!editor.isDestroyed) notify(error instanceof Error &&
      error.message.startsWith('붙여넣을 표') ? error.message : `문서에서 ${shortcut}로 붙여넣으세요.`);
  } finally {
    editor.off('transaction', changed);
    editor.off('blur', blurred);
  }
}
