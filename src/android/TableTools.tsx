import { useEffect, useState } from 'react';
import type { Editor } from '@tiptap/core';
import { closeHistory } from '@tiptap/pm/history';
import { selectedRect, TableMap } from '@tiptap/pm/tables';
import { selectContextCell, selectTableRegion, selectedCells, distributeColumns } from '../table-format';
import { MAX_ROW_HEIGHT } from '../table-row-resize';

export function TableTools({ editor, onChange, disabled }: { editor: Editor; onChange(): void; disabled: boolean }) {
  const [height, setHeight] = useState('10'), [width, setWidth] = useState('30');
  const [, refresh] = useState(0);
  useEffect(() => {
    if (editor.isActive('table')) selectContextCell(editor);
    const update = () => refresh(value => value + 1); editor.on('transaction', update);
    return () => { editor.off('transaction', update); };
  }, [editor]);
  const act = (action: () => void) => { if (!disabled && !editor.view.composing) { action(); onChange(); } };
  const button = (name: string, action: () => void) => <button disabled={disabled} onMouseDown={event => event.preventDefault()} onClick={() => act(action)}>{name}</button>;
  const setRowHeight = () => {
    const value = Number(height) * 96 / 25.4;
    if (!(value > 0 && value <= MAX_ROW_HEIGHT)) return;
    const tr = closeHistory(editor.state.tr), rows = new Set<number>();
    for (const cell of selectedCells(editor)) {
      const pos = editor.state.doc.resolve(cell.pos);
      for (let depth = pos.depth; depth > 0; depth--) if (pos.node(depth).type.name === 'tableRow') { rows.add(pos.before(depth)); break; }
    }
    for (const pos of rows) { const row = tr.doc.nodeAt(pos)!; tr.setNodeMarkup(pos, null, { ...row.attrs, rowHeight: value, wordRowRule: null }); }
    if (tr.docChanged) editor.view.dispatch(tr);
  };
  const setColumnWidth = () => {
    const value = Math.round(Number(width) * 96 / 25.4);
    if (!(value >= 30 && value <= 2000) || !selectedCells(editor).length) return;
    const { table, tableStart, left } = selectedRect(editor.state), map = TableMap.get(table), tr = closeHistory(editor.state.tr);
    for (const offset of new Set(map.map)) {
      const rect = map.findCell(offset); if (left < rect.left || left >= rect.right) continue;
      const cell = table.nodeAt(offset)!, widths: number[] = cell.attrs.colwidth ? [...cell.attrs.colwidth] : Array(cell.attrs.colspan).fill(0);
      widths[left - rect.left] = value;
      tr.setNodeMarkup(tableStart + offset, null, { ...cell.attrs, colwidth: widths });
    }
    if (tr.docChanged) editor.view.dispatch(tr);
  };
  return <>
    <p>{editor.isActive('table') ? '선택한 셀·행·열에 적용합니다. 실행 취소로 되돌릴 수 있습니다.' : '표를 넣거나 표 안의 셀을 선택하세요.'}</p>
    {button('표 삽입 (3×3)', () => { editor.chain().focus().insertTable({ rows: 3, cols: 3, withHeaderRow: true }).run(); selectContextCell(editor); })}
    {editor.isActive('table') && <>
      {button('행 선택', () => selectTableRegion(editor, 'row'))}{button('열 선택', () => selectTableRegion(editor, 'column'))}{button('표 전체 선택', () => selectTableRegion(editor, 'table'))}
      {button('아래에 행 삽입', () => { editor.chain().focus().addRowAfter().run(); selectContextCell(editor); })}
      {button('오른쪽에 열 삽입', () => { editor.chain().focus().addColumnAfter().run(); selectContextCell(editor); })}
      {button('셀 병합', () => { editor.chain().focus().mergeCells().run(); selectContextCell(editor); })}
      {button('셀 나누기', () => { editor.chain().focus().splitCell().run(); selectContextCell(editor); })}
      <label>선택 행 최소 높이 (mm)<input aria-label="선택 행 최소 높이 (mm)" type="number" min="1" max="529" value={height} onChange={event => setHeight(event.target.value)} /></label>
      {button('행 높이 적용', setRowHeight)}
      <label>선택 열 너비 (mm)<input aria-label="선택 열 너비 (mm)" type="number" min="8" max="529" value={width} onChange={event => setWidth(event.target.value)} /></label>
      {button('열 너비 적용', setColumnWidth)}{button('열 너비 같게', () => distributeColumns(editor))}
      {button('행 삭제', () => { editor.chain().focus().deleteRow().run(); selectContextCell(editor); })}
      {button('열 삭제', () => { editor.chain().focus().deleteColumn().run(); selectContextCell(editor); })}
    </>}
  </>;
}
