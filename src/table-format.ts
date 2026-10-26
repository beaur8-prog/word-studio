import { Extension, type Editor } from '@tiptap/core';
import type { Node as PMNode } from '@tiptap/pm/model';
import { CellSelection, cellAround, selectedRect, TableMap } from '@tiptap/pm/tables';
import { closeHistory } from '@tiptap/pm/history';

export const cellStyleProperties = {
  backgroundColor: 'background-color', verticalAlign: 'vertical-align',
  borderTop: 'border-top', borderRight: 'border-right', borderBottom: 'border-bottom', borderLeft: 'border-left',
  paddingTop: 'padding-top', paddingRight: 'padding-right', paddingBottom: 'padding-bottom', paddingLeft: 'padding-left',
  fontFamily: 'font-family', fontSize: 'font-size', textColor: 'color',
  whiteSpace: 'white-space',
} as const;
export type CellStyle = keyof typeof cellStyleProperties;

/** Accept a single CSS property, never a raw style attribute from a document. */
export function cellStyleValue(name: CellStyle, value: unknown): string | null {
  if (typeof value !== 'string' || value.length > 200 || /[;{}\\]|url\s*\(|var\s*\(|expression\s*\(/i.test(value)) return null;
  const style = document.createElement('span').style;
  style.setProperty(cellStyleProperties[name], value);
  return style.getPropertyValue(cellStyleProperties[name]) || null;
}

export const CellAppearance = Extension.create({
  name: 'cellAppearance',
  addGlobalAttributes() {
    return [{ types: ['tableCell', 'tableHeader'], attributes: Object.fromEntries(
      Object.entries(cellStyleProperties).map(([name, property]) => [name, {
        default: null,
        parseHTML: (el: HTMLElement) => cellStyleValue(name as CellStyle, el.style.getPropertyValue(property)),
        renderHTML: (attrs: Record<string, unknown>) => {
          const value = cellStyleValue(name as CellStyle, attrs[name]);
          return value ? { style: `${property}:${value}` } : {};
        },
      }]),
    ) }];
  },
});

export function selectedCells(editor: Editor): { node: PMNode; pos: number }[] {
  const selection = editor.state.selection;
  const cells: { node: PMNode; pos: number }[] = [];
  if (selection instanceof CellSelection) selection.forEachCell((node, pos) => cells.push({ node, pos }));
  else {
    const cell = cellAround(selection.$from);
    if (cell?.nodeAfter) cells.push({ node: cell.nodeAfter, pos: cell.pos });
  }
  return cells;
}

export function selectContextCell(editor: Editor, element?: HTMLElement): boolean {
  let pos: number | undefined;
  if (element) {
    try { pos = cellAround(editor.state.doc.resolve(editor.view.posAtDOM(element, 0)))?.pos; }
    catch { return false; }
  } else pos = selectedCells(editor)[0]?.pos;
  if (pos == null) return false;
  const current = editor.state.selection;
  if (!(current instanceof CellSelection) || !selectedCells(editor).some(cell => cell.pos === pos)) {
    editor.view.dispatch(editor.state.tr.setSelection(CellSelection.create(editor.state.doc, pos)).setMeta('addToHistory', false));
  }
  return true;
}

export function setCellStyles(editor: Editor, attrs: Partial<Record<CellStyle, string | null>>) {
  if (!editor.isEditable) return;
  const tr = closeHistory(editor.state.tr);
  for (const { node, pos } of selectedCells(editor)) tr.setNodeMarkup(pos, null, { ...node.attrs, ...attrs });
  if (tr.docChanged) editor.view.dispatch(tr);
}

export function cellTextValue(editor: Editor, property: 'fontFamily' | 'fontSize' | 'color'): string {
  const values = new Set<string>();
  for (const { node } of selectedCells(editor)) {
    const inherited = node.attrs[property === 'color' ? 'textColor' : property] || '';
    let hasText = false;
    node.descendants(child => {
      if (child.isText) { hasText = true; values.add(child.marks.find(m => m.type.name === 'textStyle')?.attrs[property] || inherited); }
    });
    if (!hasText) values.add(inherited);
  }
  return values.size === 1 ? [...values][0] : '';
}

export function formatCellText(editor: Editor, property: 'fontFamily' | 'fontSize' | 'color', value: string | null) {
  if (!editor.isEditable) return;
  const tr = closeHistory(editor.state.tr), type = editor.schema.marks.textStyle;
  for (const { node, pos } of selectedCells(editor)) {
    // Defaults on the cell also style text subsequently entered into an empty cell.
    tr.setNodeMarkup(pos, null, { ...node.attrs, [property === 'color' ? 'textColor' : property]: value });
    node.descendants((child, offset) => {
      if (child.isText) {
        const attrs = { ...child.marks.find(m => m.type === type)?.attrs, [property]: value };
        tr.addMark(pos + 1 + offset, pos + 1 + offset + child.nodeSize, type.create(attrs));
      }
    });
  }
  if (tr.docChanged) editor.view.dispatch(tr);
}

export function allCellsHaveMark(editor: Editor, name: string): boolean {
  let found = false, all = true;
  for (const { node } of selectedCells(editor)) node.descendants(child => {
    if (child.isText) { found = true; if (!child.marks.some(m => m.type.name === name)) all = false; }
  });
  return found && all;
}

export function toggleCellMark(editor: Editor, name: string, attrs?: Record<string, string>) {
  if (!editor.isEditable) return;
  const tr = closeHistory(editor.state.tr), type = editor.schema.marks[name];
  const remove = !attrs && allCellsHaveMark(editor, name);
  for (const { node, pos } of selectedCells(editor)) {
    if (remove) tr.removeMark(pos + 1, pos + node.nodeSize - 1, type);
    else tr.addMark(pos + 1, pos + node.nodeSize - 1, type.create(attrs));
  }
  if (tr.docChanged) editor.view.dispatch(tr);
}

export function setCellHighlight(editor: Editor, color: string | null) {
  if (!editor.isEditable) return;
  const tr = closeHistory(editor.state.tr), highlight = editor.schema.marks.highlight, textStyle = editor.schema.marks.textStyle;
  for (const { node, pos } of selectedCells(editor)) {
    node.descendants((child, offset) => {
      const mark = child.marks.find(mark => mark.type === textStyle);
      if (child.isText && mark?.attrs.backgroundColor) tr.addMark(pos + 1 + offset, pos + 1 + offset + child.nodeSize, textStyle.create({ ...mark.attrs, backgroundColor: null }));
    });
    if (color) tr.addMark(pos + 1, pos + node.nodeSize - 1, highlight.create({ color }));
    else tr.removeMark(pos + 1, pos + node.nodeSize - 1, highlight);
  }
  if (tr.docChanged) editor.view.dispatch(tr);
}

export function formatCellParagraphs(editor: Editor, attrs: Record<string, unknown>) {
  if (!editor.isEditable) return;
  const tr = closeHistory(editor.state.tr);
  for (const { node, pos } of selectedCells(editor)) node.descendants((child, offset) => {
    if (['paragraph', 'heading'].includes(child.type.name)) tr.setNodeMarkup(pos + 1 + offset, null, { ...child.attrs, ...attrs });
  });
  if (tr.docChanged) editor.view.dispatch(tr);
}

export function clearCellFormatting(editor: Editor) {
  if (!editor.isEditable || !selectedCells(editor).length) return;
  const tr = closeHistory(editor.state.tr);
  const { tableStart } = selectedRect(editor.state);
  for (const [offset, attrs] of borderChanges(editor, 'all', null)) tr.setNodeMarkup(tableStart + offset, null, attrs);
  for (const { node, pos } of selectedCells(editor)) {
    tr.setNodeMarkup(pos, null, { ...node.attrs, ...Object.fromEntries(Object.keys(cellStyleProperties).map(k => [k, null])) });
    node.descendants((child, offset) => {
      const start = pos + 1 + offset;
      // Comments and links are content, not visual formatting.
      if (child.isText) for (const mark of child.marks) if (!['commentAnchor', 'link'].includes(mark.type.name)) tr.removeMark(start, start + child.nodeSize, mark.type);
      if (['paragraph', 'heading'].includes(child.type.name)) tr.setNodeMarkup(start, null, { ...child.attrs, textAlign: null, indent: 0, lineSpacing: null });
    });
  }
  if (tr.docChanged) editor.view.dispatch(tr);
}

export type BorderPreset = 'all' | 'outer' | 'inner' | 'none' | 'top' | 'right' | 'bottom' | 'left' | 'horizontal' | 'vertical';
const sides = ['top', 'right', 'bottom', 'left'] as const;
const attribute = (side: string) => `border${side[0].toUpperCase()}${side.slice(1)}`;

function borderChanges(editor: Editor, preset: BorderPreset, border: string | null) {
  const rect = selectedRect(editor.state), { table, map, tableStart } = rect;
  const changes = new Map<number, Record<string, unknown>>();
  const patch = (offset: number, side: string, value: string | null) => {
    changes.set(offset, { ...(changes.get(offset) ?? table.nodeAt(offset)!.attrs), [attribute(side)]: value });
  };
  for (const { pos } of selectedCells(editor)) {
    const offset = pos - tableStart, cell = map.findCell(offset);
    for (const side of sides) {
      const outer = cell[side] === rect[side];
      const apply = preset === 'all' || preset === 'none' || (preset === 'outer' && outer) ||
        (preset === 'inner' && !outer) || (preset === side && outer) ||
        (preset === 'horizontal' && !outer && ['top', 'bottom'].includes(side)) ||
        (preset === 'vertical' && !outer && ['left', 'right'].includes(side));
      if (!apply) continue;
      const value = preset === 'none' ? '0px none #000000' : border;
      patch(offset, side, value);
      // A collapsed table border is shared. Update the adjoining edge too, so
      // clearing a border or making it thinner is not defeated by its neighbour.
      const horizontal = side === 'top' || side === 'bottom';
      const boundary = cell[side] + (side === 'top' || side === 'left' ? -1 : 0);
      const limit = horizontal ? map.height : map.width;
      if (boundary < 0 || boundary >= limit) continue;
      const start = horizontal ? cell.left : cell.top, end = horizontal ? cell.right : cell.bottom;
      for (let index = start; index < end; index++) {
        const neighbour = map.map[horizontal ? boundary * map.width + index : index * map.width + boundary];
        patch(neighbour, { top: 'bottom', bottom: 'top', left: 'right', right: 'left' }[side], value);
      }
    }
  }
  return changes;
}

export function applyCellBorders(editor: Editor, preset: BorderPreset, border: string) {
  if (!editor.isEditable || !selectedCells(editor).length) return;
  const tr = closeHistory(editor.state.tr), { tableStart } = selectedRect(editor.state);
  for (const [offset, attrs] of borderChanges(editor, preset, border)) tr.setNodeMarkup(tableStart + offset, null, attrs);
  if (tr.docChanged) editor.view.dispatch(tr);
}

export function selectTableRegion(editor: Editor, region: 'row' | 'column' | 'table') {
  const selection = editor.state.selection;
  if (!(selection instanceof CellSelection)) return;
  let next: CellSelection;
  if (region === 'row') next = CellSelection.rowSelection(selection.$anchorCell, selection.$headCell);
  else if (region === 'column') next = CellSelection.colSelection(selection.$anchorCell, selection.$headCell);
  else {
    const table = selection.$anchorCell.node(-1), start = selection.$anchorCell.start(-1), map = TableMap.get(table);
    next = CellSelection.create(editor.state.doc, start + map.map[0], start + map.map[map.map.length - 1]);
  }
  editor.view.dispatch(editor.state.tr.setSelection(next).setMeta('addToHistory', false));
}

export function distributeColumns(editor: Editor) {
  if (!editor.isEditable || !selectedCells(editor).length) return;
  const { map, tableStart, table } = selectedRect(editor.state);
  const dom = editor.view.nodeDOM(tableStart - 1) as HTMLElement | null;
  const tableDOM = dom?.matches('table') ? dom : dom?.querySelector('table');
  // CSS layout width is independent of the document zoom transform.
  const width = Math.max(30, Math.floor((tableDOM?.offsetWidth || map.width * 120) / map.width));
  const tr = closeHistory(editor.state.tr);
  for (const offset of new Set(map.map)) {
    const cell = table.nodeAt(offset)!;
    tr.setNodeMarkup(tableStart + offset, null, { ...cell.attrs, colwidth: Array(cell.attrs.colspan).fill(width) });
  }
  editor.view.dispatch(tr);
}

/** Refine the grid inside the selection; surrounding cells retain their spans and content. */
export function divideSelectedCells(editor: Editor, rows: number, columns: number): boolean {
  if (!editor.isEditable || !Number.isInteger(rows) || !Number.isInteger(columns) || rows < 1 || columns < 1 || rows > 10 || columns > 10 || rows * columns < 2) return false;
  const cells = selectedCells(editor);
  if (!cells.length) return false;
  const rect = selectedRect(editor.state), { map, table, tableStart } = rect;
  const selected = new Set(cells.map(cell => cell.pos - tableStart));
  const columnEdges = [0], rowEdges = [0];
  for (let col = 0; col < map.width; col++) columnEdges.push(columnEdges[col] + (col >= rect.left && col < rect.right ? columns : 1));
  for (let row = 0; row < map.height; row++) rowEdges.push(rowEdges[row] + (row >= rect.top && row < rect.bottom ? rows : 1));
  if (columnEdges.at(-1)! * rowEdges.at(-1)! > 20000) return false;
  const output: { column: number; node: PMNode }[][] = Array.from({ length: rowEdges.at(-1)! }, () => []);
  for (const offset of new Set(map.map)) {
    const cell = table.nodeAt(offset)!, box = map.findCell(offset), split = selected.has(offset);
    const width = columnEdges[box.right] - columnEdges[box.left], height = rowEdges[box.bottom] - rowEdges[box.top];
    const across = split ? columns : 1, down = split ? rows : 1;
    const colwidth: number[] = [];
    for (let col = box.left; col < box.right; col++) {
      const factor = columnEdges[col + 1] - columnEdges[col], old = cell.attrs.colwidth?.[col - box.left];
      colwidth.push(...Array(factor).fill(old ? Math.max(1, Math.round(old / factor)) : 0));
    }
    for (let row = 0; row < down; row++) for (let col = 0; col < across; col++) {
      const attrs = { ...cell.attrs, colspan: width / across, rowspan: height / down,
        colwidth: colwidth.some(Boolean) ? colwidth.slice(col * width / across, (col + 1) * width / across) : null };
      const node = row === 0 && col === 0 ? cell.type.create(attrs, cell.content, cell.marks) : cell.type.createAndFill(attrs)!;
      output[rowEdges[box.top] + row * height / down].push({ column: columnEdges[box.left] + col * width / across, node });
    }
  }
  const newTable = table.type.create(table.attrs, output.map((cells, row) => {
    const originalRow = rowEdges.findIndex((edge, index) => index < rowEdges.length - 1 && edge <= row && rowEdges[index + 1] > row);
    const attrs = table.child(originalRow).attrs;
    return editor.schema.nodes.tableRow.create({ ...attrs,
      rowHeight: attrs.rowHeight ? attrs.rowHeight / (rowEdges[originalRow + 1] - rowEdges[originalRow]) : null,
    }, cells.sort((a, b) => a.column - b.column).map(c => c.node));
  }));
  const nextMap = TableMap.get(newTable);
  if (nextMap.problems?.length) return false;
  const tr = closeHistory(editor.state.tr).replaceWith(tableStart - 1, tableStart - 1 + table.nodeSize, newTable);
  tr.setSelection(CellSelection.create(tr.doc,
    tableStart + nextMap.map[rowEdges[rect.top] * nextMap.width + columnEdges[rect.left]],
    tableStart + nextMap.map[(rowEdges[rect.bottom] - 1) * nextMap.width + columnEdges[rect.right] - 1]));
  editor.view.dispatch(tr);
  return true;
}
