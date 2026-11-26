import { documentScale } from './document-scale';

type Position = { x: number; y: number };
type Extent = Position & { width: number };
const twips = 15;
const supportedTableRows = 1;
const supportedColumns = 16;
const positionsMatch = (first: Extent, second: Extent, available: number) =>
  first.y === second.y && first.width > 0 && second.width > 0 &&
  first.x + first.width < second.x && second.x + second.width <= available + .5;

export function cellTablePosition(value: unknown): Position | undefined {
  const m = typeof value === 'string' && /^(\d{1,5}) (\d{1,5})$/.exec(value);
  if (!m) return;
  const x = Number(m[1]), y = Number(m[2]);
  return x <= 32767 && y <= 1440 ? { x: x / twips, y: y / twips } : undefined;
}

const outsideBody = '.document-note,.document-furniture,.document-textbox,.print-footnotes';
const inlineControl = 'br:not(.ProseMirror-trailingBreak),img:not(.ProseMirror-separator),svg,table,video,canvas,math,[data-document-textbox],[data-document-field],[data-note-reference],[data-page-number],[data-docx-tab],[data-docx-bookmark],[data-symbol-font],[data-word-overlap]';
const frameTable = (el: Element) => el instanceof HTMLTableElement ? el
  : el.matches('.tableWrapper,[data-cell-table-copy-frame]') ? el.querySelector<HTMLTableElement>(':scope > table') : null;

function renderedExtent(table: HTMLTableElement): Extent | undefined {
  const pos = cellTablePosition(table.dataset.cellTablePosition);
  const cells = table.rows[0]?.cells;
  if (!pos || table.rows.length !== supportedTableRows || !cells?.length || cells.length > supportedColumns ||
    table.dataset.tableAnchor || table.dataset.wordTableWidthAuto || table.dataset.wordTableLayout === 'autofit') return;
  let width = 0;
  for (const cell of cells) {
    const blocks = [...cell.children].filter(child => !child.matches('.ProseMirror-widget'));
    const column = Number(cell.getAttribute('colwidth') ?? cell.getAttribute('data-colwidth'));
    if (!column || cell.tagName !== 'TD' || cell.colSpan !== 1 || cell.rowSpan !== 1 ||
      blocks.length !== 1 || blocks[0].tagName !== 'P' || blocks[0].querySelector(inlineControl)) return;
    width += column;
  }
  return { ...pos, width };
}

const property = (el: HTMLElement, name: string, pixels: number) => {
  const value = `${pixels}px`;
  if (el.style.getPropertyValue(name) !== value) el.style.setProperty(name, value);
};

function clearPosition(owner: HTMLTableCellElement) {
  for (const child of [...owner.children]) if (child instanceof HTMLElement && child.classList.contains('cell-table-frame')) {
    if (child.dataset.cellTableCopyFrame) { child.replaceWith(...child.childNodes); continue; }
    child.classList.remove('cell-table-frame');
    for (const name of ['--cell-table-left', '--cell-table-top', '--cell-table-width', '--cell-table-frame-width']) child.style.removeProperty(name);
  }
}

/** Move existing view frames without moving either editable table node. */
export function layoutCellTablePairs(root: HTMLElement) {
  const owners = new Set<HTMLTableCellElement>();
  for (const frame of root.querySelectorAll('.cell-table-frame'))
    if (frame.parentElement instanceof HTMLTableCellElement) owners.add(frame.parentElement);
  for (const table of root.querySelectorAll('table[data-cell-table-position]')) {
    const cell = table.parentElement?.closest('td');
    if (cell) owners.add(cell);
  }
  for (const owner of owners) {
    const frames: HTMLElement[] = [], tables: HTMLTableElement[] = [], entries: Extent[] = [], carriers: HTMLElement[] = [];
    for (const child of owner.children) {
      if (!(child instanceof HTMLElement) || child.matches('.ProseMirror-widget')) continue;
      const table = frameTable(child);
      if (table) {
        const extent = renderedExtent(table);
        if (!extent) break;
        frames.push(child); tables.push(table); entries.push(extent);
      } else if (child.tagName === 'P' && !child.textContent?.trim() && !child.querySelector(inlineControl)) carriers.push(child);
      else break;
    }
    const scale = documentScale(owner), style = getComputedStyle(owner);
    const metric = (value: string) => parseFloat(value) || 0;
    const left = metric(style.borderLeftWidth) + metric(style.paddingLeft);
    const top = metric(style.borderTopWidth) + metric(style.paddingTop);
    const right = metric(style.borderRightWidth) + metric(style.paddingRight);
    const available = owner.getBoundingClientRect().width / scale - right;
    const children = [...owner.children].filter(child => !child.matches('.ProseMirror-widget')).length;
    if (owner.rowSpan !== 1 || owner.closest(outsideBody) || carriers.length > 3 || frames.length !== 2 ||
      children !== frames.length + carriers.length || !positionsMatch(entries[0], entries[1], available) || entries[0].x < left - .5 || entries[0].y < top - .5) {
      clearPosition(owner); continue;
    }
    frames.forEach((initial, i) => {
      const { width } = entries[i];
      // Print HTML has a bare table. The live editor already has a TableView
      // frame whose geometry can change without reparsing its editable cell.
      let frame = initial;
      if (initial === tables[i]) {
        frame = document.createElement('div'); frame.dataset.cellTableCopyFrame = 'true';
        initial.replaceWith(frame); frame.append(initial); frames[i] = frame;
      }
      property(frame, '--cell-table-width', width);
      if (!frame.classList.contains('cell-table-frame')) frame.classList.add('cell-table-frame');
    });
    const widths = tables.map(table => table.getBoundingClientRect().width / scale);
    if (!positionsMatch({ ...entries[0], width: widths[0] }, { ...entries[1], width: widths[1] }, available)) {
      clearPosition(owner); continue;
    }
    frames.forEach((frame, i) => {
      const { x, y } = entries[i];
      property(frame, '--cell-table-frame-width', widths[i]);
      property(frame, '--cell-table-left', i === 0 ? x - left : x - entries[0].x - widths[0]);
      property(frame, '--cell-table-top', y - top);
    });
    const origin = owner.getBoundingClientRect();
    if (tables.some((table, i) => {
      const rect = table.getBoundingClientRect();
      return Math.abs((rect.left - origin.left) / scale - entries[i].x) > 1 ||
        Math.abs((rect.top - origin.top) / scale - entries[i].y) > 1 ||
        rect.right > origin.right - right * scale + .5 || rect.bottom > origin.bottom + .5;
    })) clearPosition(owner);
  }
}
