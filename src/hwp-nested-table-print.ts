import {
  hwpNestedCellFaceAttribute, hwpNestedCellToken, hwpNestedLayoutFingerprint, hwpNestedParagraphToken, hwpNestedTableLayoutAttribute,
  hwpNestedTableToken, parseHwpNestedTableLayout, type HwpNestedTableCut, type HwpNestedTableGap,
} from './hwp-nested-table-layout';

type GridCell = { cell: HTMLTableCellElement; row: number; column: number; end: number };

const paragraph = 'p,h1,h2,h3,h4,h5,h6';
const closeMargin = (a: number, b: number) => Math.abs(a - b) <= .08;
const pixels = (value: string) => value === '' || value === '0' || /^\d+(?:\.\d+)?px$/.test(value)
  ? parseFloat(value) || 0 : NaN;
const columns = (table: Element) => table.querySelectorAll(':scope > colgroup > col').length;

function gapPlan(blocks: HTMLElement[], gaps: HwpNestedTableGap[]) {
  const margins: { left: HTMLElement; right: HTMLElement; bottom: number; top: number }[] = [];
  for (const [at, after, before, sourceTrailing] of gaps) {
    const left = blocks[at], right = blocks[at + 1];
    const currentBottom = pixels(left.style.marginBottom), currentTop = pixels(right.style.marginTop);
    if (!Number.isFinite(currentBottom) || !Number.isFinite(currentTop) ||
        (currentTop > 0 && !closeMargin(currentTop, before))) return;
    if (sourceTrailing !== null && (!left.matches('table') || !closeMargin(currentBottom, sourceTrailing))) return;
    if (sourceTrailing === null && (!left.matches(paragraph) || (currentBottom > 0 && !closeMargin(currentBottom, after)))) return;
    let bottom = sourceTrailing !== null ? currentBottom + after : currentBottom || after;
    let top = currentTop || before;
    if (bottom > 0 && top > 0) {
      bottom += top;
      top = 0;
    }
    margins.push({ left, right, bottom, top });
  }
  return margins;
}

function cutPlan(table: HTMLElement, proof: HwpNestedTableCut) {
  if (!(table instanceof HTMLTableElement) || table.rows.length !== proof.rows || columns(table) !== proof.columns ||
      table.tBodies.length !== 1 || table.tHead || table.tFoot || table.caption ||
      table.matches('[data-table-anchor],[data-cell-table-position]') || table.querySelector('table')) return;
  const rows = Array.from(table.rows), placed: GridCell[] = [], occupied = Array<number>(proof.columns).fill(0);
  for (const [row, entry] of rows.entries()) {
    let column = 0;
    for (const cell of entry.cells) {
      while (occupied[column] > row) column++;
      const end = row + cell.rowSpan, lastColumn = column + cell.colSpan;
      if (cell.rowSpan < 1 || end > rows.length || lastColumn > proof.columns ||
          occupied.slice(column, lastColumn).some(n => n > row)) return;
      placed.push({ cell, row, column, end });
      for (let col = column; col < lastColumn; col++) occupied[col] = end;
      column = lastColumn;
    }
    if (occupied.some(n => n <= row)) return;
  }
  const tokens = placed.map(({ cell, row, column }) => hwpNestedCellToken(row, column, cell.rowSpan, cell.colSpan, cell.textContent ?? ''));
  if (hwpNestedLayoutFingerprint(tokens) !== proof.fingerprint) return;
  const boundary = proof.boundary, crossing = placed.filter(cell => cell.row < boundary && cell.end > boundary);
  if (crossing.length !== proof.merges.length || crossing.some(({ cell, row, column }, i) =>
    row !== proof.merges[i][0] || column !== proof.merges[i][1] || cell.rowSpan !== proof.merges[i][2] ||
    cell.matches('[data-hwp-cell-raster-fill]') ||
    cell.querySelector('img,svg,[data-note-reference],[data-document-field],[data-document-textbox]'))) return;
  return { table, rows, placed, crossing, boundary };
}

function divideNestedTable(plan: NonNullable<ReturnType<typeof cutPlan>>) {
  const { table, rows, placed, crossing, boundary } = plan;
  const tail = table.cloneNode(false) as HTMLTableElement;
  tail.removeAttribute('id');
  for (const grid of table.querySelectorAll(':scope > colgroup')) tail.append(grid.cloneNode(true));
  const body = table.tBodies[0].cloneNode(false) as HTMLTableSectionElement;
  tail.append(body);
  table.style.marginBottom = '0';
  for (const { cell, row, column, end } of crossing) {
    const continuation = cell.cloneNode(false) as HTMLTableCellElement;
    continuation.removeAttribute('id');
    cell.rowSpan = boundary - row;
    continuation.rowSpan = end - boundary;
    const following = placed.find(item => item.row === boundary && item.column > column);
    rows[boundary].insertBefore(continuation, following?.cell ?? null);
  }
  body.append(...rows.slice(boundary));
  table.after(tail);
}

/** Apply optional source geometry to a print copy. Editing and exporting
 * still use the single original table, rows and merged cells. */
export function layoutHwpNestedTables(source: HTMLElement) {
  for (const table of source.querySelectorAll<HTMLTableElement>(`table[${hwpNestedTableLayoutAttribute}]`)) {
    const layout = parseHwpNestedTableLayout(table.getAttribute(hwpNestedTableLayoutAttribute));
    const row = table.rows[0], cell = row?.cells[0];
    if (!layout || table.parentElement?.closest('td,th') || table.rows.length !== 1 || row.cells.length !== 1 ||
        cell.tagName !== 'TD' || cell.colSpan !== 1 || cell.rowSpan !== 1) continue;
    const blocks = Array.from(cell.children as HTMLCollectionOf<HTMLElement>).filter(element =>
      !(element.matches('p') && !element.textContent?.trim() && element.nextElementSibling?.matches('table') &&
        parseFloat(element.style.lineHeight) <= 1));
    if (blocks.length !== layout.blocks || blocks.some(block => !block.matches(`${paragraph},table`))) continue;
    const tokens = blocks.map(block => block instanceof HTMLTableElement
      ? hwpNestedTableToken(block.rows.length, columns(block)) : hwpNestedParagraphToken(block.textContent ?? ''));
    if (hwpNestedLayoutFingerprint(tokens) !== layout.fingerprint) continue;
    const gaps = gapPlan(blocks, layout.gaps), cuts = layout.cuts.map(cut => cutPlan(blocks[cut.block], cut));
    if (!gaps || cuts.some(cut => !cut)) continue;
    for (const { left, right, bottom, top } of gaps) {
      if (pixels(left.style.marginBottom) !== bottom) left.style.marginBottom = `${bottom}px`;
      if (pixels(right.style.marginTop) !== top) right.style.marginTop = `${top}px`;
    }
    for (const cut of cuts) if (cut) divideNestedTable(cut);
    for (const face of layout.faces ?? []) {
      const table = blocks[face.block];
      const row = table instanceof HTMLTableElement ? table.rows[face.row] : undefined;
      if (!row || row.cells.length !== columns(table) ||
          Array.from(row.cells).some(cell => cell.colSpan !== 1 || cell.rowSpan !== 1)) continue;
      const cell = row.cells[face.column], paragraph = cell?.firstElementChild;
      if (cell?.tagName === 'TD' && cell.children.length === 1 && paragraph?.tagName === 'P')
        paragraph.setAttribute(hwpNestedCellFaceAttribute, JSON.stringify(face));
    }
  }
}
