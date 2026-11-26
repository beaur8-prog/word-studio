import { paragraphLineOffsets } from './paragraph-fragments';
import { tableAnchor } from './table-anchor-data';

const topPadding = '--hwp-centered-padding-top';
const bottomPadding = '--hwp-centered-padding-bottom';
const centeredTables = 'table:is([data-hwp-centered-borders="true"],[data-hwp-centered-borders="inline"],[data-hwp-centered-borders="continuation"])';
let nextTableScope = 0;

type CellInset = { cell: HTMLTableCellElement; row: number; column: number; top: string; bottom: string };

type CellEdges = { element: HTMLTableCellElement; start: number; end: number; endRow: number;
  top: number; bottom: number; paddingTop: number; paddingBottom: number };

function rowEdges(rows: HTMLTableRowElement[], inline: boolean, continuation: boolean): CellEdges[][] | null {
  const result: CellEdges[][] = [];
  const allowedBorders = inline ? ['solid', 'none', 'double'] : ['solid', 'none'];
  let open: CellEdges[] = [];
  let columns = 0;
  for (const [index, row] of rows.entries()) {
    open = open.filter(previous => previous.endRow > index);
    const occupied = open.slice().sort((a, b) => a.start - b.start);
    const cells: CellEdges[] = [];
    let column = 0, interval = 0;
    const skipSpans = () => {
      while (occupied[interval]?.start === column) column = occupied[interval++].end;
    };
    for (const cell of row.cells) {
      skipSpans();
      const css = getComputedStyle(cell);
      if ((!inline && !continuation && cell.children.length !== 1) ||
          cell.rowSpan < 1 || index + cell.rowSpan > rows.length || !cell.children.length ||
          [...cell.children].some(child => child.tagName !== 'P') ||
          !allowedBorders.includes(css.borderTopStyle) || !allowedBorders.includes(css.borderBottomStyle) ||
          !cell.style.paddingTop.endsWith('px') || !cell.style.paddingBottom.endsWith('px')) return null;
      const end = column + cell.colSpan;
      if (occupied[interval] && occupied[interval].start < end) return null;
      const edge = { element: cell, start: column, end, endRow: index + cell.rowSpan,
        top: parseFloat(css.borderTopWidth), bottom: parseFloat(css.borderBottomWidth),
        paddingTop: parseFloat(cell.style.paddingTop), paddingBottom: parseFloat(cell.style.paddingBottom) };
      cells.push(edge);
      if (cell.rowSpan > 1) open.push(edge);
      column = end;
    }
    skipSpans();
    if (!index) columns = column;
    if (!cells.length || column !== columns || interval !== occupied.length) return null;
    result.push(cells);
  }
  return result;
}

/** Paint the usual collapsed borders. Only the half-stroke that Chromium
 * reserves inside a cell belongs to its original padding budget. */
export function layoutHwpCenteredTableBorders(table: HTMLTableElement, paintInlineGrid = true): void {
  if (!table.matches(centeredTables)) return;
  const active = planInsets(table, paintInlineGrid);
  if (!active) { delete table.dataset.hwpCenteredBordersActive; return; }
  for (const { cell, top, bottom } of active) {
    const previous = cell.style;
    if (previous.getPropertyValue(topPadding) !== top) previous.setProperty(topPadding, top);
    if (previous.getPropertyValue(bottomPadding) !== bottom) previous.setProperty(bottomPadding, bottom);
  }
  if (table.dataset.hwpCenteredBordersActive !== 'true') table.dataset.hwpCenteredBordersActive = 'true';
}

/** ProseMirror redraws cell style from the model. Keep measured padding in
 * a view stylesheet, scoped to the live table and its copied headers. */
export function layoutEditorHwpCenteredBorders(root: HTMLElement, style: HTMLStyleElement) {
  const rules: string[] = [];
  for (const table of root.querySelectorAll<HTMLTableElement>(centeredTables)) {
    const insets = planInsets(table);
    if (!insets) { delete table.dataset.hwpCenteredBordersActive; continue; }
    if (!table.dataset.hwpBorderScope) table.dataset.hwpBorderScope = String(++nextTableScope);
    if (table.dataset.hwpCenteredBordersActive !== 'true') table.dataset.hwpCenteredBordersActive = 'true';
    const scope = `table[data-hwp-centered-borders-active="true"][data-hwp-border-scope="${table.dataset.hwpBorderScope}"]`;
    const selectors = new Map<string, string[]>();
    for (const { row, column, top, bottom } of insets) {
      const properties = `padding-top:${top}!important;padding-bottom:${bottom}!important`;
      const group = selectors.get(properties) ?? [];
      group.push(`${scope}>:is(tbody,thead)>tr:nth-child(${row + 1} of :not(.ProseMirror-widget,.pagination-gap))>:nth-child(${column + 1})`);
      selectors.set(properties, group);
    }
    for (const [declaration, cells] of selectors) rules.push(`${cells.join(',')}{${declaration}}`);
  }
  const css = rules.join('\n');
  if (style.textContent !== css) style.textContent = css;
}

function planInsets(table: HTMLTableElement, paintInlineGrid = true): CellInset[] | null {
  const anchor = tableAnchor(table.dataset.tableAnchor);
  const inline = table.dataset.hwpCenteredBorders === 'inline';
  const continuation = table.dataset.hwpCenteredBorders === 'continuation';
  if (inline && anchor || table.parentElement?.closest('td,th,[data-column-count]:not([data-column-count="1"])') ||
    anchor && (anchor.reference !== 'paragraph' || anchor.wrap || anchor.overlay) || table.dataset.wordTableWidthAuto === 'true' || getComputedStyle(table).borderCollapse !== 'collapse' ||
    table.querySelector('table,img,svg,[data-note-reference],[data-document-textbox],[data-docx-tab]') ||
    !continuation && table.querySelector('br:not(.ProseMirror-trailingBreak)')) return null;
  const rows = [...table.rows].filter(row => !row.matches('.ProseMirror-widget,.pagination-gap'));
  if (rows.length < 2 || !rows[0].cells.length) return null;
  const edges = rowEdges(rows, inline, continuation);
  if (!edges) return null;
  if ([...rows[0].cells].some(cell => cell.tagName !== 'TH')) {
    const cells = edges.flat();
    const separators = cells.filter(cell => cell.element.rowSpan > 1);
    if (!paintInlineGrid || !inline || rows.length < 3 || !separators.length ||
        cells.some(cell => cell.element.tagName !== 'TD') ||
        separators.some(cell => !edges[0].includes(cell) || cell.endRow !== rows.length || cell.top || cell.bottom ||
          cell.element.children.length !== 1 || cell.element.textContent?.trim()) ||
        cells.some(cell => cell.element.rowSpan === 1 && !(cell.top > 0 && cell.bottom > 0))) return null;
  }
  const ending: CellEdges[][] = Array.from({ length: rows.length + 1 }, () => []);
  for (const cells of edges) for (const cell of cells) ending[cell.endRow].push(cell);

  const result: CellInset[] = [];
  for (const [row, cells] of edges.entries()) for (const [column, cell] of cells.entries()) {
    const overlaps = (other: CellEdges) => other.start < cell.end && cell.start < other.end;
    const above = ending[row].filter(overlaps);
    const below = edges[cell.endRow]?.filter(overlaps) ?? [];
    const top = Math.max(cell.top, ...above.map(other => other.bottom)) / 2;
    const bottom = Math.max(cell.bottom, ...below.map(other => other.top)) / 2;
    if (!continuation && !(cell.paddingTop >= top && cell.paddingBottom >= bottom)) return null;
    result.push({ cell: cell.element, row, column, top: `${Math.max(0, cell.paddingTop - top)}px`,
      bottom: `${Math.max(0, cell.paddingBottom - bottom)}px` });
  }
  return result;
}

/** The logical HWP grid uses the authored row minimum or the exact line
 * plus its cell insets. A collapsed CSS border and mixed inline baselines
 * can add a different fractional amount in each Chromium version. */
function centeredRowsFit(table: HTMLTableElement, last: HTMLTableRowElement, bottom: number): boolean {
  const pixels = (value: string) => value.endsWith('px') ? parseFloat(value) : NaN;
  const paragraphs: HTMLParagraphElement[] = [];
  let logicalBottom = table.rows[0].getBoundingClientRect().top;
  for (const row of table.rows) {
    if (row.dataset.wordRowRule === 'exact') return false;
    let height = pixels(row.style.height) || 0;
    for (const cell of row.cells) {
      const paragraph = cell.firstElementChild;
      if (!(paragraph instanceof HTMLParagraphElement) ||
          paragraph.dataset.wordLineRule !== 'exact' && !paragraph.hasAttribute('data-exact-line-height') ||
          paragraph.dataset.printParagraphFragment || paragraph.querySelector('br,img,svg,sub,sup,[data-word-overlap],[data-page-number],[data-document-field]'))
        return false;
      const line = pixels(paragraph.style.lineHeight);
      const padding = pixels(cell.style.paddingTop) + pixels(cell.style.paddingBottom);
      const margins = (parseFloat(paragraph.style.marginTop) || 0) + (parseFloat(paragraph.style.marginBottom) || 0);
      if (!(line > 0) || !(padding >= 0) || margins !== 0) return false;
      height = Math.max(height, pixels(cell.style.height) || 0, line + padding);
      paragraphs.push(paragraph);
    }
    logicalBottom += height;
    if (logicalBottom > bottom) return false;
  }
  // A new line or resized row must reserve normal DOM space. Keep the
  // original paint and permit only the bottom inset to reach into the margin.
  if (paragraphs.some(paragraph => paragraphLineOffsets(paragraph).length)) return false;
  const range = document.createRange();
  for (const cell of last.cells) {
    const bounds = cell.getBoundingClientRect();
    range.selectNodeContents(cell.firstElementChild!);
    for (const box of range.getClientRects()) {
      if (box.width > 0 && box.height > 0 && (box.bottom > bottom || box.left < bounds.left || box.right > bounds.right))
        return false;
    }
  }
  return true;
}

/** The stroke's outside half can enter free page margin. Its physical box
 * remains inside the sheet and clear of page furniture. */
export function tableBorderFits(element: HTMLElement, body: HTMLElement, bottom: number): boolean {
  if (!(element instanceof HTMLTableElement) || element.dataset.hwpCenteredBordersActive !== 'true') return false;
  const row = element.rows[element.rows.length - 1], sheet = body.closest('.print-page');
  if (!row || !sheet || sheet.querySelector('.print-footnotes [data-document-note]')) return false;
  const rowBottom = row.getBoundingClientRect().bottom, outside = element.getBoundingClientRect().bottom;
  const footer = sheet.querySelector('.print-footer .document-furniture')?.getBoundingClientRect();
  return rowBottom < outside && outside <= sheet.getBoundingClientRect().bottom &&
    (!footer?.height || outside <= footer.top) &&
    (rowBottom <= bottom || centeredRowsFit(element, row, bottom));
}
