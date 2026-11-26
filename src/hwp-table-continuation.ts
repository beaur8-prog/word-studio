import { hwpWrapFingerprint, normalizeHwpWrapText } from './hwp-text-advances';

export type HwpTableContinuation = {
  row: number;
  rows: number;
  bodyWidth: number;
  bodyHeight: number;
  tableTop: number;
  cells: string[];
  overhang?: { page: number; end: number };
};

export function parseHwpTableContinuation(value: unknown): HwpTableContinuation | undefined {
  if (typeof value !== 'string' || value.length > 260) return;
  try {
    const parsed = JSON.parse(value);
    const { row, rows, bodyWidth, bodyHeight, tableTop, cells, overhang } = parsed;
    if (!Number.isInteger(row) || !Number.isInteger(rows) || row < 1 || row + 1 >= rows || rows > 10000 ||
        ![bodyWidth, bodyHeight].every(value => Number.isFinite(value) && value > 0 && value <= 10000) ||
        !Number.isFinite(tableTop) || tableTop < 0 || tableTop >= bodyHeight ||
        !Array.isArray(cells) || cells.length < 2 || cells.length > 6 ||
        cells.some(value => typeof value !== 'string' || !/^[a-f0-9]{16}$/.test(value))) return;
    if (overhang !== undefined && (!overhang || !Number.isSafeInteger(overhang.page) ||
        overhang.page < 1 || overhang.page > 10000 || !Number.isFinite(overhang.end) ||
        overhang.end <= bodyHeight || overhang.end > 10000)) return;
    return { row, rows, bodyWidth, bodyHeight, tableTop, cells,
      ...(overhang ? { overhang: { page: overhang.page, end: overhang.end } } : {}) };
  } catch { /* The saved table can still use its ordinary layout. */ }
}

export function matchesHwpTableFragment(
  table: HTMLTableElement,
  body: HTMLElement,
  first: HwpTableContinuation,
  rowIndex: (row: HTMLTableRowElement) => number | undefined,
) {
  const last = table.rows[table.rows.length - 1];
  if (!last || rowIndex(last) !== first.row || last.cells.length !== first.cells.length) return false;
  if (![...last.cells].every((cell, index) =>
    cell.tagName === 'TD' && cell.rowSpan === 1 && cell.colSpan === 1 &&
    !cell.querySelector('table,img,svg,br,[data-note-reference],[data-docx-tab],[data-document-field]') &&
    hwpWrapFingerprint(cell.textContent ?? '') === first.cells[index])) return false;
  const frame = body.getBoundingClientRect(), box = table.getBoundingClientRect();
  // Native geometry is recorded to a tenth of a pixel. DOCX also rounds
  // margins to twips and each browser rounds the two layout edges.
  const precision = .1 + 1 / 15 + 2 / 64;
  return Math.abs(frame.width - first.bodyWidth) <= precision &&
    Math.abs(frame.height - first.bodyHeight) <= precision &&
    Math.abs(box.top - frame.top - first.tableTop) <= precision;
}

/** Include explicit breaks and paragraph boundaries in the original witness. */
export function hwpOverhangCellFingerprint(paragraphs: readonly string[]) {
  return hwpWrapFingerprint(JSON.stringify(paragraphs.map(text => text.split('\n').map(normalizeHwpWrapText))));
}

function printedOverhangFingerprint(cell: HTMLTableCellElement) {
  const paragraphs: string[] = [];
  for (const p of cell.children) {
    if (p.tagName !== 'P') return;
    const walker = document.createTreeWalker(p, NodeFilter.SHOW_ELEMENT | NodeFilter.SHOW_TEXT);
    let text = '';
    while (walker.nextNode()) {
      const node = walker.currentNode;
      if (node instanceof HTMLBRElement) text += '\n';
      else if (node.nodeType === Node.TEXT_NODE) text += node.textContent ?? '';
    }
    paragraphs.push(text);
  }
  return hwpOverhangCellFingerprint(paragraphs);
}

type HwpOverhang = { table: HTMLTableElement; row: HTMLTableRowElement; body: HTMLElement; bottom: number; cells: string[] };

/** Check the real sheet and the intact editable row before offering a cut.
 * This only extends the current table's split boundary. */
export function firstTableOverhang(table: HTMLTableElement, row: HTMLTableRowElement,
  body: HTMLElement, proof: HwpTableContinuation, pageNumber: number): HwpOverhang | undefined {
  const details = proof.overhang, paper = body.closest('.print-page');
  if (!details || details.page !== pageNumber || !paper || body.children.length !== 1 ||
      table.closest('[data-column-count],[data-document-textbox]') ||
      row.cells.length !== proof.cells.length || [...row.cells].some((cell, i) =>
        cell.tagName !== 'TD' || cell.rowSpan !== 1 || cell.colSpan !== 1 ||
        cell.querySelector('table,img,svg,[data-note-reference],[data-docx-tab],[data-document-field]') ||
        printedOverhangFingerprint(cell) !== proof.cells[i])) return;
  const box = table.getBoundingClientRect(), frame = body.getBoundingClientRect(), sheet = paper.getBoundingClientRect();
  const framePrecision = .1 + 1 / 15 + 2 / 64;
  const stroke = Math.max(0, ...[...table.rows[0].cells].map(cell => parseFloat(getComputedStyle(cell).borderTopWidth) || 0));
  const end = frame.top + details.end;
  if (Math.abs(frame.width - proof.bodyWidth) > framePrecision ||
      Math.abs(frame.height - proof.bodyHeight) > framePrecision ||
      Math.abs(box.top - frame.top - proof.tableTop) > framePrecision + stroke / 2 ||
      box.left < frame.left - framePrecision || box.right > frame.right + framePrecision ||
      end >= sheet.bottom || !clearBottomFurniture(paper)) return;
  return { table, row, body, bottom: end, cells: proof.cells };
}

function clearBottomFurniture(paper: Element) {
  const footer = paper.querySelector('.print-footer'), notes = paper.querySelector('.print-footnotes');
  return !footer?.textContent?.trim() && !footer?.querySelector('img,svg,table') && !notes?.children.length;
}

export function activeHwpTableOverhang(overhang: HwpOverhang | undefined, table: Element | null, body: HTMLElement) {
  if (!overhang || overhang.table !== table || overhang.body !== body ||
      overhang.table.rows[overhang.table.rows.length - 1] !== overhang.row ||
      !clearBottomFurniture(body.closest('.print-page')!)) return;
  return overhang;
}

export function visibleHwpTableOverhang(overhang: HwpOverhang | undefined): number | undefined {
  if (!overhang) return;
  const { table, body, row, bottom, cells } = overhang, frame = body.getBoundingClientRect();
  if (table.getBoundingClientRect().bottom > bottom + .5) return;
  const range = document.createRange();
  for (const [i, cell] of [...row.cells].entries()) {
    if (!cell.textContent?.trim() || cell.querySelector('br') || printedOverhangFingerprint(cell) === cells[i]) return;
    range.selectNodeContents(cell);
    if ([...range.getClientRects()].some(rect => rect.width > 0 && rect.height > 0 &&
      (rect.bottom > bottom || rect.left < frame.left - .16 || rect.right > frame.right + .16))) return;
  }
  return bottom;
}
