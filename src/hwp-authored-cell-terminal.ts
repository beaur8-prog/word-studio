import { parseHwpCellLine } from './hwp-cell-line';
import { layoutHwpCellLines } from './hwp-cell-line-metrics';
import { hwpWrapFingerprint } from './hwp-text-advances';

/** Compensate for a browser's extra text strut only in a print copy. Text,
 * sizing, padding and the authored row minimum stay in the editable model. */
export function layoutHwpAuthoredCellTerminals(table: HTMLTableElement) {
  const candidates = [...table.querySelectorAll<HTMLParagraphElement>('td > p[data-hwp-cell-line^="2 "]')];
  if (!candidates.length) return;
  const mergedRows = new Set<HTMLTableRowElement>(), rows = Array.from(table.rows);
  for (const [index, row] of rows.entries()) for (const cell of row.cells) {
    if (cell.rowSpan > 1) for (const merged of rows.slice(index, index + cell.rowSpan)) mergedRows.add(merged);
  }
  const intrinsicTolerance = .12;
  const range = document.createRange();
  const rects = (p: HTMLParagraphElement) => {
    const cell = p.parentElement!.getBoundingClientRect();
    range.selectNodeContents(p);
    return [...range.getClientRects()].filter(rect => rect.width > .1 && rect.height > 1)
      .map(rect => [rect.left - cell.left, rect.top - cell.top, rect.right - cell.left, rect.bottom - cell.top]);
  };
  layoutHwpCellLines(table);
  const measured = candidates.flatMap(p => {
    const line = parseHwpCellLine(p.dataset.hwpCellLine), proof = line?.authoredTerminal;
    const cell = p.parentElement;
    const row = cell?.closest('tr');
    if (!line || !proof || !(cell instanceof HTMLTableCellElement) || cell.rowSpan !== 1 || cell.children.length !== 1 ||
        p.closest('table') !== table || !row || row.querySelector('th') ||
        p.style.height || p.querySelector('a,br,img,svg,sub,sup,[data-docx-tab],[data-page-number],[data-note-reference],[data-hwp-cell-prefix]') ||
        hwpWrapFingerprint(p.textContent ?? '') !== proof.text) return [];
    const size = cell.getBoundingClientRect(), box = p.getBoundingClientRect(), style = getComputedStyle(p);
    const alignment = getComputedStyle(cell).verticalAlign, centered = alignment === 'middle';
    if ((alignment !== 'top' && !centered) || centered !== !!proof.middle || (centered && mergedRows.has(row)) ||
        Math.abs(size.width - proof.width) > 1 ||
        Math.abs(parseFloat(style.lineHeight) - line.end) > intrinsicTolerance ||
        box.height <= line.end + intrinsicTolerance) return [];
    const ink = rects(p);
    // Plain terminal text occupies one line. A changed font may introduce a
    // second line even though the original and current strings still match.
    if (!ink.length || ink.some(r => Math.max(r[1], ink[0][1]) >= Math.min(r[3], ink[0][3]))) return [];
    return [{ p, cell, end: line.end, row, minimum: proof.row, ink, centered }];
  });
  for (const item of measured) item.p.style.height = `${item.end}px`;
  const rejected = new Set(measured.filter(({ p, cell, minimum, ink }) => {
    const size = cell.getBoundingClientRect(), actual = rects(p);
    return size.height < minimum - intrinsicTolerance || actual.length !== ink.length ||
      actual.some((r, index) => r[0] < -.5 || r[1] < -.5 || r[2] > size.width + .5 || r[3] > size.height - .5 ||
        r.some((value, edge) => Math.abs(value - ink[index][edge]) > intrinsicTolerance));
  }));
  // Two centered cells can affect the same row height. Keep their geometry
  // together if any trial in that row changes the previously measured ink.
  const centeredRows = new Set(measured.filter(item => item.centered).map(item => item.row));
  const rejectedRows = new Set([...rejected].map(item => item.row).filter(row => centeredRows.has(row)));
  for (const item of measured) if (rejected.has(item) || rejectedRows.has(item.row)) item.p.style.removeProperty('height');
}
