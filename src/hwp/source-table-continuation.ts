import type { HwpDocument } from '@rhwp/core';
import { hwpOverhangCellFingerprint, type HwpTableContinuation } from '../hwp-table-continuation';
import type { HwpIntactRow } from '../hwp-intact-row';
import { sourceCellBoundsTolerance, type HwpPaintedLine } from './paragraph-layout';
import { hwpWrapFingerprint, normalizeHwpWrapText } from '../hwp-text-advances';

type Cell = { index: number; row: number; col: number; rowSpan: number; colSpan: number; isHeader?: boolean };
type CellBox = { cellIdx: number; pageIndex: number; row: number; y: number; h: number };
type TextRun = {
  text: string; x: number; y: number; w: number; h: number; secIdx: number; paraIdx: number; parentParaIdx: number; headerFooter?: boolean;
  cellPath?: { controlIndex: number; cellIndex: number; cellParaIndex: number }[];
};

export function sourceTableContinuation(
  doc: HwpDocument, section: number, paragraph: number, controlIndex: number,
  rows: number, cells: Cell[], boxes: CellBox[],
): HwpTableContinuation | undefined {
  if (!boxes.length || !cells.some(cell => cell.row === 0 && cell.isHeader)) return;
  try {
    let first = Infinity, last = 0;
    for (const box of boxes) { first = Math.min(first, box.pageIndex); last = Math.max(last, box.pageIndex); }
    if (!(first < last)) return;
    const page = JSON.parse(doc.getPageInfo(first));
    if (page.columns?.length !== 1) return;
    const top = Number(page.marginTop) + Number(page.marginHeader);
    const bottom = Number(page.footerArea?.y ?? (page.height - page.marginBottom - page.marginFooter));
    const width = Number(page.bodyRight) - Number(page.bodyLeft);
    const firstBoxes = boxes.filter(box => box.pageIndex === first);
    const header = firstBoxes.filter(box => box.row === 0);
    if (!header.length || !(width > 0 && bottom > top)) return;
    const tableTop = Math.min(...header.map(box => box.y));
    if (tableTop < top - .1 || tableTop >= bottom) return;
    const runs: TextRun[] = JSON.parse(doc.getPageTextLayout(first)).runs;
    const painted = runs.filter(run => run.secIdx === section && run.parentParaIdx === paragraph &&
      !run.headerFooter && run.cellPath?.length === 1 && run.cellPath[0].controlIndex === controlIndex && !!run.text);
    if (!painted.length || painted.some(run => !(run.y >= top - .1 && run.y + run.h <= bottom + .1))) return;
    const cellById = new Map(cells.map(cell => [cell.index, cell]));
    const row = Math.max(...painted.map(run => cellById.get(run.cellPath![0].cellIndex)?.row ?? -1));
    if (row < 1 || row + 1 >= rows) return;
    const boundary = cells.filter(cell => cell.row === row).sort((a, b) => a.col - b.col);
    if (boundary.length < 2 || boundary.length > 6 || boundary.some((cell, col) =>
      cell.col !== col || cell.isHeader || cell.rowSpan !== 1 || cell.colSpan !== 1 ||
      !firstBoxes.some(box => box.cellIdx === cell.index) ||
      cells.some(other => other.row < row && other.row + other.rowSpan > row))) return;
    const fingerprints: string[] = [];
    for (const cell of boundary) {
      const count = doc.getCellParagraphCount(section, paragraph, controlIndex, cell.index);
      if (count < 1 || count > 12) return;
      let cellText = '';
      for (let cellParaIndex = 0; cellParaIndex < count; cellParaIndex++) {
        const path = JSON.stringify([{ controlIndex, cellIndex: cell.index, cellParaIndex }]);
        const literal = doc.getTextInCellByPath(section, paragraph, path, 0, 0xffffffff);
        if (!literal.trim() || literal.length > 2000 || /[\t\r\n\uD800-\uDFFF\p{Bidi_Control}\ufffc]/u.test(literal)) return;
        const metric = JSON.parse(doc.getParagraphLineMetrics(section, paragraph, path));
        if (!metric.stored || metric.controlCount || !metric.lines?.length) return;
        const parts = painted.filter(run => run.cellPath![0].cellIndex === cell.index &&
          run.cellPath![0].cellParaIndex === cellParaIndex);
        if (parts.map(run => run.text).join('') !== literal) return;
        cellText += literal;
      }
      fingerprints.push(hwpWrapFingerprint(cellText));
    }
    return { row, rows, bodyWidth: width, bodyHeight: bottom - top, tableTop: Math.max(0, tableTop - top), cells: fingerprints };
  } catch { /* Skip originals without complete native layout evidence. */ }
}

/** Some HWP tables use the empty area immediately below the first body.
 * Preserve only a complete, observed first cut of two intact text cells. */
export function sourceFirstTableOverhang(doc: HwpDocument, section: number, paragraph: number, control: number,
  rows: number, cells: readonly Cell[], boxes: readonly (CellBox & { x: number; w: number })[]): HwpTableContinuation | undefined {
  try {
    const first = Math.min(...boxes.map(box => box.pageIndex));
    if (!Number.isFinite(first) || !cells.some(cell => cell.row === 0 && cell.isHeader)) return;
    const paper = JSON.parse(doc.getPageInfo(first)), tolerance = sourceCellBoundsTolerance;
    if (paper.columns?.length !== 1) return;
    const top = Number(paper.marginTop) + Number(paper.marginHeader), bottom = Number(paper.footerArea?.y);
    const width = Number(paper.bodyRight) - Number(paper.bodyLeft);
    if (!(width > 0 && bottom > top && bottom < paper.height)) return;
    const placed = boxes.filter(box => box.pageIndex === first);
    const tableTop = Math.min(...placed.map(box => box.y)), tableEnd = Math.max(...placed.map(box => box.y + box.h));
    if (tableTop < top - tolerance || tableTop >= bottom || tableEnd <= bottom + tolerance || tableEnd >= paper.height) return;
    const own = (run: TextRun) => run.secIdx === section && run.parentParaIdx === paragraph && !run.headerFooter &&
      run.cellPath?.length === 1 && run.cellPath[0].controlIndex === control;
    const runs: TextRun[] = JSON.parse(doc.getPageTextLayout(first)).runs;
    const paint = runs.filter(run => run.text.trim() && own(run));
    const footerPlaceholder = (run: TextRun) => run.paraIdx === 0xffffffff && /^[\s–—-]+$/u.test(run.text);
    if (runs.some(run => run.text.trim() && !own(run) && !footerPlaceholder(run))) return;
    const shapes = JSON.parse(doc.getPageControlLayout(first)).controls;
    if (shapes.length !== 1 || shapes[0].type !== 'table' || shapes[0].secIdx !== section ||
        shapes[0].paraIdx !== paragraph || shapes[0].controlIdx !== control) return;
    const overflow = paint.filter(run => run.y + run.h > bottom + tolerance);
    const overflowCells = new Set(overflow.map(run => run.cellPath![0].cellIndex));
    const row = Math.max(...placed.map(box => box.row));
    if (row < 1 || row + 1 >= rows || !overflow.length || overflow.some(run =>
      run.y < bottom - tolerance || run.y + run.h > tableEnd + tolerance ||
      run.x < paper.bodyLeft - tolerance || run.x + run.w > paper.bodyRight + tolerance)) return;
    const boundary = cells.filter(cell => cell.row === row).sort((a,b) => a.col - b.col);
    if (boundary.length < 2 || boundary.length > 6 || boundary.length !== overflowCells.size ||
        cells.some(cell => cell.row < row && cell.row + cell.rowSpan > row)) return;
    const fingerprints: string[] = [], heads: { end: number; fingerprint: string }[] = [];
    for (const [i, cell] of boundary.entries()) {
      if (cell.isHeader || cell.col !== i || cell.rowSpan !== 1 || cell.colSpan !== 1 || !overflowCells.has(cell.index)) return;
      const parts = boxes.filter(box => box.cellIdx === cell.index).sort((a,b) => a.pageIndex - b.pageIndex);
      if (parts.length !== 2 || parts[0].pageIndex !== first || parts[1].pageIndex !== first + 1 ||
          Math.abs(parts[0].y + parts[0].h - tableEnd) > tolerance ||
          Math.abs(parts[0].w - parts[1].w) > tolerance || parts[0].y >= bottom) return;
      const painted = overflow.filter(run => run.cellPath![0].cellIndex === cell.index);
      if (painted.some(run => run.cellPath![0].cellParaIndex !== 0 || Math.abs(run.y - painted[0].y) > tolerance)) return;
      const count = doc.getCellParagraphCount(section, paragraph, control, cell.index), literals: string[] = [];
      if (count < 1 || count > 16) return;
      for (let cellParaIndex = 0; cellParaIndex < count; cellParaIndex++) {
        const address = JSON.stringify([{ controlIndex: control, cellIndex: cell.index, cellParaIndex }]);
        const literal = doc.getTextInCellByPath(section, paragraph, address, 0, 0xffffffff);
        const saved = JSON.parse(doc.getParagraphLineMetrics(section, paragraph, address));
        if (!saved.stored || saved.controlCount || !saved.lines?.length || literal.length > 2000 ||
            /[\t\r\uD800-\uDFFF\p{Bidi_Control}\ufffc]/u.test(literal)) return;
        if (!cellParaIndex) {
          const prefix = normalizeHwpWrapText(painted.map(run => run.text).join(''));
          const full = normalizeHwpWrapText(literal.split('\n')[0]);
          if (saved.lines.length < 2 || !prefix || prefix.length >= full.length || !full.startsWith(prefix)) return;
          const first = paint.filter(run => run.cellPath![0].cellIndex === cell.index);
          if (first.length === painted.length) heads.push({ end: prefix.length, fingerprint: hwpWrapFingerprint(prefix) });
        }
        literals.push(literal);
      }
      fingerprints.push(hwpOverhangCellFingerprint(literals));
    }
    // The first physical fragment starts at the implicit object inset,
    // without an author-specified indent in its empty host.
    const style = JSON.parse(doc.getTableProperties(section, paragraph, control));
    const host = JSON.parse(doc.getParaPropertiesAt(section, paragraph));
    const defaultHostOrigin = heads.length === fingerprints.length && !style.treatAsChar && !style.hasCaption &&
      style.horzRelTo === 'Column' && style.horzAlign === 'Left' && style.horzOffset === 0 && style.outerLeft === 283 &&
      style.textWrap === 'TopAndBottom' && style.restrictInPage && !style.allowOverlap &&
      Math.abs(shapes[0].x - paper.bodyLeft - style.outerLeft / 75) <= tolerance &&
      host.marginLeft === 0 && host.marginRight === 0 && host.indent === 0 &&
      doc.getParagraphLength(section, paragraph) === 0 &&
      JSON.parse(doc.getParagraphLineMetrics(section, paragraph, '[]')).controlCount === 1;
    return { row, rows, bodyWidth: width, bodyHeight: bottom - top, tableTop: Math.max(0, tableTop - top),
      cells: fingerprints, overhang: { page: first + 1, end: tableEnd - top,
        ...(defaultHostOrigin ? { defaultHostOrigin: true } : {}),
        ...(heads.length === fingerprints.length ? { heads } : {}) } };
  } catch { /* Without a complete source witness, let the table paginate normally. */ }
}

/** Repeated-header flow can carry a complete ordinary row to the next sheet.
 * Verify both the native cell edges and every authored paragraph before
 * recording the placement. The stored minimum and permission are unchanged. */
export function sourceIntactTopRow(
  doc: HwpDocument, section: number, paragraph: number, controlIndex: number,
  cells: readonly (Cell & { height: number; verticalAlign: number; textDirection: number;
    diagonalSlash: number; diagonalBackSlash: number })[],
  boxes: readonly (CellBox & { w: number })[], minimum: number,
  lines: Map<string, HwpPaintedLine[]>,
  body: { columns: number | undefined; top: number; bottom: number },
): HwpIntactRow | undefined {
  const first = boxes[0], tolerance = sourceCellBoundsTolerance;
  if (!first || cells.length < 2 || cells.length > 6 || cells.length !== boxes.length || body.columns !== 1 ||
      first.y < body.top - tolerance || first.y + first.h > body.bottom + tolerance ||
      !(first.h > 0 && first.h <= 2000 && Math.abs(minimum - first.h) <= .1) ||
      Math.abs(first.h - Math.max(...cells.map(cell => cell.height / 75))) > .1) return;
  const widths: number[] = [], text: string[][] = [];
  for (const [index, cell] of cells.entries()) {
    if (cell.col !== index || cell.isHeader || cell.colSpan !== 1 || cell.rowSpan !== 1 || cell.verticalAlign !== 0 ||
        cell.textDirection || cell.diagonalSlash || cell.diagonalBackSlash) return;
    const box = boxes.find(box => box.cellIdx === cell.index);
    if (!box || box.pageIndex !== first.pageIndex || Math.abs(box.y - first.y) > .1 || Math.abs(box.h - first.h) > .1 ||
        !(box.w > 0 && box.w <= 10000)) return;
    const paragraphs: string[] = [];
    try {
      const count = doc.getCellParagraphCount(section, paragraph, controlIndex, cell.index);
      if (count < 1 || count > 16) return;
      for (let p = 0; p < count; p++) {
        const path = JSON.stringify([{ controlIndex, cellIndex: cell.index, cellParaIndex: p }]);
        const literal = doc.getTextInCellByPath(section, paragraph, path, 0, 0xffffffff);
        if (literal.length > 2000 || /[\x00-\x08\x0a-\x1f\ufffc]/u.test(literal)) return;
        const metric = JSON.parse(doc.getParagraphLineMetrics(section, paragraph, path));
        if (!metric.stored || metric.controlCount || !metric.lines?.length) return;
        const painted = lines.get(`${section}:${paragraph}:${controlIndex}:${cell.index}:${p}`);
        if (literal.trim() && (!painted?.length || painted.some(line => line.page !== first.pageIndex ||
            line.y < box.y - tolerance || line.y + line.h > box.y + box.h + tolerance))) return;
        paragraphs.push(literal);
      }
    } catch { return; }
    widths.push(box.w); text.push(paragraphs);
  }
  return { height: first.h, widths, fingerprint: hwpWrapFingerprint(JSON.stringify(text)) };
}

type PlainCell = Cell & { textDirection: number; diagonalSlash: number; diagonalBackSlash: number };
type TableBox = CellBox & { w: number };
type BodyBounds = { columns: number | undefined; top: number; bottom: number };

/** Repeated headers in this continuation model start on the body edge. */
export function sourceAlignedHeaderEnds(cells: readonly Cell[], boxes: readonly CellBox[], body: (page: number) => BodyBounds) {
  const ends = new Map<number, number>();
  const headers = cells.filter(cell => cell.isHeader);
  if (headers.length < 2 || headers.some(cell => cell.row !== 0 || cell.rowSpan !== 1)) return ends;
  const ids = new Set(headers.map(cell => cell.index));
  const byPage = new Map<number, CellBox[]>();
  for (const box of boxes) if (ids.has(box.cellIdx)) {
    const parts = byPage.get(box.pageIndex) ?? [];
    parts.push(box);
    byPage.set(box.pageIndex, parts);
  }
  for (const [page, parts] of byPage) {
    const [first] = parts, paper = body(page);
    if (paper.columns === 1 && parts.length === ids.size && new Set(parts.map(box => box.cellIdx)).size === ids.size &&
        Math.abs(first.y - paper.top) <= sourceCellBoundsTolerance && first.h > 0 && first.y + first.h < paper.bottom &&
        parts.every(box => Math.abs(box.y - first.y) <= .1 && Math.abs(box.h - first.h) <= .1))
      ends.set(page, first.y + first.h);
  }
  return ends;
}

/** Read the physical minimum only from two equal ordinary cell fragments
 * inside an already witnessed page-spanning merge. */
export function sourceContinuedRowReserve(
  cells: readonly (PlainCell & { height: number; paddingBottom: number })[],
  boxes: readonly TableBox[], height: number, body: (page: number) => BodyBounds,
  afterHeader: (page: number) => number | undefined,
) {
  const ordinary = cells.filter(cell => cell.rowSpan === 1);
  if (ordinary.length < 3 || !(height > 0 && height <= 2000) ||
      ordinary.some(cell => cell.isHeader || cell.colSpan !== 1 || cell.textDirection ||
        cell.diagonalBackSlash || cell.diagonalSlash || Math.abs(cell.height / 75 - height) > .12)) return;
  const segments = ordinary.map(cell => boxes.filter(box => box.cellIdx === cell.index)
    .sort((a, b) => a.pageIndex - b.pageIndex));
  if (segments.some(parts => parts.length !== 2)) return;
  const [head, tail] = segments[0];
  const first = body(head.pageIndex), second = body(tail.pageIndex);
  const topBelowHeader = afterHeader(tail.pageIndex);
  const pad = Math.max(...ordinary.map(cell => cell.paddingBottom / 75));
  const rightAtBottom = first.bottom - head.y - head.h;
  const tolerance = sourceCellBoundsTolerance;
  if (tail.pageIndex !== head.pageIndex + 1 || first.columns !== 1 || second.columns !== 1 ||
      !(head.h > 0 && tail.h > 0) || head.y < first.top - tolerance ||
      rightAtBottom < -tolerance || rightAtBottom > pad + tolerance ||
      topBelowHeader === undefined || Math.abs(tail.y - topBelowHeader) > tolerance ||
      tail.y + tail.h > second.bottom + tolerance) return;
  if (segments.some(parts => parts.some((box, index) => {
    const native = index ? tail : head;
    return box.pageIndex !== native.pageIndex || !Number.isFinite(box.w) || box.w <= 0 ||
      Math.abs(box.y - native.y) > tolerance || Math.abs(box.h - native.h) > tolerance ||
      Math.abs(box.w - parts[0].w) > tolerance;
  }))) return;
  const observed = Math.round((head.h + tail.h) * 100) / 100;
  if (observed < height - .5) return { observed, headHeight: head.h };
}

/** A blank vertical merge keeps the cuts established by its ordinary
 * neighbours, including a long first row that continues across sheets. */
export function sourceEmptyMergeCuts(
  doc: HwpDocument, address: [number, number, number], owner: PlainCell,
  rows: readonly (readonly PlainCell[])[], boxes: readonly TableBox[],
  body: (page: number) => BodyBounds, painted: Map<string, HwpPaintedLine[]>,
): number[] {
  const [section, paragraph, controlIndex] = address;
  const tolerance = sourceCellBoundsTolerance;
  const plain = (cell: PlainCell) => !cell.isHeader && cell.colSpan === 1 && !cell.textDirection &&
    !cell.diagonalSlash && !cell.diagonalBackSlash;
  if (!plain(owner) || rows.length !== owner.rowSpan || rows.some(row =>
    row.length < 1 || row.some(cell => !plain(cell) || cell.rowSpan !== 1))) return [];
  const close = (a: number, b: number) => Math.abs(a - b) <= tolerance;
  const withinBody = ({ pageIndex, y, h, w }: TableBox) => {
    const { columns, top, bottom } = body(pageIndex);
    return columns === 1 && w > 0 && h > 0 && y >= top - tolerance && y + h <= bottom + tolerance;
  };
  const span = boxes.filter(box => box.cellIdx === owner.index).sort((a, b) => a.pageIndex - b.pageIndex);
  if (span.length < 3 || !span.every((box, i) =>
    box.pageIndex === span[0].pageIndex + i && close(box.w, span[0].w) && withinBody(box))) return [];
  const partsFor = (cell: PlainCell) => boxes.filter(box => box.cellIdx === cell.index)
    .sort((a, b) => a.pageIndex - b.pageIndex);

  const proof = (cell: PlainCell, parts: TableBox[], wholeParagraphs = false) => {
    const count = doc.getCellParagraphCount(...address, cell.index);
    const pages = new Set<number>();
    for (let p = 0; p < count; p++) {
      const path = JSON.stringify([{ controlIndex, cellIndex: cell.index, cellParaIndex: p }]);
      const literal = doc.getTextInCellByPath(section, paragraph, path, 0, 0xffffffff);
      const metric = JSON.parse(doc.getParagraphLineMetrics(section, paragraph, path));
      const unsupportedText = wholeParagraphs ? /[\x00-\x09\x0b-\x1f\ufffc]/u : /[\x00-\x1f\ufffc]/u;
      if (!metric.stored || metric.controlCount || !metric.lines?.length || unsupportedText.test(literal)) return false;
      if (!literal.trim()) continue;
      const lines = painted.get(`${section}:${paragraph}:${controlIndex}:${cell.index}:${p}`);
      if (!lines || lines.length !== metric.lines.length ||
          (wholeParagraphs && lines.some(line => line.page !== lines[0].page))) return false;
      const characters = wholeParagraphs ? [...literal] : [];
      for (const [i, line] of lines.entries()) {
        const box = parts.find(part => part.pageIndex === line.page);
        const storedStart = metric.lines[i].textStart;
        // The paint map gathers the first visible glyph. A stored start may
        // include spaces before it when the complete paragraph travels intact.
        const visibleStart = line.charStart === storedStart || (wholeParagraphs && typeof line.charStart === 'number' &&
          Number.isInteger(storedStart) && storedStart >= 0 && line.charStart > storedStart && line.charStart < characters.length &&
          characters.slice(storedStart, line.charStart).every(char => char === ' '));
        if (!box || !visibleStart || !(line.h > 0 &&
            line.y >= box.y - tolerance && line.y + line.h <= box.y + box.h + tolerance)) return false;
        pages.add(line.page);
      }
    }
    return pages.size === parts.length;
  };

  const twoPageCut = (head: TableBox, tail: TableBox, afterRow: number, wholeParagraphs = false) =>
    rows.find(group => group[0].row > afterRow && group.some(cell => {
      const parts = partsFor(cell);
      if (parts.length !== 2 || parts[0].pageIndex !== head.pageIndex || parts[1].pageIndex !== tail.pageIndex ||
          !parts.every(withinBody) || !close(parts[0].w, parts[1].w) ||
          parts[0].y < head.y - tolerance || !close(parts[0].y + parts[0].h, head.y + head.h) ||
          !close(parts[1].y, tail.y) || parts[1].y + parts[1].h > tail.y + tail.h + tolerance) return false;
      return proof(cell, parts, wholeParagraphs);
    }));

  try {
    const ownerPath = JSON.stringify([{ controlIndex, cellIndex: owner.index, cellParaIndex: 0 }]);
    if (doc.getCellParagraphCount(...address, owner.index) !== 1 ||
        doc.getTextInCellByPath(section, paragraph, ownerPath, 0, 0xffffffff) !== '' ||
        JSON.parse(doc.getParagraphLineMetrics(section, paragraph, ownerPath)).controlCount !== 0) return [];
    if (span.length > 3) {
      const head = rows[0];
      for (const cell of head) {
        const parts = partsFor(cell);
        if (parts.length < 3 || parts.length >= span.length || !parts.every((part, i) =>
          part.pageIndex === span[i].pageIndex && close(part.w, parts[0].w) && close(part.y, span[i].y) &&
          withinBody(part) && (i < parts.length - 1 ? close(part.h, span[i].h) : part.h <= span[i].h + tolerance))) continue;
        if (!head.every(other => {
          const neighbours = partsFor(other);
          return neighbours.length === parts.length && neighbours.every((box, i) => box.pageIndex === parts[i].pageIndex &&
            close(box.y, parts[i].y) && close(box.h, parts[i].h) && withinBody(box));
        })) continue;
        if (!proof(cell, parts, true)) continue;
        const cuts = [cell.row];
        for (let page = parts.length - 1; page < span.length - 1; page++) {
          const row = twoPageCut(span[page], span[page + 1], cuts.at(-1)!, true);
          if (row) cuts.push(row[0].row);
        }
        return cuts;
      }
      return [];
    }
    const cuts: number[] = [];
    for (const [i, head] of span.slice(0, -1).entries()) {
      const tail = span[i + 1];
      const row = twoPageCut(head, tail, cuts.at(-1) ?? -1);
      if (!row) return [];
      cuts.push(row[0].row);
    }
    return cuts;
  } catch { return []; }
}
