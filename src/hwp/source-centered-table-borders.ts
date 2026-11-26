import type { HwpDocument } from '@rhwp/core';
import { sourceCellBoundsTolerance, type HwpPaintedLine } from './paragraph-layout';

type Cell = { index: number; row: number; height: number; rowSpan: number; isHeader: boolean; textDirection: number;
  diagonalSlash: number; diagonalBackSlash: number; paddingTop: number; paddingBottom: number;
  borderTop: { type: number }; borderBottom: { type: number } };
type Box = { cellIdx: number; pageIndex: number; row: number; y: number; h: number };

/** Intact rows can contain several paragraphs and still measure from
 * stroke centre to stroke centre. Require the complete native row grid and
 * visible text inside the authored insets on every source sheet. */
export function sourceCentersMultilineRowBorders(doc: HwpDocument, section: number, paragraph: number, control: number,
  table: Record<string, unknown>, cells: Cell[], boxes: Box[], singleColumn: (page: number) => boolean,
  painted: ReadonlyMap<string, HwpPaintedLine[]>) {
  if (table.treatAsChar || Number(table.pageBreak) !== 1 || !table.repeatHeader || table.hasCaption ||
      table.cellSpacing !== 0 || !table.restrictInPage || table.allowOverlap || table.textWrap !== 'TopAndBottom' ||
      table.vertRelTo !== 'Para' || table.vertAlign !== 'Top' || table.vertOffset !== 0 || table.horzRelTo !== 'Column' ||
      cells.some(c => c.rowSpan < 1 || c.textDirection || c.diagonalSlash || c.diagonalBackSlash ||
        ![0, 1].includes(c.borderTop?.type) || ![0, 1].includes(c.borderBottom?.type))) return false;
  const close = (a: number, b: number) => Math.abs(a - b) <= sourceCellBoundsTolerance;
  const placements = new Map<number, Box[]>(), pages = new Set<number>();
  for (const box of boxes) {
    const positions = placements.get(box.cellIdx) ?? [];
    positions.push(box);
    placements.set(box.cellIdx, positions);
    pages.add(box.pageIndex);
  }
  if (pages.size < 3 || placements.size !== cells.length || ![...pages].every(singleColumn)) return false;
  type Row = { box: Box; stored: number; tight: boolean };
  const rows = new Map<number, Map<number, Row>>();
  const marginsFit = (top: number, bottom: number) =>
    top >= -sourceCellBoundsTolerance && bottom >= -sourceCellBoundsTolerance;
  let multiline = false;
  for (const cell of cells) {
    const positions = placements.get(cell.index);
    if (!positions || cell.isHeader && (cell.row !== 0 || cell.rowSpan !== 1 ||
        new Set(positions.map(box => box.pageIndex)).size !== pages.size) || !cell.isHeader && positions.length !== 1)
      return false;
    const count = doc.getCellParagraphCount(section, paragraph, control, cell.index);
    if (!count || count > 64) return false;
    const ink: HwpPaintedLine[] = [];
    for (let index = 0; index < count; index++) {
      const address = JSON.stringify([{ controlIndex: control, cellIndex: cell.index, cellParaIndex: index }]);
      const { stored, controlCount, lines, tables } = JSON.parse(doc.getParagraphLineMetrics(section, paragraph, address));
      if (!stored || tables?.length || !lines?.length) return false;
      const physical = painted.get(`${section}:${paragraph}:${control}:${cell.index}:${index}`) ?? [];
      // A control with no measured advance can share the native line.
      if (controlCount && !(controlCount === 1 && !cell.isHeader && positions.length === 1 &&
          lines.length === 1 && physical.length === 1 && physical[0].charStart === 0 &&
          close(physical[0].h, lines[0].textHeight))) return false;
      if (!physical.length && lines.length === 1 &&
          !doc.getTextInCellByPath(section, paragraph, address, 0, 0xffffffff).trim()) continue;
      if (!positions.every(box => physical.filter(line => line.page === box.pageIndex).length === lines.length) ||
          physical.length !== positions.length * lines.length) return false;
      if (!cell.isHeader && lines.length > 1) multiline = true;
      ink.push(...physical);
    }
    for (const box of positions) {
      if (box.row !== cell.row || !(box.h > 0) || cell.height / 75 > box.h + sourceCellBoundsTolerance) return false;
      const bands = ink.filter(line => line.page === box.pageIndex);
      const top = Math.min(...bands.map(line => line.y)) - box.y - cell.paddingTop / 75;
      const bottom = box.y + box.h - Math.max(...bands.map(line => line.y + line.h)) - cell.paddingBottom / 75;
      if (!marginsFit(top, bottom) || bands.some(line => !(line.h > 0))) return false;
      if (cell.rowSpan > 1) continue;
      const byPage = rows.get(box.pageIndex) ?? new Map<number, Row>();
      const row = byPage.get(cell.row) ?? { box, stored: 0, tight: false };
      if (!close(row.box.y, box.y) || !close(row.box.h, box.h)) return false;
      row.stored = Math.max(row.stored, cell.height / 75);
      row.tight ||= !cell.isHeader && bands.length > 0 && close(top, 0) && close(bottom, 0) &&
        cell.borderTop.type === 1 && cell.borderBottom.type === 1;
      byPage.set(cell.row, row);
      rows.set(box.pageIndex, byPage);
    }
  }
  if (!multiline || rows.size !== pages.size) return false;
  for (const page of rows.values()) {
    const ordered = [...page.values()].sort((a, b) => a.box.y - b.box.y);
    if (ordered[0]?.box.row !== 0 || !ordered.some(row => row.tight) ||
        ordered.some((row, index) => !row.tight && !close(row.stored, row.box.h) ||
          index > 0 && !close(ordered[index - 1].box.y + ordered[index - 1].box.h, row.box.y))) return false;
  }
  for (const cell of cells.filter(cell => cell.rowSpan > 1)) {
    const box = placements.get(cell.index)![0], page = rows.get(box.pageIndex)!;
    const first = page.get(cell.row)?.box, last = page.get(cell.row + cell.rowSpan - 1)?.box;
    if (!first || !last || !close(first.y, box.y) || !close(last.y + last.h, box.y + box.h)) return false;
  }
  return true;
}

/** The source may have centered logical cells whose text starts at the
 * authored top inset on every page of a split. Confirm the shared border
 * convention against painted text before using it on the continuing grid. */
export function sourceCentersContinuedBorders(doc: HwpDocument, section: number, paragraph: number, control: number,
  cells: Cell[], boxes: Box[], painted: ReadonlyMap<string, HwpPaintedLine[]>) {
  if (cells.some(cell => cell.textDirection || cell.diagonalSlash || cell.diagonalBackSlash ||
      ![0, 1].includes(cell.borderTop?.type) || ![0, 1].includes(cell.borderBottom?.type))) return false;
  const fragments = new Map<number, Box[]>();
  for (const box of boxes) {
    const entries = fragments.get(box.cellIdx) ?? [];
    entries.push(box);
    fragments.set(box.cellIdx, entries);
  }
  let found = false;
  for (const cell of cells) {
    const placements = fragments.get(cell.index);
    if (!placements || cell.rowSpan !== 1 || cell.isHeader || placements.length < 2) continue;
    const byPage = new Map<number, number>();
    const count = doc.getCellParagraphCount(section, paragraph, control, cell.index);
    for (let index = 0; index < count; index++) {
      for (const line of painted.get(`${section}:${paragraph}:${control}:${cell.index}:${index}`) ?? []) {
        const box = placements.find(part => part.pageIndex === line.page);
        if (!box || !(line.h > 0 && line.y >= box.y - .16 && line.y + line.h <= box.y + box.h + .16)) return false;
        byPage.set(line.page, Math.min(byPage.get(line.page) ?? Infinity, line.y));
      }
    }
    if (byPage.size < 2) continue;
    if (!(cell.paddingTop > 0 && cell.borderTop.type === 1 &&
      placements.every(part => byPage.has(part.pageIndex) &&
        Math.abs(byPage.get(part.pageIndex)! - part.y - cell.paddingTop / 75) <= .16))) return false;
    found = true;
  }
  return found;
}

/** An intact inline grid can have merged cells. For each row, require a
 * terminal line whose em and native insets already fill its stored height. */
function sourceCentersInlineBorders(doc: HwpDocument, section: number, paragraph: number, control: number,
  cells: Cell[], boxes: Box[], pages: number[], table: Record<string, unknown>,
  painted: ReadonlyMap<string, HwpPaintedLine[]>) {
  if (pages.length !== 1 || boxes.length !== cells.length ||
      cells.some(cell => ![0, 1, 8].includes(cell.borderTop?.type) || ![0, 1, 8].includes(cell.borderBottom?.type))) return false;
  const page = JSON.parse(doc.getPageInfo(pages[0]));
  const top = Number(page.marginTop) + Number(page.marginHeader);
  const bottom = Number(page.footerArea?.y ?? (page.height - page.marginBottom - page.marginFooter));
  const byCell = new Map(boxes.map(box => [box.cellIdx, box]));
  if (byCell.size !== cells.length) return false;
  const rows = new Set<number>();
  for (const cell of cells) {
    const box = byCell.get(cell.index);
    const height = cell.height / 75;
    if (!box || box.row !== cell.row || !(height > 0) || Math.abs(height - box.h) > .1 ||
        !(box.y >= top - .1 && box.y + box.h <= bottom + .1)) return false;
    rows.add(cell.row);
  }
  if (rows.size < 2) return false;
  const tight = new Set<number>();
  for (const cell of cells) {
    if (tight.has(cell.row) || cell.rowSpan !== 1 || !(cell.borderTop.type || cell.borderBottom.type) ||
        doc.getCellParagraphCount(section, paragraph, control, cell.index) !== 1) continue;
    const path = JSON.stringify([{ controlIndex: control, cellIndex: cell.index, cellParaIndex: 0 }]);
    const { stored, controlCount, lines } = JSON.parse(doc.getParagraphLineMetrics(section, paragraph, path));
    if (stored && !controlCount && lines?.length === 1 && lines[0].textHeight > 0 &&
        Math.abs(cell.height - (lines[0].textHeight * 75 + cell.paddingTop + cell.paddingBottom)) <= 1)
      tight.add(cell.row);
  }
  if (tight.size === rows.size) return true;
  // A headerless grid can also share stroke centres when a complete empty
  // separator spans its rows and the native text uses their authored insets.
  if (table.horzRelTo !== 'Para' || table.vertOffset !== 0 || table.pageBreak !== 2 ||
      cells.some(cell => cell.isHeader) || rows.size < 3) return false;
  const separators = cells.filter(cell => cell.rowSpan > 1);
  if (!separators.length || separators.some(cell => cell.row !== 0 || cell.rowSpan !== rows.size ||
      cell.borderTop.type || cell.borderBottom.type)) return false;
  for (const cell of separators) {
    const path = JSON.stringify([{ controlIndex: control, cellIndex: cell.index, cellParaIndex: 0 }]);
    const { stored, controlCount, lines } = JSON.parse(doc.getParagraphLineMetrics(section, paragraph, path));
    if (doc.getCellParagraphCount(section, paragraph, control, cell.index) !== 1 ||
        doc.getTextInCellByPath(section, paragraph, path, 0, 0xffffffff) !== '' ||
        !stored || controlCount || lines?.length !== 1) return false;
  }
  const matched = new Set<number>();
  for (const cell of cells) {
    if (cell.rowSpan > 1) continue;
    const box = byCell.get(cell.index)!;
    if (cell.borderTop.type !== 1 || cell.borderBottom.type !== 1 ||
        !(cell.paddingTop > 0 && cell.paddingBottom > 0)) return false;
    let first = Infinity, end = -Infinity;
    const count = doc.getCellParagraphCount(section, paragraph, control, cell.index);
    if (!(count > 0 && count <= 32)) return false;
    for (let index = 0; index < count; index++) {
      const path = JSON.stringify([{ controlIndex: control, cellIndex: cell.index, cellParaIndex: index }]);
      const { stored, controlCount, lines } = JSON.parse(doc.getParagraphLineMetrics(section, paragraph, path));
      const ink = painted.get(`${section}:${paragraph}:${control}:${cell.index}:${index}`);
      if (!stored || controlCount || !lines?.length) return false;
      // The paint map omits authored blank paragraphs. Their exact native
      // row is still checked above and the editor keeps the blank's pitch.
      if (!ink && lines.length === 1 && !doc.getTextInCellByPath(section, paragraph, path, 0, 0xffffffff).trim()) continue;
      if (!ink || ink.length !== lines.length) return false;
      for (const line of ink) {
        if (line.page !== box.pageIndex || !(line.h > 0)) return false;
        first = Math.min(first, line.y);
        end = Math.max(end, line.y + line.h);
      }
    }
    const topGap = first - box.y - cell.paddingTop / 75;
    const bottomGap = box.y + box.h - end - cell.paddingBottom / 75;
    if (!(topGap >= -.12 && bottomGap >= -.12)) return false;
    if (Math.abs(topGap) <= .12 && Math.abs(bottomGap) <= .12) matched.add(cell.row);
  }
  return matched.has(0) && [...matched].some(row => row > 0);
}

/** A row-breaking HWP table can measure its grid between stroke centers.
 * Prove that its authored minima and terminal line extents alone account for
 * the source heights before using that convention in our collapsed grid. */
export function sourceCentersTableBorders(doc: HwpDocument, section: number, paragraph: number, control: number,
  table: Record<string, unknown>, cells: Cell[], boxes: Box[], singleColumn: (page: number) => boolean,
  painted: ReadonlyMap<string, HwpPaintedLine[]>): boolean | 'inline' {
  if (!table.repeatHeader || table.hasCaption || table.cellSpacing !== 0 ||
    table.vertRelTo !== 'Para' || table.vertAlign !== 'Top' || table.textWrap !== 'TopAndBottom') return false;
  const pages = [...new Set(boxes.map(box => box.pageIndex))];
  if (!pages.every(singleColumn) || cells.some(cell => cell.textDirection || cell.diagonalSlash || cell.diagonalBackSlash)) return false;
  if (table.treatAsChar) return sourceCentersInlineBorders(doc, section, paragraph, control, cells, boxes, pages, table, painted)
    ? 'inline' : false;
  if (Number(table.pageBreak) !== 1 || pages.length < 2 || cells.some(cell => cell.rowSpan < 1 ||
    ![0, 1].includes(cell.borderTop?.type) || ![0, 1].includes(cell.borderBottom?.type))) return false;
  if (!cells.some(cell => cell.borderTop.type || cell.borderBottom.type)) return false;
  const heights = new Map<number, number>();
  const spans = cells.filter(cell => cell.rowSpan > 1);
  const controlLines = new Map<number, { line: HwpPaintedLine; padding: number }>();
  for (const cell of cells) {
    if (cell.rowSpan > 1) continue;
    if (doc.getCellParagraphCount(section, paragraph, control, cell.index) !== 1) return false;
    const path = JSON.stringify([{ controlIndex: control, cellIndex: cell.index, cellParaIndex: 0 }]);
    const { stored, controlCount, lines, tables } = JSON.parse(doc.getParagraphLineMetrics(section, paragraph, path));
    if (!stored || lines?.length !== 1 || !(lines[0].textHeight > 0)) return false;
    if (controlCount) {
      // A zero-advance control such as a bookmark can share a proven line.
      // Require independent paint for its entire single-line height.
      const ink = painted.get(`${section}:${paragraph}:${control}:${cell.index}:0`);
      if (!spans.length || cell.isHeader || controlCount !== 1 || tables?.length !== 0 ||
          ink?.length !== 1 || ink[0].charStart !== 0 || Math.abs(ink[0].h - lines[0].textHeight) > .16) return false;
      controlLines.set(cell.index, { line: ink[0], padding: cell.paddingTop / 75 });
    }
    const textHeight = lines[0].textHeight + (cell.paddingTop + cell.paddingBottom) / 75;
    heights.set(cell.row, Math.max(heights.get(cell.row) ?? 0, cell.height / 75, textHeight));
  }
  if (new Set(boxes.map(box => box.cellIdx)).size !== cells.length) return false;
  if (!spans.length) return boxes.every(box => Math.abs(box.h - (heights.get(box.row) ?? Infinity)) <= .1);

  // A vertical merge does not add a stroke inside each of its logical rows.
  // Require a complete, unbroken native span over independently proven rows.
  const byCell = new Map<number, Box[]>();
  const rowsByPage = new Map<number, Map<number, Box>>();
  const merged = new Set(spans.map(cell => cell.index));
  for (const box of boxes) {
    if (merged.has(box.cellIdx)) {
      const owned = byCell.get(box.cellIdx) ?? [];
      owned.push(box);
      byCell.set(box.cellIdx, owned);
      continue;
    }
    const control = controlLines.get(box.cellIdx);
    if (control && (control.line.page !== box.pageIndex || Math.abs(control.line.y - box.y - control.padding) > .16)) return false;
    if (Math.abs(box.h - (heights.get(box.row) ?? Infinity)) > .1) return false;
    const rows = rowsByPage.get(box.pageIndex) ?? new Map<number, Box>();
    const prior = rows.get(box.row);
    if (prior && Math.abs(prior.y - box.y) > .1) return false;
    rows.set(box.row, box);
    rowsByPage.set(box.pageIndex, rows);
  }
  for (const cell of spans) {
    const owned = byCell.get(cell.index);
    if (cell.isHeader || owned?.length !== 1 || doc.getCellParagraphCount(section, paragraph, control, cell.index) !== 1) return false;
    const [box] = owned;
    const rows = rowsByPage.get(box.pageIndex);
    let end = box.y;
    for (let row = cell.row; row < cell.row + cell.rowSpan; row++) {
      const band = rows?.get(row);
      if (!band || Math.abs(band.y - end) > .16) return false;
      end = band.y + band.h;
    }
    if (Math.abs(end - box.y - box.h) > .16 || cell.height / 75 > box.h + .1) return false;
    const path = JSON.stringify([{ controlIndex: control, cellIndex: cell.index, cellParaIndex: 0 }]);
    const { stored, controlCount, lines } = JSON.parse(doc.getParagraphLineMetrics(section, paragraph, path));
    if (!stored || controlCount || !lines?.length || lines.some((line: { textHeight: number }) => !(line.textHeight > 0))) return false;
    const ink = painted.get(`${section}:${paragraph}:${control}:${cell.index}:0`);
    if (!ink && !doc.getTextInCellByPath(section, paragraph, path, 0, 0xffffffff).trim() &&
        lines.length === 1 && lines[0].textHeight + (cell.paddingTop + cell.paddingBottom) / 75 <= box.h + .1) continue;
    if (!ink || ink.length !== lines.length || ink.some(line => line.page !== box.pageIndex || !(line.h > 0) ||
        line.y < box.y + cell.paddingTop / 75 - .16 || line.y + line.h > end - cell.paddingBottom / 75 + .16)) return false;
  }
  return true;
}
