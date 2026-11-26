import type { HwpDocument } from '@rhwp/core';
import {
  hwpNestedCellToken, hwpNestedLayoutFingerprint, hwpNestedParagraphToken, hwpNestedTableToken, parseHwpNestedTableLayout,
  type HwpNestedCellFace, type HwpNestedTableCut, type HwpNestedTableGap,
} from '../hwp-nested-table-layout';
import { normalizeHwpWrapText } from '../hwp-text-advances';
import { sourceNestedCellFaces } from './source-nested-cell-face';

type Box = { pageIndex: number; x: number; y: number; w: number; h: number };
type CellBox = Box & { cellIdx: number; row: number; col: number; rowSpan: number; colSpan: number };
type PageRun = { x: number; y: number; w: number; h: number; text: string; displayText?: string;
  secIdx: number; parentParaIdx?: number; headerFooter?: boolean;
  cellPath?: { controlIndex: number; cellIndex: number; cellParaIndex: number }[] };
type Position = { page: number; y: number };
type SourceBlock = { token: string; before: number; after: number; first?: Position; end?: Position;
  trailing: number | null; cut?: HwpNestedTableCut };

const tolerance = .16;
const close = (a: number, b: number) => Math.abs(a - b) <= tolerance;
const within = (inner: Pick<Box, 'x' | 'y' | 'w' | 'h'>, outer: Box) => inner.w > 0 && inner.h > 0 &&
  inner.x >= outer.x - tolerance && inner.y >= outer.y - tolerance &&
  inner.x + inner.w <= outer.x + outer.w + tolerance && inner.y + inner.h <= outer.y + outer.h + tolerance;

function sourceCut(doc: HwpDocument, section: number, paragraph: number, block: number, rows: number, columns: number,
  cells: number, address: (cell: number, p?: number) => string, sheets: readonly [number, CellBox[]][],
  runs: ReadonlyMap<number, PageRun[]>): HwpNestedTableCut | undefined {
  if (sheets.length !== 2 || cells > 1000) return;
  const [[firstPage, first], [secondPage, second]] = sheets;
  if (secondPage !== firstPage + 1) return;
  // Reader fragment indices restart on the next page. Their row and column
  // still identify the authored grid, including empty continuation cells.
  const boundary = Math.max(...first.map(cell => cell.row)) + 1;
  if (!(boundary > 0 && boundary < rows) || Math.min(...second.filter(cell => cell.row >= boundary).map(cell => cell.row)) !== boundary)
    return;
  const merges: HwpNestedTableCut['merges'] = [];
  const tokens: { row: number; col: number; token: string }[] = [];
  for (let index = 0; index < cells; index++) {
    const { row, col, rowSpan, colSpan } = JSON.parse(doc.getCellInfoByPath(section, paragraph, address(index)));
    const count = doc.getCellParagraphCountByPath(section, paragraph, address(index));
    if (!(count > 0 && count <= 50)) return;
    const crosses = row < boundary && row + rowSpan > boundary;
    let original = '';
    for (let p = 0; p < count; p++) {
      const path = address(index, p);
      if (crosses) {
        const metric = JSON.parse(doc.getParagraphLineMetrics(section, paragraph, path));
        if (!metric.stored || metric.controlCount) return;
      }
      original += doc.getTextInCellByPath(section, paragraph, path, 0, 0xffffffff);
    }
    tokens.push({ row, col, token: hwpNestedCellToken(row, col, rowSpan, colSpan, original) });
    if (!crosses) continue;
    const head = first.find(b => b.row === row && b.col === col), tail = second.find(b => b.row === row && b.col === col);
    if (!head || !tail || !close(head.x, tail.x) || !close(head.w, tail.w)) return;
    const painted = (box: CellBox) => (runs.get(box.pageIndex) ?? [])
      .filter(run => within(run, box) && normalizeHwpWrapText(run.displayText ?? run.text))
      .sort((a, b) => a.y - b.y || a.x - b.x)
      .map(run => normalizeHwpWrapText(run.displayText ?? run.text)).join('');
    if (painted(head) !== normalizeHwpWrapText(original) || painted(tail)) return;
    merges.push([row, col, rowSpan]);
  }
  if (second.some(box => box.row < boundary && !merges.some(([row, col]) => row === box.row && col === box.col))) return;
  merges.sort((a, b) => a[0] - b[0] || a[1] - b[1]);
  const fingerprint = hwpNestedLayoutFingerprint(tokens.sort((a, b) => a.row - b.row || a.col - b.col).map(c => c.token));
  return { block, rows, columns, boundary, fingerprint, merges };
}

/** Observe the source geometry of a continuing one-cell table that contains
 * editable paragraphs and nested tables. Only same-sheet gaps and witnessed
 * nested row boundaries become optional rendering metadata. */
export function sourceNestedTableLayout(doc: HwpDocument, section: number, paragraph: number, control: number,
  outer: readonly Box[]): string | undefined {
  if (outer.length < 3 || outer.length > 128) return;
  try {
    const boxes = [...outer].sort((a, b) => a.pageIndex - b.pageIndex);
    if (!boxes.every((box, i) => box.pageIndex === boxes[0].pageIndex + i && close(box.w, boxes[0].w) && box.h > 0)) return;
    const count = doc.getCellParagraphCount(section, paragraph, control, 0);
    if (!(count >= 4 && count <= 2048)) return;
    const pages = new Map(boxes.map(box => [box.pageIndex, box]));
    const runs = new Map<number, PageRun[]>();
    const paragraphRuns = new Map<number, (PageRun & { page: number })[]>();
    for (const box of boxes) {
      const page = JSON.parse(doc.getPageInfo(box.pageIndex));
      const top = page.marginTop + page.marginHeader, bottom = page.footerArea?.y;
      if (page.columns?.length !== 1 || !Number.isFinite(top) || !Number.isFinite(bottom) ||
          box.y < top - tolerance || box.y + box.h > bottom + tolerance)
        return;
      const painted = (JSON.parse(doc.getPageTextLayout(box.pageIndex)).runs as PageRun[])
        .filter(run => run.secIdx === section && run.parentParaIdx === paragraph && !run.headerFooter);
      runs.set(box.pageIndex, painted);
      for (const run of painted) {
        const path = run.cellPath, [owner] = path ?? [];
        if (path?.length !== 1 || owner.controlIndex !== control || owner.cellIndex !== 0 || run.h <= 0 ||
            run.y < box.y - tolerance || run.y + run.h > box.y + box.h + tolerance) continue;
        const list = paragraphRuns.get(owner.cellParaIndex) ?? [];
        list.push({ ...run, page: box.pageIndex });
        paragraphRuns.set(owner.cellParaIndex, list);
      }
    }
    const blocks: SourceBlock[] = [], faces: HwpNestedCellFace[] = [];
    for (let i = 0; i < count; i++) {
      const owner = { controlIndex: control, cellIndex: 0, cellParaIndex: i }, path = JSON.stringify([owner]);
      const literal = doc.getTextInCellByPath(section, paragraph, path, 0, 0xffffffff);
      const properties = JSON.parse(doc.getCellParaPropertiesAtByPath(section, paragraph, path));
      const metrics = JSON.parse(doc.getParagraphLineMetrics(section, paragraph, path));
      const before = Number(properties.spacingBefore) / 2, after = Number(properties.spacingAfter) / 2;
      if (!metrics.stored || !metrics.lines?.length || ![before, after].every(n => Number.isFinite(n) && n >= 0 && n <= 200)) return;
      if (!metrics.controlCount) {
        const runs = (paragraphRuns.get(i) ?? []).sort((a, b) => a.page - b.page || a.y - b.y);
        const head = runs[0], tail = runs.at(-1), firstLine = metrics.lines[0], lastLine = metrics.lines.at(-1);
        blocks.push({ token: hwpNestedParagraphToken(literal), before, after, trailing: null,
          first: head && close(head.h, firstLine.height) ? { page: head.page, y: head.y } : undefined,
          end: tail && close(tail.h, lastLine.height) ? { page: tail.page, y: tail.y + lastLine.height + lastLine.spacing } : undefined });
        continue;
      }
      const [table] = metrics.tables ?? [];
      if (metrics.controlCount !== 1 || metrics.tables?.length !== 1 || typeof table.inline !== 'boolean' || normalizeHwpWrapText(literal) ||
          metrics.lines.length !== 1 || !(table.outerTop >= 0 && table.outerBottom >= 0)) return;
      const address = (index: number, p = 0) => JSON.stringify([owner,
        { controlIndex: table.index, cellIndex: index, cellParaIndex: p }]);
      const dims = JSON.parse(doc.getTableDimensionsByPath(section, paragraph, address(0)));
      const nestedBoxes = JSON.parse(doc.getTableCellBboxesByPath(section, paragraph, address(0))) as CellBox[];
      const byPage = new Map<number, CellBox[]>();
      for (const box of nestedBoxes) {
        const parent = pages.get(box.pageIndex);
        if (!parent || !within(box, parent)) return;
        const group = byPage.get(box.pageIndex) ?? [];
        group.push(box);
        byPage.set(box.pageIndex, group);
      }
      const sheets = [...byPage].sort((a, b) => a[0] - b[0]);
      if (!sheets.length) return;
      const [firstPage, first] = sheets[0], [lastPage, last] = sheets.at(-1)!;
      if (sheets.length === 1 && table.inline && dims.rowCount > 2 && dims.cellCount <= 400)
        faces.push(...sourceNestedCellFaces(doc, section, paragraph, i, address, first, runs.get(firstPage) ?? []));
      const trailing = table.outerBottom + (table.inline ? Math.max(0, metrics.lines[0].spacing) : 0);
      blocks.push({ token: hwpNestedTableToken(dims.rowCount, dims.colCount), before, after, trailing,
        first: { page: firstPage, y: Math.min(...first.map(box => box.y)) - table.outerTop },
        end: { page: lastPage, y: Math.max(...last.map(box => box.y + box.h)) + trailing },
        cut: sourceCut(doc, section, paragraph, i, dims.rowCount, dims.colCount, dims.cellCount, address, sheets, runs) });
    }
    const gaps: HwpNestedTableGap[] = [];
    for (let i = 0; i + 1 < blocks.length; i++) {
      const left = blocks[i], right = blocks[i + 1];
      const after = left.after, before = right.before;
      if (after + before > 0 && left.end && right.first && left.end.page === right.first.page &&
          close(right.first.y - left.end.y, after + before))
        gaps.push([i, after, before, left.trailing === null ? null : Math.round(left.trailing * 10000) / 10000]);
    }
    const cuts = blocks.flatMap(block => block.cut ? [block.cut] : []);
    const layout = JSON.stringify({ blocks: count, fingerprint: hwpNestedLayoutFingerprint(blocks.map(block => block.token)), gaps, cuts, ...(faces.length ? { faces } : {}) });
    if (parseHwpNestedTableLayout(layout)) return layout;
  } catch { /* Optional source geometry leaves the editable table unchanged. */ }
}
