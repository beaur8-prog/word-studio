import type { HwpDocument } from '@rhwp/core';
import type { HwpAuthoredCellTerminal } from '../hwp-cell-line';
import { hwpWrapFingerprint } from '../hwp-text-advances';
import type { HwpPaintedLine } from './paragraph-layout';

type CellBox = { cellIdx: number; pageIndex: number; y: number; h: number; w: number };
type CellPadding = { paddingTop: number; paddingBottom: number; paddingLeft: number; paddingRight: number };

/** Repeated headings and whole-row cuts identify the floating, continuous
 * table layout. A body cell must be witnessed as a single box. */
export function sourceTableTerminalBoxes(properties: Record<string, unknown>, boxes: readonly CellBox[]) {
  const inset = Number(properties.outerTop), offset = Number(properties.vertOffset);
  const symmetricInset = inset > 0 && offset > 0 && offset <= inset &&
    properties.outerBottom === inset && properties.outerLeft === inset && properties.outerRight === inset &&
    properties.vertAlign === 'Top' && properties.horzRelTo === 'Column' &&
    properties.horzAlign === 'Left' && properties.horzOffset === 0;
  const flushPlacement = properties.vertOffset === 0 && properties.outerTop === 0 && properties.outerBottom === 0;
  if (properties.treatAsChar || !properties.repeatHeader || properties.pageBreak !== 1 || properties.cellSpacing !== 0 ||
      properties.hasCaption || properties.allowOverlap || properties.textWrap !== 'TopAndBottom' ||
      properties.vertRelTo !== 'Para' || (!flushPlacement && !symmetricInset) ||
      new Set(boxes.map(box => box.pageIndex)).size < 2) return;
  const byCell = new Map<number, CellBox | null>();
  for (const box of boxes) byCell.set(box.cellIdx, byCell.has(box.cellIdx) ? null : box);
  return byCell;
}

export function sourceCellTerminal(doc: HwpDocument, section: number, paragraph: number, control: number, index: number,
  cell: CellPadding, box: CellBox, painted: readonly HwpPaintedLine[], middle: boolean): HwpAuthoredCellTerminal | undefined {
  // Reader boxes are reported to a tenth of a px. Compare the independently
  // stored line, its painted location and the cell's available unpadded area.
  const tolerance = .12;
  const [line] = painted;
  if (painted.length !== 1 || line.page !== box.pageIndex || line.y < box.y - tolerance ||
      line.y + line.h > box.y + box.h + tolerance || box.w > 2000 || box.h > 2000) return;
  const innerHeight = box.h - (cell.paddingTop + cell.paddingBottom) / 75;
  const innerWidth = box.w - (cell.paddingLeft + cell.paddingRight) / 75;
  if (line.h > innerHeight + tolerance || !(innerWidth > 0)) return;
  const address = JSON.stringify([{ controlIndex: control, cellIndex: index, cellParaIndex: 0 }]);
  const metrics = JSON.parse(doc.getParagraphLineMetrics(section, paragraph, address));
  const stored = metrics.lines?.[0];
  if (!metrics.stored || metrics.controlCount || metrics.lines?.length !== 1 || stored.textStart !== 0 ||
      Math.abs(stored.height - line.h) > tolerance || Math.abs(stored.width - innerWidth) > tolerance) return;
  const text = doc.getTextInCellByPath(section, paragraph, address, 0, 0xffffffff);
  if (!text.trim() || text.length > 256 || /[\u0000-\u001f\ufffc]/u.test(text)) return;
  return { text: hwpWrapFingerprint(text), width: box.w, row: box.h, ...(middle ? { middle: true } : {}) };
}
