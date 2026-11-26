import type { HwpDocument } from '@rhwp/core';
import {
  hwpRelatedParagraphToken, hwpRelatedTableToken, parseHwpRelatedTableLayout, type HwpRelatedTableLayout,
} from '../hwp-related-table-layout';
import { physicalLineHeight, sourceCellBoundsTolerance, sourceSplitCellBefore, type HwpPaintedLine } from './paragraph-layout';

type Position = { type: string; page: number; x: number; y: number; w: number; h: number };
type Box = { cellIdx: number; row: number; col: number; rowSpan: number; colSpan: number;
  pageIndex: number; x: number; y: number; w: number; h: number };
type Cell = { row: number; col: number; rowSpan: number; colSpan: number; height: number; width: number;
  isHeader: boolean; verticalAlign: number; textDirection: number; diagonalSlash: boolean; diagonalBackSlash: boolean;
  applyInnerMargin: boolean; paddingTop: number; paddingBottom: number };
const tolerance = sourceCellBoundsTolerance;
const close = (a: number, b: number) => Math.abs(a - b) <= tolerance;
const round = (value: number) => Math.round(value * 100) / 100;
const px = (value: number) => round(value / 75);

function bodyBounds(doc: HwpDocument, pageIndex: number) {
  const page = JSON.parse(doc.getPageInfo(pageIndex));
  return { left: page.bodyLeft, right: page.bodyRight, top: page.marginTop + page.marginHeader,
    bottom: page.footerArea?.y, columns: page.columns?.length || 1 };
}

function readCellParagraphs(doc: HwpDocument, address: [number, number, number, number]) {
  const [section, paragraph, controlIndex, cellIndex] = address;
  return Array.from({ length: doc.getCellParagraphCount(...address) }, (_, i) =>
    doc.getTextInCellByPath(section, paragraph, JSON.stringify([{ controlIndex, cellIndex, cellParaIndex: i }]), 0, 0xffffffff));
}

function inlineTableProof(doc: HwpDocument, section: number, paragraph: number, control: number, position: Position) {
  const address = [section, paragraph, control] as const;
  const properties = JSON.parse(doc.getTableProperties(...address));
  const dims = JSON.parse(doc.getTableDimensions(...address));
  if (position.type !== 'table' || !properties.treatAsChar || properties.pageBreak !== 2 || properties.hasCaption ||
    properties.allowOverlap || !properties.restrictInPage || properties.cellSpacing || properties.textWrap !== 'TopAndBottom' ||
    properties.vertRelTo !== 'Para' || properties.vertAlign !== 'Top' || properties.horzRelTo !== 'Para' ||
    dims.colCount !== 1 || dims.rowCount !== dims.cellCount || dims.rowCount < 1 || dims.rowCount > 48) return;
  const boxes = JSON.parse(doc.getTableCellBboxes(...address)) as Box[];
  if (boxes.length !== dims.rowCount || !boxes.every((box, row) => box.pageIndex === position.page &&
    box.row === row && box.cellIdx === row && box.col === 0 && box.rowSpan === 1 && box.colSpan === 1 &&
    close(box.x, position.x) && close(box.w, position.w) && box.h > 0)) return;
  const text = boxes.flatMap(box => readCellParagraphs(doc, [section, paragraph, control, box.cellIdx]));
  if (!text.some(value => value.trim())) return;
  return { rows: dims.rowCount, text: hwpRelatedTableToken(text) };
}

/** The two adjacent source controls start a new physical sheet together.
 * Their independent width/placement proof is only used alongside a split
 * row with two witnessed cell flows in the immediately preceding table. */
export function sourceRelatedTableLayout(doc: HwpDocument, section: number, paragraph: number, preceding: number,
  controls: readonly [number, number], positions: readonly [Position, Position],
  painted: Map<string, HwpPaintedLine[]>): string | undefined {
  try {
    const [left, right] = positions;
    if (left.page !== right.page || left.page < 1 || !close(left.y, right.y) || right.x - left.x - left.w < .5) return;
    const paper = bodyBounds(doc, left.page);
    const height = paper.bottom - paper.top, width = paper.right - paper.left;
    if (paper.columns !== 1 || !positions.every(pos => pos.x >= paper.left && pos.x + pos.w <= paper.right &&
      pos.y >= paper.top && pos.y - paper.top <= 6 && pos.h >= height * .8 && pos.y + pos.h <= paper.bottom + tolerance)) return;
    const pair = controls.map((control, i) => {
      const position = positions[i];
      const proof = inlineTableProof(doc, section, paragraph, control, position);
      return proof && { ...proof, left: round(position.x - paper.left), width: round(position.w) };
    });
    if (!pair[0] || !pair[1]) return;

    const prior = [section, paragraph - 1, preceding] as const;
    const dims = JSON.parse(doc.getTableDimensions(...prior));
    const previous = JSON.parse(doc.getTableProperties(...prior));
    if (dims.colCount !== 2 || dims.rowCount < 3 || dims.rowCount > 2000 || dims.cellCount > dims.rowCount * 2 ||
      previous.treatAsChar || previous.pageBreak !== 2 || previous.hasCaption || previous.allowOverlap ||
      !previous.restrictInPage || previous.cellSpacing || previous.textWrap !== 'TopAndBottom') return;
    const boxes = JSON.parse(doc.getTableCellBboxes(...prior)) as Box[];
    if (Math.max(...boxes.map(box => box.pageIndex)) !== left.page - 1) return;
    const byRow = new Map<number, Box[]>();
    for (const box of boxes) {
      const fragments = byRow.get(box.row) ?? [];
      fragments.push(box);
      byRow.set(box.row, fragments);
    }
    const choices: HwpRelatedTableLayout[] = [];
    for (const [row, fragments] of byRow) {
      if (row < 1 || row >= dims.rowCount - 1 || fragments.length !== 6 ||
        fragments.some(box => box.rowSpan !== 1 || box.colSpan !== 1)) continue;
      const columns = [0, 1].map(col => fragments.filter(box => box.col === col).sort((a, b) => a.pageIndex - b.pageIndex));
      const [first, second] = columns;
      if (first.length !== 3 || second.length !== 3 || !first.every((piece, i) => {
        const other = second[i], bounds = bodyBounds(doc, piece.pageIndex);
        return piece.cellIdx === first[0].cellIdx && other.cellIdx === second[0].cellIdx &&
          bounds.columns === 1 && piece.pageIndex === other.pageIndex && piece.pageIndex === first[0].pageIndex + i &&
          piece.y >= bounds.top - tolerance && piece.y + piece.h <= bounds.bottom + tolerance &&
          piece.h > 0 && close(piece.y, other.y) && close(piece.h, other.h) && close(piece.x + piece.w, other.x);
      })) continue;
      const measures = columns.map(column => {
        const address: [number, number, number, number] = [...prior, column[0].cellIdx];
        const cell: Cell = { ...JSON.parse(doc.getCellInfo(...address)), ...JSON.parse(doc.getCellProperties(...address)) };
        if (cell.row !== row || cell.col !== column[0].col || cell.rowSpan !== 1 || cell.colSpan !== 1 ||
          cell.isHeader || cell.verticalAlign !== 0 || cell.textDirection || cell.diagonalSlash || cell.diagonalBackSlash ||
          !close(px(cell.width), column[0].w) || !(column.reduce((total, b) => total + b.h, 0) > px(cell.height) + tolerance)) return;
        const padding = (side: 'paddingTop' | 'paddingBottom') => px(cell.applyInnerMargin && cell[side] >= 0 ? cell[side] : previous[side]);
        const proven = sourceSplitCellBefore(doc, address, painted, column, padding('paddingTop'), padding('paddingBottom'), true);
        if (!proven) return;
        const paragraphs = readCellParagraphs(doc, address);
        if (paragraphs.length !== proven.before.length || paragraphs.some(text => !text.trim() || /[\x00-\x09\x0b-\x1f\ufffc]/u.test(text))) return;
        const before = paragraphs.map((text, i) => {
          const properties = JSON.parse(doc.getCellParaPropertiesAt(...address, i));
          const char = JSON.parse(doc.getCellCharPropertiesAt(...address, i, 0));
          return { gap: proven.before[i], text: hwpRelatedParagraphToken(text), font: round(char.fontSize / 75),
            pitch: round(physicalLineHeight(properties, char)), breaks: text.split('\n').length - 1 };
        });
        return { fills: proven.fillsFragments, width: column[0].w, before };
      });
      const [a, b] = measures;
      if (a && b && (a.fills || b.fills))
        choices.push({ bodyWidth: round(width), bodyHeight: round(height), rows: dims.rowCount, row,
          rowWidths: [a.width, b.width], before: [a.before, b.before], pair: [pair[0], pair[1]] });
    }
    if (choices.length !== 1) return;
    const value = JSON.stringify(choices[0]);
    return parseHwpRelatedTableLayout(value) ? value : undefined;
  } catch { /* The ordinary editable tables remain available without native proof. */ }
}
