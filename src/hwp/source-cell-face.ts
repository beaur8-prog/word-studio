import type { HwpDocument } from '@rhwp/core';
import { hwpCompactCellFace, hwpCompactHeaderFace, hwpWrapFingerprint, normalizeHwpWrapText, type HwpSourceWrap } from '../hwp-text-advances';
import { sourceCellBoundsTolerance } from './paragraph-layout';
import type { NativeCellInkRun } from './source-caption-cell-line';

type SourceAddress = { section: number; paragraph: number; control: number; cell: number; index: number };
type SourceBox = { w: number; y: number; h: number };
type SourceStyle = { spacingBefore: number; spacingAfter: number; marginLeft: number; marginRight: number; indent: number; alignment: string };
type PaintedText = NativeCellInkRun & { displayText?: string };
export type CompactCellFacePosition = 'edge' | 'interior' | 'header' | 'continued' | 'scaledContinuation';

/** The stored line count and the reader's actual visible lines must agree.
 * A one-line member of a multi-paragraph row, or a single authored hard
 * break, is an unambiguous target for a substitute when the face is absent. */
export function sourceHwpCompactCellFace(doc: HwpDocument, address: SourceAddress, box: SourceBox,
  count: number, style: SourceStyle, position: CompactCellFacePosition | { intactColumns: number; paragraphs?: 2 },
  readPaint: () => readonly PaintedText[], scaledLineCount?: 2 | 3): HwpSourceWrap | undefined {
  const intactRow = typeof position === 'object'
    ? { height: box.h, columns: position.intactColumns, ...(position.paragraphs && { paragraphs: position.paragraphs }) }
    : undefined;
  const scaled = position === 'scaledContinuation' || !!intactRow;
  const continued = position === 'continued' || scaled;
  if (count < 1 || count > (continued ? 64 : 12) || style.spacingBefore || style.spacingAfter) return;
  const { section, paragraph, control, cell, index } = address;
  const header = position === 'header';
  const hanging = position === 'interior' || continued ? -style.indent : 0;
  if (header && (count !== 1 || style.alignment !== 'center' || style.indent || style.marginLeft || style.marginRight) ||
      (position === 'interior' || continued) && (!scaled && count < 2 || !continued && (count !== 2 || index !== 1) ||
        style.alignment !== 'justify' || (scaled ? hanging < 0 : hanging <= 0) || style.marginLeft || style.marginRight)) return;
  const key = JSON.stringify([{ controlIndex: control, cellIndex: cell, cellParaIndex: index }]);
  const text = doc.getTextInCellByPath(section, paragraph, key, 0, 0xffffffff);
  const literal = normalizeHwpWrapText(text), hard = text.indexOf('\n');
  if ((header ? !/^[가-힣]{2}$/u.test(text) : literal.length < (scaled ? 8 : 12) || literal.length > 160 || !/[가-힣]/u.test(text)) ||
      /[\x00-\x09\x0b-\x1f\uD800-\uDFFF\p{Bidi_Control}\ufffc]/u.test(text) ||
      (intactRow?.paragraphs && hard >= 0) ||
      (position === 'edge' ? hard < 0 ? count < 2 || style.marginLeft || style.marginRight || style.indent :
        count !== 1 || hard !== text.lastIndexOf('\n') : scaled ? hard !== text.lastIndexOf('\n') : hard >= 0)) return;
  const metric = JSON.parse(doc.getParagraphLineMetrics(section, paragraph, key));
  const lines = metric.lines;
  if (!metric.stored || metric.controlCount || !lines?.length ||
      (continued ? lines.length > (intactRow?.paragraphs ? 6 : scaled ? 2 : 3) : lines.length !== (hard < 0 ? 1 : 2)) ||
      intactRow?.paragraphs && lines.length < 2 || lines[0].textStart !== 0 ||
      hard >= 0 && lines[1]?.textStart !== hard + 1) return;
  const { width, height: size, spacing } = lines[0], pitch = size + spacing;
  const close = (a: number, b: number) => Math.abs(a - b) <= sourceCellBoundsTolerance;
  if (!(width > 0 && width <= box.w && size > 0 && pitch > size && pitch <= box.h) || (intactRow ? hanging >= width : hanging > width / 4) ||
      header && (width < 2 * size || width > 2.5 * size) ||
      lines.some((line: { width: number; height: number; spacing: number }) =>
        !close(line.width, width) || !close(line.height, size) || !close(line.spacing, spacing))) return;
  const runs = JSON.parse(doc.getCharShapeRunsInCellByPath(section, paragraph, key, 0, text.length));
  const family = header ? hwpCompactHeaderFace : hwpCompactCellFace;
  const ratio = header ? 100 : 95;
  if (!runs.length || header && runs.length !== 1 || !scaled && hanging && runs.length < 2 || runs.some((run: { startOffset: number }) => {
    const char = JSON.parse(doc.getCellCharPropertiesAtByPath(section, paragraph, key, run.startOffset));
    return char.bold || char.italic || !close(char.fontSize / 75, size) ||
      char.fontFamilies?.length !== 7 || !char.fontFamilies.every((face: string) => face === family) ||
      char.ratios?.length !== 7 || !char.ratios.every((scale: number) => scale === ratio) ||
      char.relativeSizes?.length !== 7 || !char.relativeSizes.every((relative: number) => relative === 100);
  })) return;

  const bands = new Map<number, PaintedText[]>();
  for (const run of readPaint()) {
    const path = run.cellPath?.[0];
    if (run.secIdx !== section || run.parentParaIdx !== paragraph || run.cellPath?.length !== 1 ||
        path?.controlIndex !== control || path.cellIndex !== cell || path.cellParaIndex !== index || !run.text) continue;
    if (run.y < box.y - sourceCellBoundsTolerance || run.y + run.h > box.y + box.h + sourceCellBoundsTolerance) return;
    const line = bands.get(run.y) ?? [];
    line.push(run); bands.set(run.y, line);
  }
  if (bands.size !== lines.length) return;
  const ends = continued ? [
    ...lines.slice(1).map((line: { textStart: number }) => normalizeHwpWrapText(text.slice(0, line.textStart)).length),
    literal.length,
  ] : hard < 0 ? [literal.length] : [normalizeHwpWrapText(text.slice(0, hard)).length, literal.length];
  let counted = 0;
  const seen = [...bands].sort(([a], [b]) => a - b).map(([, line]) => normalizeHwpWrapText(
    line.sort((a, b) => a.x - b.x).map(run => run.displayText ?? run.text).join('')));
  if (seen.join('') !== literal || seen.some((line, i) => (counted += line.length) !== ends[i])) return;
  if (scaled) return { fingerprint: hwpWrapFingerprint(text), ends, cellWidth: { outer: box.w, text: width },
    cellWidthFallback: { scale: ratio, size, pitch, hanging, ...(hard >= 0 && { hardBreak: true }),
      ...(intactRow && { intactRow: { ...intactRow, textWidth: width } }) } };
  return { fingerprint: hwpWrapFingerprint(text), ends, cellWidth: { outer: box.w, text: width },
    authoredCellFace: { size, pitch, rowHeight: box.h, ...(hanging > 0 && { hanging }), ...(header && { header: true }),
      ...(continued && { continued: true }),
      ...(continued && lines.length === scaledLineCount && { scale: ratio }) } };
}
