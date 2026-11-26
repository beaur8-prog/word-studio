import type { HwpDocument } from '@rhwp/core';
import { hwpScaledCellFace, hwpWrapFingerprint, normalizeHwpWrapText, type HwpSourceWrap } from '../hwp-text-advances';
import { sourceCellBoundsTolerance } from './paragraph-layout';

type CellBox = { x: number; y: number; w: number; h: number };
export type NativeCellInkRun = { secIdx: number; parentParaIdx: number; x: number; y: number; w: number; h: number; text?: string;
  cellPath?: { controlIndex: number; cellIndex: number; cellParaIndex: number }[] };

export function sourceScaledCellLine(doc: HwpDocument, section: number, paragraph: number, control: number, cell: number,
  box: CellBox, leftPadding: number, rightPadding: number, rowHeight: number,
  painted: readonly NativeCellInkRun[]): HwpSourceWrap | undefined {
  const tolerance = sourceCellBoundsTolerance;
  const width = box.w - leftPadding - rightPadding;
  if (!(width > 0 && rowHeight > 0 && Math.abs(box.h - rowHeight) <= tolerance)) return;
  const address = JSON.stringify([{ controlIndex: control, cellIndex: cell, cellParaIndex: 0 }]);
  const literal = doc.getTextInCellByPath(section, paragraph, address, 0, 0xffffffff);
  if (!/^[가-힣0-9]+(?: [가-힣0-9]+)*$/u.test(literal) || literal.length > 60) return;
  const text = normalizeHwpWrapText(literal);
  if (text.length < 2 || text.length > 40 || !/[가-힣]/u.test(text)) return;
  const metrics = JSON.parse(doc.getParagraphLineMetrics(section, paragraph, address));
  if (!metrics.stored || metrics.controlCount || metrics.lines?.length !== 1 || metrics.lines[0].textStart !== 0 ||
      Math.abs(metrics.lines[0].width - width) > tolerance) return;
  const runs = JSON.parse(doc.getCharShapeRunsInCellByPath(section, paragraph, address, 0, literal.length));
  if (runs.length !== 1 || runs[0].startOffset !== 0 || runs[0].endOffset !== literal.length) return;
  const { fontFamily, fontSize, ratios, spacings, relativeSizes } =
    JSON.parse(doc.getCellCharPropertiesAtByPath(section, paragraph, address, 0));
  const scale = ratios?.[0], size = fontSize / 75;
  if (fontFamily !== hwpScaledCellFace || !Array.isArray(ratios) || ratios.length !== 7 ||
      !Number.isInteger(scale) || scale < 50 || scale >= 100 || ratios.some(value => value !== scale) ||
      !Array.isArray(spacings) || spacings.length !== 7 || spacings.some(value => value !== 0) ||
      !Array.isArray(relativeSizes) || relativeSizes.length !== 7 || relativeSizes.some(value => value !== 100) ||
      !(size > 0 && size <= rowHeight && rowHeight <= 180)) return;
  const hits = painted.filter(run => {
    const path = run.cellPath?.[0];
    return run.secIdx === section && run.parentParaIdx === paragraph && run.cellPath?.length === 1 &&
      path?.controlIndex === control && path.cellIndex === cell && path.cellParaIndex === 0 &&
      !!normalizeHwpWrapText(run.text || '');
  });
  if (hits.length !== 1) return;
  const [run] = hits;
  if (run.text !== literal || !(run.w > 0 && run.h > 0) ||
      !(run.x >= box.x + leftPadding - tolerance && run.x + run.w <= box.x + box.w - rightPadding + tolerance &&
        run.y >= box.y - tolerance && run.y + run.h <= box.y + box.h + tolerance)) return;
  return { fingerprint: hwpWrapFingerprint(literal), ends: [text.length], cellWidth: { outer: box.w, text: width },
    scaledCellLine: { scale, size, rowHeight } };
}

export function sourceCellCaptionLine(doc: HwpDocument, section: number, paragraph: number, control: number, cell: number,
  box: CellBox, leftPadding: number, rightPadding: number, painted: readonly NativeCellInkRun[], repeatedHeader = false): HwpSourceWrap | undefined {
  if (![box.x, box.y, box.w, box.h, leftPadding, rightPadding].every(Number.isFinite)) return;
  const ordinal = repeatedHeader ? 0 : 1;
  const address = JSON.stringify([{ controlIndex: control, cellIndex: cell, cellParaIndex: ordinal }]);
  const literal = doc.getTextInCellByPath(section, paragraph, address, 0, 0xffffffff);
  if (literal.length > 60 || /[\u0000-\u001f\u007f-\u009f\uD800-\uDFFF\p{Cf}\ufffc]/u.test(literal)) return;
  const text = normalizeHwpWrapText(literal);
  if ((repeatedHeader ? !/^[가-힣]{2}$/u.test(text) : text.length <= 12 || text.length > 40) || literal.trim() !== literal) return;

  const metrics = JSON.parse(doc.getParagraphLineMetrics(section, paragraph, address));
  const width = box.w - leftPadding - rightPadding, tolerance = sourceCellBoundsTolerance;
  if (!metrics.stored || metrics.controlCount || metrics.lines?.length !== 1 || metrics.lines[0].textStart !== 0 ||
      !(width > 0) || Math.abs(metrics.lines[0].width - width) > tolerance) return;
  const runs = JSON.parse(doc.getCharShapeRunsInCellByPath(section, paragraph, address, 0, literal.length));
  if (runs.length !== 1 || runs[0].startOffset !== 0 || runs[0].endOffset !== literal.length) return;
  const { ratios, spacings, relativeSizes } = JSON.parse(doc.getCellCharPropertiesAtByPath(section, paragraph, address, 0));
  if (!Array.isArray(ratios) || ratios.length !== 7 || ratios.some(value => value !== 100) ||
      !Array.isArray(spacings) || spacings.length !== 7 || spacings.some(value => value !== 0) ||
      !Array.isArray(relativeSizes) || relativeSizes.length !== 7 || relativeSizes.some(value => value !== 100)) return;

  const hits = painted.filter(run => {
    const item = run.cellPath?.[0];
    return run.secIdx === section && run.parentParaIdx === paragraph && run.cellPath?.length === 1 &&
      item?.controlIndex === control && item.cellIndex === cell && item.cellParaIndex === ordinal &&
      !!normalizeHwpWrapText(run.text || '');
  });
  if (hits.length !== 1) return;
  const [run] = hits;
  if (normalizeHwpWrapText(run.text || '') !== text || ![run.x, run.y, run.w, run.h].every(Number.isFinite) || !(run.w > 0 && run.h > 0) ||
      run.x < box.x + leftPadding - tolerance || run.x + run.w > box.x + box.w - rightPadding + tolerance ||
      run.y < box.y - tolerance || run.y + run.h > box.y + box.h + tolerance) return;
  return { fingerprint: hwpWrapFingerprint(literal), ends: [text.length], cellWidth: { outer: box.w, text: width },
    ...(repeatedHeader ? { singleHeaderLine: true } : { singleCellLine: true }) };
}
