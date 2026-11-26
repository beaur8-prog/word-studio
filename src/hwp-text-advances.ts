import { floatingTextBox, type HwpFloatingTextBox } from './hwp-floating-textbox';
import { noteTerminal, type HwpNoteTerminal } from './hwp-note-terminal';

export const hwpSourceSpacesAttribute = 'data-hwp-source-spaces';
export const hwpSourceSpacesNamespace = 'urn:wordstudio:hwp-source-spaces:1';
export const hwpSourceWrapAttribute = 'data-hwp-source-wrap';
export const hwpScaledCellFace = 'KoPub돋움체 Bold';
export const hwpCompactCellFace = 'KoPub돋움체 Light';
export const hwpCompactHeaderFace = 'KoPub돋움체 Bold';
export const hwpCellFaceFallbackSelector = `p[${hwpSourceWrapAttribute}*="authoredCellFace"]`;
export const hwpCellWidthFallbackSelector = `p[${hwpSourceWrapAttribute}*="cellWidthFallback"]`;
export const hwpSingleBodyFace = 'KoPub바탕체 Light';
export const hwpBodyFallbackSelector = `p[${hwpSourceWrapAttribute}*="singleBodyFallback"],` +
  `p[${hwpSourceWrapAttribute}*="authoredBodyLines"][style*="${hwpSingleBodyFace}"]`;

export function hwpShortCellLabel(text: string) {
  return /^[가-힣]{2,12}$/u.test(text);
}

export type HwpSourceOrigin = { page: number; x: number; y: number };
export type HwpSourceCellWidth = { outer: number; text: number };
export type HwpSourceScaleRun = { start: number; scale: number; width: number; size: number; font: string };
export type HwpNestedHangingScale = { scale: number; hanging: number; size: number };
export type HwpScaledCellLine = { scale: number; size: number; rowHeight: number };
export type HwpCellWidthFallback = { scale: number; hanging: number; size: number; pitch: number; hardBreak?: true;
  intactRow?: { height: number; columns: number; textWidth?: number; paragraphs?: 2 } };
export type HwpAuthoredBodyLines = { origin: HwpSourceOrigin; width: number; hanging: number; size: number; pitch: number };
export type HwpAuthoredNativeLine = HwpAuthoredBodyLines & { top: number; ascent: number; typography: string };
export type HwpForcedBodyEnd = { baseline: number; room: number; width: number; spaceBefore: number; pitch: number;
  font: number; hanging: number; preceding: string; precedingLines: number };
export type HwpSourceWrap = { fingerprint: string; ends: number[]; origin?: HwpSourceOrigin; cellWidth?: HwpSourceCellWidth;
  bodyWidth?: number; columnWidth?: number; columnPitches?: [number, number]; scaleRun?: HwpSourceScaleRun; nestedHanging?: HwpNestedHangingScale; singleCellLine?: true; singleHeaderLine?: true;
  scaledCellLine?: HwpScaledCellLine; splitCell?: true; singleBodyFallback?: true; columnCellHeight?: number;
  columnCellPitch?: number; columnCellStories?: number; expandedRowHeight?: number; expandedRowLine?: { hanging: number; pitch: number; scale: number }; authoredBodyLines?: HwpAuthoredBodyLines;
  authoredNativeLine?: HwpAuthoredNativeLine;
  forcedBodyEnd?: HwpForcedBodyEnd;
  noteTerminal?: HwpNoteTerminal;
  cellWidthFallback?: HwpCellWidthFallback;
  floatingTextBox?: HwpFloatingTextBox;
  inlineTextBox?: { width: number; height: number; textWidth: number; pitch: number; hanging: number; size: number; family: string };
  authoredCellFace?: { size: number; pitch: number; rowHeight: number; hanging?: number; header?: true; continued?: true; scale?: number } };
export function normalizeHwpWrapText(text: string) {
  return text.normalize('NFKC').replace(/[\s\u0000-\u001f\ufffc]/gu, '');
}

export function hwpWrapFingerprint(text: string) {
  let a = 0x811c9dc5, b = 0x9e3779b9;
  for (let index = 0; index < text.length; index++) {
    const unit = text.charCodeAt(index);
    a = Math.imul(a ^ unit, 0x01000193);
    b = Math.imul(b ^ unit, 0x5bd1e995);
  }
  return (a >>> 0).toString(16).padStart(8, '0') + (b >>> 0).toString(16).padStart(8, '0');
}

function validOrigin(origin: HwpSourceOrigin | undefined): origin is HwpSourceOrigin {
  return !!origin && Number.isInteger(origin.page) && origin.page > 0 &&
    Number.isFinite(origin.x) && origin.x >= 0 && Number.isFinite(origin.y) && origin.y >= 0;
}

export function parseHwpSourceWrap(value: unknown): HwpSourceWrap | undefined {
  if (typeof value !== 'string' || value.length > 1600) return;
  try {
    const source = JSON.parse(value);
    if (source && typeof source.fingerprint === 'string' && /^[0-9a-f]{16}$/.test(source.fingerprint) &&
        Array.isArray(source.ends) && source.ends.length > 0 && source.ends.length <= 80 &&
        source.ends.every((end: unknown, index: number) => Number.isInteger(end) &&
          Number(end) > (index ? source.ends[index - 1] : 0) && Number(end) <= 20000)) {
      const { origin, cellWidth, bodyWidth, columnWidth, columnPitches, scaleRun, nestedHanging, singleCellLine, singleHeaderLine, scaledCellLine, splitCell,
        columnCellHeight, columnCellPitch, columnCellStories, expandedRowHeight, expandedRowLine, singleBodyFallback, authoredBodyLines, authoredNativeLine, authoredCellFace,
        forcedBodyEnd, inlineTextBox, cellWidthFallback } = source;
      const floatingBox = floatingTextBox(source.floatingTextBox, source.ends.length);
      if (value.length > 800 && !floatingBox) return;
      if (source.floatingTextBox !== undefined && (!floatingBox || origin || cellWidth || bodyWidth || columnWidth ||
          authoredBodyLines || authoredCellFace || forcedBodyEnd || inlineTextBox || source.noteTerminal ||
          source.ends.at(-1) > 1000)) return;
      const terminalNote = noteTerminal(source.noteTerminal);
      if (source.noteTerminal !== undefined && (!terminalNote || origin || cellWidth || bodyWidth || columnWidth ||
          authoredBodyLines || authoredCellFace || forcedBodyEnd || inlineTextBox || source.ends.length !== 2 ||
          source.ends[1] > 160)) return;
      if (singleBodyFallback !== undefined && (singleBodyFallback !== true || source.ends.length !== 1 ||
          !origin || bodyWidth === undefined || cellWidth || columnWidth !== undefined || source.ends[0] < 12 || source.ends[0] > 160)) return;
      if (!singleBodyFallback && !authoredNativeLine && !forcedBodyEnd && !inlineTextBox && !floatingBox && source.ends.length === 1 && (!cellWidth || origin !== undefined || bodyWidth !== undefined ||
          source.ends[0] < 2 || source.ends[0] > (scaleRun || authoredCellFace || cellWidthFallback ? 160 : singleCellLine || scaledCellLine ? 40 : 12))) return;
      if (origin !== undefined && !validOrigin(origin)) return;
      if (cellWidth !== undefined && (!cellWidth || !Number.isFinite(cellWidth.outer) || !Number.isFinite(cellWidth.text) ||
          cellWidth.text <= 0 || cellWidth.text > cellWidth.outer || cellWidth.outer > 10000)) return;
      if (bodyWidth !== undefined && (cellWidth || !Number.isFinite(bodyWidth) || bodyWidth <= 0 || bodyWidth > 10000)) return;
      if (columnWidth !== undefined && (origin || cellWidth || bodyWidth || source.ends.length < 2 ||
          !Number.isFinite(columnWidth) || columnWidth <= 0 || columnWidth > 10000)) return;
      if (columnPitches !== undefined && (columnWidth === undefined || source.ends.length !== 2 ||
          !Array.isArray(columnPitches) || columnPitches.length !== 2 ||
          !columnPitches.every(pitch => Number.isFinite(pitch) && pitch >= 4 && pitch <= 480) ||
          columnPitches[0] <= columnPitches[1])) return;
      if (forcedBodyEnd !== undefined && (!forcedBodyEnd || origin || cellWidth || bodyWidth || columnWidth ||
          singleBodyFallback || authoredBodyLines || source.ends.length !== 1 || source.ends[0] > 160 ||
          ![forcedBodyEnd.baseline, forcedBodyEnd.room, forcedBodyEnd.width, forcedBodyEnd.spaceBefore, forcedBodyEnd.pitch,
            forcedBodyEnd.font, forcedBodyEnd.hanging].every(value => Number.isFinite(value) && value > 0 && value <= 2000) ||
          forcedBodyEnd.baseline > forcedBodyEnd.font || forcedBodyEnd.room < forcedBodyEnd.font - .16 ||
          forcedBodyEnd.room >= forcedBodyEnd.pitch || forcedBodyEnd.pitch > forcedBodyEnd.font * 2 ||
          forcedBodyEnd.hanging >= forcedBodyEnd.width / 4 || forcedBodyEnd.spaceBefore > forcedBodyEnd.pitch ||
          !/^[0-9a-f]{16}$/.test(forcedBodyEnd.preceding) ||
          !Number.isInteger(forcedBodyEnd.precedingLines) || forcedBodyEnd.precedingLines < 2 || forcedBodyEnd.precedingLines > 8)) return;
      if (authoredBodyLines !== undefined && (!authoredBodyLines || cellWidth || columnWidth !== undefined ||
          source.ends.length < 2 || !validOrigin(authoredBodyLines.origin) ||
          !Number.isFinite(authoredBodyLines.width) || authoredBodyLines.width <= 0 || authoredBodyLines.width > 10000 ||
          !Number.isFinite(authoredBodyLines.hanging) || authoredBodyLines.hanging <= 0 ||
          authoredBodyLines.hanging >= authoredBodyLines.width ||
          !Number.isFinite(authoredBodyLines.size) || authoredBodyLines.size < 4 || authoredBodyLines.size > 120 ||
          !Number.isFinite(authoredBodyLines.pitch) || authoredBodyLines.pitch < authoredBodyLines.size ||
          authoredBodyLines.pitch > authoredBodyLines.size * 4)) return;
      if (authoredNativeLine !== undefined && (!authoredNativeLine || authoredBodyLines || cellWidth || columnWidth !== undefined ||
          forcedBodyEnd || inlineTextBox || floatingBox || source.ends.length !== 1 ||
          source.ends[0] < 12 || source.ends[0] > 160 || !validOrigin(authoredNativeLine.origin) ||
          typeof authoredNativeLine.typography !== 'string' || !/^[0-9a-f]{16}$/.test(authoredNativeLine.typography) ||
          ![authoredNativeLine.width, authoredNativeLine.hanging, authoredNativeLine.size,
            authoredNativeLine.pitch, authoredNativeLine.ascent, authoredNativeLine.top]
            .every(value => Number.isFinite(value) && value >= 0 && value <= 10000) ||
          !(authoredNativeLine.size >= 4 && authoredNativeLine.size <= 120 &&
            authoredNativeLine.pitch > authoredNativeLine.size && authoredNativeLine.pitch <= authoredNativeLine.size * 4 &&
            authoredNativeLine.hanging > 0 && authoredNativeLine.hanging < authoredNativeLine.width &&
            authoredNativeLine.ascent > 0 && authoredNativeLine.ascent <= authoredNativeLine.size))) return;
      if (scaleRun !== undefined && (!scaleRun || source.ends.length !== 1 || !cellWidth ||
          !Number.isInteger(scaleRun.start) || scaleRun.start < 1 || scaleRun.start >= 160 ||
          !Number.isInteger(scaleRun.scale) || scaleRun.scale < 50 || scaleRun.scale >= 100 ||
          !Number.isFinite(scaleRun.width) || scaleRun.width <= 0 || scaleRun.width >= cellWidth.text ||
          !Number.isFinite(scaleRun.size) || scaleRun.size < 4 || scaleRun.size > 120 ||
          typeof scaleRun.font !== 'string' || !scaleRun.font.trim() || scaleRun.font.length > 80)) return;
      if (nestedHanging !== undefined && (!nestedHanging || !cellWidth || origin || bodyWidth || scaleRun ||
          source.ends.length < 2 || !Number.isInteger(nestedHanging.scale) || nestedHanging.scale < 50 ||
          nestedHanging.scale >= 100 || !Number.isFinite(nestedHanging.hanging) || nestedHanging.hanging <= 0 ||
          nestedHanging.hanging >= cellWidth.text || !Number.isFinite(nestedHanging.size) ||
          nestedHanging.size < 4 || nestedHanging.size > 120)) return;
      if (singleCellLine !== undefined && (singleCellLine !== true || !cellWidth || scaleRun || nestedHanging ||
          source.ends.length !== 1 || source.ends[0] <= 12)) return;
      if (singleHeaderLine !== undefined && (singleHeaderLine !== true || !cellWidth || origin || bodyWidth || columnWidth ||
          scaleRun || nestedHanging || singleCellLine || scaledCellLine || authoredCellFace || cellWidthFallback ||
          inlineTextBox || floatingBox || terminalNote || source.ends.length !== 1 || source.ends[0] !== 2)) return;
      if (scaledCellLine !== undefined && (!scaledCellLine || !cellWidth || scaleRun || nestedHanging || singleCellLine ||
          source.ends.length !== 1 || !Number.isInteger(scaledCellLine.scale) || scaledCellLine.scale < 50 || scaledCellLine.scale >= 100 ||
          !Number.isFinite(scaledCellLine.size) || scaledCellLine.size < 4 || scaledCellLine.size > 120 ||
          !Number.isFinite(scaledCellLine.rowHeight) || scaledCellLine.rowHeight < scaledCellLine.size || scaledCellLine.rowHeight > 180)) return;
      if (splitCell !== undefined && (splitCell !== true || !cellWidth || origin || nestedHanging ||
          source.ends.length < 2)) return;
      if (columnCellHeight !== undefined && (!cellWidth || origin || nestedHanging || splitCell ||
          !Number.isFinite(columnCellHeight) || columnCellHeight <= 0 || columnCellHeight > 10000 ||
          columnCellPitch === undefined && source.ends.length !== 2)) return;
      if (columnCellPitch !== undefined && (columnCellHeight === undefined || source.ends.length < 3 ||
          !Number.isFinite(columnCellPitch) || columnCellPitch <= 0 || columnCellPitch * (source.ends.length - 1) >= columnCellHeight)) return;
      if (columnCellStories !== undefined && (columnCellPitch === undefined ||
          !Number.isSafeInteger(columnCellStories) || columnCellStories < 2 || columnCellStories > 8)) return;
      if (expandedRowHeight !== undefined && (!cellWidth || origin || splitCell || columnCellHeight !== undefined ||
          nestedHanging || source.ends.length < 2 || !Number.isFinite(expandedRowHeight) ||
          expandedRowHeight <= 0 || expandedRowHeight > 2000)) return;
      if (inlineTextBox !== undefined && (!inlineTextBox || forcedBodyEnd || cellWidth || origin || bodyWidth || columnWidth ||
          authoredBodyLines || authoredCellFace || expandedRowHeight || source.ends.length > 12 || source.ends.at(-1) > 1000 ||
          !Number.isFinite(inlineTextBox.width) || inlineTextBox.width < 100 || inlineTextBox.width > 2000 ||
          !Number.isFinite(inlineTextBox.height) || inlineTextBox.height < 80 || inlineTextBox.height > 1800 ||
          !Number.isFinite(inlineTextBox.textWidth) || inlineTextBox.textWidth <= 0 || inlineTextBox.textWidth >= inlineTextBox.width ||
          !Number.isFinite(inlineTextBox.hanging) || inlineTextBox.hanging < 0 || inlineTextBox.hanging >= inlineTextBox.textWidth ||
          !Number.isFinite(inlineTextBox.size) || inlineTextBox.size < 4 || inlineTextBox.size > 96 ||
          typeof inlineTextBox.family !== 'string' || !inlineTextBox.family.trim() || inlineTextBox.family.length > 100 ||
          !Number.isFinite(inlineTextBox.pitch) || inlineTextBox.pitch < inlineTextBox.size ||
          inlineTextBox.pitch * source.ends.length >= inlineTextBox.height)) return;
      if (expandedRowLine !== undefined && (!expandedRowLine || !expandedRowHeight || !cellWidth ||
          source.ends.length !== 2 || source.ends[0] < 20 || source.ends[1] - source.ends[0] > 3 ||
          !Number.isFinite(expandedRowLine.hanging) || expandedRowLine.hanging <= 0 || expandedRowLine.hanging >= cellWidth.text / 4 ||
          !Number.isFinite(expandedRowLine.pitch) || expandedRowLine.pitch <= 0 || expandedRowLine.pitch >= expandedRowHeight ||
          !Number.isInteger(expandedRowLine.scale) || expandedRowLine.scale < 90 || expandedRowLine.scale >= 100)) return;
      if (cellWidthFallback !== undefined && (!cellWidthFallback || !cellWidth || origin || bodyWidth || columnWidth ||
          nestedHanging || scaleRun || scaledCellLine || singleCellLine || splitCell ||
          source.ends.length > (cellWidthFallback.intactRow?.paragraphs === 2 ? 6 : 2) ||
          source.ends.at(-1) < 8 || source.ends.at(-1) > 160 ||
          !Number.isInteger(cellWidthFallback.scale) || cellWidthFallback.scale < 90 || cellWidthFallback.scale >= 100 ||
          !Number.isFinite(cellWidthFallback.hanging) || cellWidthFallback.hanging < 0 ||
          cellWidthFallback.hanging >= (cellWidthFallback.intactRow
            ? cellWidthFallback.intactRow.textWidth ?? cellWidth.text : cellWidth.text / 4) ||
          cellWidthFallback.intactRow !== undefined && (!cellWidthFallback.intactRow ||
            !Number.isFinite(cellWidthFallback.intactRow.height) || cellWidthFallback.intactRow.height < cellWidthFallback.pitch ||
            cellWidthFallback.intactRow.height > 10000 || !Number.isSafeInteger(cellWidthFallback.intactRow.columns) ||
            cellWidthFallback.intactRow.columns < 1 ||
            cellWidthFallback.intactRow.paragraphs !== undefined &&
              (cellWidthFallback.intactRow.paragraphs !== 2 || cellWidthFallback.hardBreak !== undefined ||
                cellWidthFallback.intactRow.textWidth === undefined || source.ends.length < 2 ||
                source.ends.length * cellWidthFallback.pitch > cellWidthFallback.intactRow.height) ||
            cellWidthFallback.intactRow.textWidth !== undefined &&
              (!Number.isFinite(cellWidthFallback.intactRow.textWidth) || cellWidthFallback.intactRow.textWidth <= 0 ||
                cellWidthFallback.intactRow.textWidth > cellWidth.outer)) ||
          !Number.isFinite(cellWidthFallback.size) || cellWidthFallback.size < 4 || cellWidthFallback.size > 120 ||
          !Number.isFinite(cellWidthFallback.pitch) || cellWidthFallback.pitch <= cellWidthFallback.size ||
          cellWidthFallback.pitch > cellWidthFallback.size * 4 ||
          cellWidthFallback.hardBreak !== undefined && (cellWidthFallback.hardBreak !== true || source.ends.length !== 2))) return;
      if (authoredCellFace !== undefined && (!authoredCellFace || expandedRowHeight !== undefined && (!expandedRowLine || !authoredCellFace.continued ||
          Math.abs(authoredCellFace.rowHeight - expandedRowHeight) > .1 ||
          Math.abs(authoredCellFace.hanging - expandedRowLine.hanging) > .1) || !cellWidth || origin || columnWidth !== undefined ||
          columnCellHeight !== undefined || nestedHanging || scaleRun || scaledCellLine || singleCellLine || splitCell ||
          source.ends.length > (authoredCellFace.continued ? 3 : 2) ||
          authoredCellFace.continued !== undefined && (authoredCellFace.continued !== true || authoredCellFace.header ||
            authoredCellFace.hanging === undefined) || source.ends.at(-1) < (authoredCellFace.header ? 2 : 12) || source.ends.at(-1) > 160 ||
          authoredCellFace.scale !== undefined && (!authoredCellFace.continued || ![2, 3].includes(source.ends.length) ||
            !Number.isInteger(authoredCellFace.scale) || authoredCellFace.scale < 90 || authoredCellFace.scale >= 100) ||
          authoredCellFace.header !== undefined && (authoredCellFace.header !== true || authoredCellFace.hanging !== undefined ||
            source.ends.length !== 1 || source.ends[0] !== 2) ||
          authoredCellFace.hanging !== undefined && (!authoredCellFace.continued && source.ends.length !== 1 ||
            !Number.isFinite(authoredCellFace.hanging) ||
            authoredCellFace.hanging <= 0 || authoredCellFace.hanging > cellWidth.text / 4) ||
          !Number.isFinite(authoredCellFace.size) || authoredCellFace.size < 4 || authoredCellFace.size > 120 ||
          !Number.isFinite(authoredCellFace.pitch) || authoredCellFace.pitch <= authoredCellFace.size ||
          authoredCellFace.pitch > authoredCellFace.size * 4 || !Number.isFinite(authoredCellFace.rowHeight) ||
          authoredCellFace.rowHeight < authoredCellFace.pitch || authoredCellFace.rowHeight > 2000)) return;
      return { fingerprint: source.fingerprint, ends: source.ends,
        ...(terminalNote && { noteTerminal: terminalNote }),
        ...(floatingBox && { floatingTextBox: floatingBox }),
        ...(forcedBodyEnd && { forcedBodyEnd: { baseline: forcedBodyEnd.baseline, room: forcedBodyEnd.room,
          width: forcedBodyEnd.width, spaceBefore: forcedBodyEnd.spaceBefore, pitch: forcedBodyEnd.pitch, font: forcedBodyEnd.font,
          hanging: forcedBodyEnd.hanging, preceding: forcedBodyEnd.preceding, precedingLines: forcedBodyEnd.precedingLines } }),
        ...(inlineTextBox && { inlineTextBox: { width: inlineTextBox.width, height: inlineTextBox.height,
          textWidth: inlineTextBox.textWidth, pitch: inlineTextBox.pitch, hanging: inlineTextBox.hanging,
          size: inlineTextBox.size, family: inlineTextBox.family } }),
        ...(expandedRowHeight !== undefined && { expandedRowHeight }),
        ...(expandedRowLine && { expandedRowLine: { hanging: expandedRowLine.hanging, pitch: expandedRowLine.pitch,
          scale: expandedRowLine.scale } }),
        ...(cellWidthFallback && { cellWidthFallback: { scale: cellWidthFallback.scale,
          hanging: cellWidthFallback.hanging, size: cellWidthFallback.size, pitch: cellWidthFallback.pitch,
          ...(cellWidthFallback.hardBreak === true && { hardBreak: true as const }),
          ...(cellWidthFallback.intactRow && { intactRow: {
            height: cellWidthFallback.intactRow.height, columns: cellWidthFallback.intactRow.columns,
            ...(cellWidthFallback.intactRow.paragraphs === 2 && { paragraphs: 2 as const }),
            ...(cellWidthFallback.intactRow.textWidth !== undefined && { textWidth: cellWidthFallback.intactRow.textWidth }) } }) } }),
        ...(authoredCellFace && { authoredCellFace: { size: authoredCellFace.size,
          pitch: authoredCellFace.pitch, rowHeight: authoredCellFace.rowHeight,
          ...(authoredCellFace.hanging !== undefined && { hanging: authoredCellFace.hanging }),
          ...(authoredCellFace.header === true && { header: true as const }),
          ...(authoredCellFace.continued === true && { continued: true as const }),
          ...(authoredCellFace.scale !== undefined && { scale: authoredCellFace.scale }) } }),
        ...(columnCellHeight !== undefined && { columnCellHeight }),
        ...(singleBodyFallback && { singleBodyFallback: true as const }),
        ...(columnCellPitch !== undefined && { columnCellPitch }),
        ...(columnCellStories !== undefined && { columnCellStories }),
        ...(singleCellLine && { singleCellLine: true as const }),
        ...(singleHeaderLine && { singleHeaderLine: true as const }),
        ...(scaledCellLine && { scaledCellLine: { scale: scaledCellLine.scale, size: scaledCellLine.size,
          rowHeight: scaledCellLine.rowHeight } }),
        ...(origin && { origin: { page: origin.page, x: origin.x, y: origin.y } }),
        ...(authoredBodyLines && { authoredBodyLines: {
          origin: { page: authoredBodyLines.origin.page, x: authoredBodyLines.origin.x, y: authoredBodyLines.origin.y },
          width: authoredBodyLines.width, hanging: authoredBodyLines.hanging,
          size: authoredBodyLines.size, pitch: authoredBodyLines.pitch } }),
        ...(authoredNativeLine && { authoredNativeLine: {
          origin: { page: authoredNativeLine.origin.page, x: authoredNativeLine.origin.x, y: authoredNativeLine.origin.y },
          width: authoredNativeLine.width, hanging: authoredNativeLine.hanging, size: authoredNativeLine.size,
          pitch: authoredNativeLine.pitch, ascent: authoredNativeLine.ascent, top: authoredNativeLine.top,
          typography: authoredNativeLine.typography } }),
        ...(cellWidth && { cellWidth: { outer: cellWidth.outer, text: cellWidth.text } }),
        ...(splitCell && { splitCell: true as const }),
        ...(bodyWidth !== undefined && { bodyWidth }),
        ...(columnWidth !== undefined && { columnWidth }),
        ...(columnPitches && { columnPitches: [columnPitches[0], columnPitches[1]] as [number, number] }),
        ...(nestedHanging && { nestedHanging: { scale: nestedHanging.scale, hanging: nestedHanging.hanging, size: nestedHanging.size } }),
        ...(scaleRun && { scaleRun: { start: scaleRun.start, scale: scaleRun.scale, width: scaleRun.width,
          size: scaleRun.size, font: scaleRun.font } }) };
    }
  } catch { /* Optional editing metadata does not prevent opening the file. */ }
}
