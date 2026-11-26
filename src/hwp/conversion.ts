import { hwpNativeGrid, hwpNativeGridAttribute } from '../hwp-native-grid';
import { restoreHwpExportTextBoxes, restoreNativeBoxPaint, type NativeBoxPaint, type HwpExportTextBox } from './export-textboxes';
import { sourceFloatingNestedTextBox, sourceInlineTextBox } from './source-inline-textbox';
import { hwpRelatedTableLayoutAttribute } from '../hwp-related-table-layout';
import { sourceRelatedTableLayout } from './source-related-table-layout';
import { hwpNestedTableLayoutAttribute } from '../hwp-nested-table-layout';
import { sourceNestedTableLayout } from './source-nested-table-layout';
import { boxedDigitHtml } from '../hwp-boxed-digit';
import { circledCompositionHtml } from './circled-composition';
import { characterScaleProperty } from '../character-scale';
import { hwpScaledCellFace, hwpShortCellLabel, hwpSourceWrapAttribute, hwpWrapFingerprint, normalizeHwpWrapText, parseHwpSourceWrap, type HwpSourceWrap } from '../hwp-text-advances';
import type { HwpDocument } from '@rhwp/core';
import { activeHwpHorizontalMargin, parseHwpHorizontalMargins, type PageSettings } from '../types';
import documentLimits from '../../electron/document-limits.json';
import type { HwpProgressReporter } from './progress';
import type { HwpLayoutMetadata } from './layout-metadata';
import { cellContents, restoreCellContent, restoreHardLineBreaks, entities, type CellPath } from './cell-content';
import { serializedParts } from './serialized-parts';
import { olePreview } from './ole-preview';
import { objectImageSizes } from './object-image-sizes';
import { headerDefinitions, headerPreviewPage, headerLayer, hiddenStories, editableHeaderStories, initialBlankHeaderStory, initialPageNumberStory } from './header-stories';
import { stageHwpFields } from './export-fields';
import { importHwpNotes, stageHwpNotes } from './note-stories';
import { hwpFootnoteLayoutAttribute, hwpFootnoteMarkAttribute, parseHwpFootnoteLayout } from '../hwp-footnote-layout';
import { stageHwpColumns } from './export-columns';
import { stageHwpSections } from './export-sections';
import { stageHwpFurniture } from './export-furniture';
import { stageHwpPictures } from './export-pictures';
import { hwpBookmark, hwpExplicitTabStops, hwpTabAttribute, stageHwpTextNavigation } from './text-navigation';
import { hwpTableGrid } from './table-grid';
import { sourceCentersContinuedBorders, sourceCentersMultilineRowBorders, sourceCentersTableBorders } from './source-centered-table-borders';
import { sourceAlignedHeaderEnds, sourceContinuedRowReserve, sourceEmptyMergeCuts, sourceFirstTableOverhang, sourceIntactTopRow, sourceTableContinuation } from './source-table-continuation';
import { hwpIntactRowAttribute, type HwpIntactRow } from '../hwp-intact-row';
import { sourceCenteredCellLabels, sourceCompactContinuedCell } from './source-centered-cell-labels';
import { sourceCellCaptionLine, sourceScaledCellLine, type NativeCellInkRun } from './source-caption-cell-line';
import { sourceHwpCompactCellFace, type CompactCellFacePosition } from './source-cell-face';
import { encodeHwpCellRasterFill, hwpCellRasterFillAttribute, type HwpCellRasterPicture } from '../hwp-cell-raster-fill';
import { hwpCenteredCellLabelsAttribute } from '../hwp-centered-cell-labels';
import { sourceHwpBodyEndFit, sourceHwpColumnEndFit, sourceHwpForcedBodyEnd } from './source-column-end';
import { sourceFootnoteTerminal } from './source-footnote-terminal';
import { sourceCellTerminal, sourceTableTerminalBoxes } from './source-cell-terminal';
import { nestedTableGeometry } from './nested-tables';
import { hwpCellBorders } from './cell-borders';
import { physicalLineHeight, paragraphCss, sourceLinePitches, sourcePreciseColumnPitch, cellParagraphCss, paragraphAutoSpace, hwpBodyTextAtLineTop, sourceStoredMixedLine, sourceMixedBodyLinePair, sourceUniformMixedLinePair, sourceTerminalTextAtLineTop, sourceTextAfterInlineTableFits, sourceInitialTableLead, sourceHostBeforeFloatingTable, sourceHostOnPriorPage, sourceTerminalEmptyLine, sourceTerminalEmptyLineAfterTable, sourceInlineStrokeHeight, sourceKeptInlineTable, sourceKeptInlineTableAfterBody, sourceMixedCellLeading, sourceBlankAutoCellFits, sourceIntactCellBefore, sourceSplitCellBefore, sourcePaintedCellBefore, sourceWholeParagraphPeerRow, sourceInlinePeerInMiddleFragment, sourceColumnCellWrap, sourceBeforeNestedCellTail, sourceCellBoundsTolerance, sourceScaledCellRun, sourceBodyPictureFlow, sourcePaperHeadingLine, sourceTightInlinePicture, sourceInlineColumnTableIndent, type HwpPaintedLine, connectedHwpBorderGroup } from './paragraph-layout';
import { nativeLineTypography, preserveHwpBodyLineEnds, preserveHwpColumnLineEnds, preserveHwpSourceSpaces, restoreUniformLatinMetrics, restoreScriptFonts, sourceFontRuns, type HwpScriptTypography } from './script-fonts';
import { restoreHwpFixedSpaces } from './fixed-spaces';
import { metafilePreview } from './metafile-preview';
import { emptyLineCompatibilityWarning } from '../page-empty-lines';
import { encodeHwpBodyLine } from '../hwp-body-line';
import { restoreNativeContinuationPlacement, restoreNativePaperHeaderTable, restoreNativeRasterTablePair } from './export-table-placement';
import { restoreStandalonePictureLines } from './export-picture-line';
import { repairBinaryTerminalEmptyPage, restorePastedEmptyParagraphs, restoreTerminalEmptyFonts, withoutPastedEmptyCarriers } from './export-empty-paragraph-format';
import { restoreAuthoredNativeSingleLine, restoreAuthoredTerminalWrap, restoreTerminalPageBreak } from './export-terminal-page-break';
import { restoreNativeEmptyTableRows } from './export-empty-table-rows';
import { restoreNativePageHead } from './export-page-head';
import { sourceHwpPageHead } from './source-page-head';
import { stageHwpTables } from './export-tables';
import { tableCaptionDefinitions, renderTableCaptions, restoreTableCaptions } from './table-captions';
import { hwpInlineTableTop, hwpInlineTableTopAttribute, hwpTableTopGap } from '../table-anchor-data';
import { hwpMarginCenteredTableAnchor } from './table-position';
import { elementRanges, topLevelHwpPasteBlocks } from './markup-ranges';
import { stageHwpCharacterStyles } from './export-character-width';
import { planHwpGridRows } from './export-grid-rows';
import { repairBinaryDuplicatePage } from './export-binary-pages';
import { restoreNativeTableLead, type NativePrintPlan } from './export-table-lead';

export interface EditableHwp { html: string; page: PageSettings; warnings: string[]; editorPages?: number }
export const HWP_CONVERSION_NOTICE = '본문, 기본 글자·문단 서식, 표와 그림을 변환합니다. 한글과 Word의 조판 방식이 달라 줄바꿈과 쪽 수가 달라질 수 있습니다. 원본 파일은 유지하고 변환한 문서는 새 파일로 저장하세요.';
// Expanded HTML includes base64 pictures, so a 21 MB image document already
// needs about 28 MB before markup. This is separate from the source-file cap.
export const HWP_HTML_LIMIT = documentLimits.hwpEditableHtmlBytes;
export function hwpHtmlBytes(value: string): number {
  // Do not allocate another document-sized buffer just to measure the Worker input.
  let bytes = value.length;
  for (let i = 0; i < value.length; i++) {
    const c = value.charCodeAt(i);
    if (c <= 0x7f) continue;
    if (c <= 0x7ff) bytes++;
    else {
      bytes += 2;
      if (c >= 0xd800 && c <= 0xdbff && i + 1 < value.length) {
        const next = value.charCodeAt(i + 1);
        if (next >= 0xdc00 && next <= 0xdfff) i++;
      }
    }
  }
  return bytes;
}
export function checkHwpHtmlSize(bytes: number) {
  if (bytes > HWP_HTML_LIMIT) throw new Error(`편집용 변환 데이터가 약 ${Math.ceil(bytes / 1024 / 1024 * 10) / 10} MiB로 한도 ${HWP_HTML_LIMIT / 1024 / 1024} MiB를 넘습니다. 원본 파일 크기가 아니라 압축을 푼 본문·그림을 편집 형식으로 변환한 크기입니다. 그림을 줄이거나 문서를 나눠 주세요.`);
}
const escape = (text: string) => text.replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]!);
const fragment = (html: string) => html.match(/<!--StartFragment-->([\s\S]*?)<!--EndFragment-->/)?.[1].trim() ?? html;
type Control = { list: number; para: number; controlIndex: number; pos: number; ctrlId: string; userDesc: string; props: { TreatAsChar?: number } };
type PageDef = { width: number; height: number; landscape: boolean; marginTop: number; marginRight: number; marginBottom: number; marginLeft: number; marginHeader: number; marginFooter: number; marginGutter: number };
const mm = (value: number) => Math.round(value * 25.4 / 7200 * 100) / 100;
const hu = (value: number) => Math.round(value * 7200 / 25.4);
const px = (value: number) => Math.round(value / 75 * 100) / 100;
const cssStyle = (tag: string, css: string) => tag.includes('style="') ? tag.replace(/style="([^"]*)"/, (_, old) => `style="${old};${css}"`) : tag.replace(/>$/, ` style="${css}">`);


function tableGeometry(doc: HwpDocument, section: number, para: number, control: number, html: string, pitches: Map<string, number>, splitCellLines: Map<string, HwpPaintedLine[]>, nextPageBreak: boolean, tabs?: (cell: number, paragraph: number) => string, hwpxFontSpace?: readonly boolean[], rasterFills?: Readonly<Record<number, HwpCellRasterPicture>>): string {
  const dims = JSON.parse(doc.getTableDimensions(section, para, control));
  if (dims.cellCount > 50000 || dims.colCount > 1000 || dims.rowCount > 10000) throw new Error('편집할 표의 셀 수가 너무 많습니다.');
  const tableProperties = JSON.parse(doc.getTableProperties(section, para, control));
  const cells = Array.from({ length: dims.cellCount }, (_, index) => ({
    index, ...JSON.parse(doc.getCellInfo(section, para, control, index)), ...JSON.parse(doc.getCellProperties(section, para, control, index)),
  })).map(cell => {
    // Stored cell padding is inactive unless hasMargin is enabled. Inherited
    // table padding can be much smaller; using the inactive values wraps
    // short headings into extra lines and increases every following row.
    for (const side of ['Left', 'Right', 'Top', 'Bottom']) {
      const key = 'padding' + side;
      if (!cell.applyInnerMargin || Number(cell[key]) < 0) cell[key] = Math.max(0, Number(tableProperties[key]) || 0);
    }
    return cell;
  }).sort((a, b) => a.row - b.row || a.col - b.col);
  // Solve cumulative grid edges from both merged and unmerged cells. Dividing
  // every merged width equally expands the fine grids used by Hancom forms.
  const grid = (axis: 'col' | 'row', length: number) => {
    if (axis === 'row' && cells.some(c => c.height === 0)) {
      // Zero-height cells are automatic rows, not zero-length grid edges.
      // Pinning such an edge puts all remaining merged-cell height in the
      // final row, then CSS grows the earlier automatic rows again.
      const heights: (number | undefined)[] = Array.from({ length }, (_, row) => {
        const values = cells.filter(c => c.row === row && c.rowSpan === 1 && c.height > 0).map(c => c.height);
        return values.length ? Math.max(...values) : undefined;
      });
      const constraints = cells.filter(c => c.rowSpan > 1 && c.height > 0)
        .sort((a, b) => a.rowSpan - b.rowSpan).map(c => ({ start: c.row, count: c.rowSpan, height: c.height }));
      constraints.push({ start: 0, count: length, height: Math.max(0, Number(tableProperties.tableHeight) || 0) });
      for (const constraint of constraints) {
        const indexes = Array.from({ length: constraint.count }, (_, i) => constraint.start + i).filter(i => i < length);
        const unknown = indexes.filter(i => heights[i] === undefined);
        if (!unknown.length || !constraint.height) continue;
        const remaining = Math.max(0, constraint.height - indexes.reduce((sum, i) => sum + (heights[i] ?? 0), 0));
        for (const i of unknown) heights[i] = remaining / unknown.length;
      }
      return heights.map(height => Math.max(1, px(height ?? 0)));
    }
    return hwpTableGrid(cells, axis, length);
  };
  const columns = grid('col', dims.colCount), rows = grid('row', dims.rowCount);
  // CELL (native value 2) explicitly permits breaking within a cell, even
  // if the stored row currently fits one sheet. An edit/font change must not
  // turn that permission into an atomic DOCX row. For other policies retain
  // observed source continuations; repeated headers are not such evidence.
  // Inline tables do not activate the multiple-page option. Keep native
  // observed continuations, but do not grant a new split to an inline form.
  const declaredCellBreak = Number(tableProperties.pageBreak) === 2 && !tableProperties.treatAsChar;
  const continuationRows = new Set<number>();
  const observedSplitRows = new Set<number>();
  const rowReserves = new Map<number, number>();
  const mergesByRow = new Map<number, number[]>();
  const cellsByRow = new Map<number, typeof cells>();
  for (const cell of cells) {
    const own = cellsByRow.get(cell.row) ?? [];
    own.push(cell); cellsByRow.set(cell.row, own);
  }
  for (const cell of cells) if (cell.rowSpan > 1) {
    for (let row = cell.row; row < Math.min(dims.rowCount, cell.row + cell.rowSpan); row++) {
      const merges = mergesByRow.get(row) ?? [];
      merges.push(cell.index);
      mergesByRow.set(row, merges);
    }
  }
  const pageBoxes = JSON.parse(doc.getTableCellBboxes(section, para, control)) as { cellIdx: number; pageIndex: number; row: number; x: number; y: number; h: number; w: number }[];
  const nestedLayout = declaredCellBreak && !tableProperties.hasCaption && dims.cellCount === 1 &&
    dims.rowCount === 1 && dims.colCount === 1 && !cells[0].isHeader && !cells[0].textDirection
    ? sourceNestedTableLayout(doc, section, para, control, pageBoxes) : undefined;
  const terminalBoxes = sourceTableTerminalBoxes(tableProperties, pageBoxes);
  const authoredCellWidths = new Map<number, number>();
  const cellPages = new Map<number, Set<number>>();
  const boxesByRow = new Map<number, typeof pageBoxes>();
  for (const box of pageBoxes) {
    const width = authoredCellWidths.get(box.cellIdx);
    if (width === undefined) authoredCellWidths.set(box.cellIdx, box.w);
    else if (Math.abs(width - box.w) > .1) authoredCellWidths.set(box.cellIdx, NaN);
    const pages = cellPages.get(box.cellIdx) ?? new Set<number>(); pages.add(box.pageIndex); cellPages.set(box.cellIdx, pages);
    const own = boxesByRow.get(box.row) ?? [];
    own.push(box); boxesByRow.set(box.row, own);
  }
  const pageBodies = new Map<number, { columns: number | undefined; top: number; bottom: number }>();
  const sourceBody = (pageIndex: number) => {
    let body = pageBodies.get(pageIndex);
    if (body) return body;
    const page = JSON.parse(doc.getPageInfo(pageIndex));
    body = { columns: page.columns?.length, top: Number(page.marginTop) + Number(page.marginHeader),
      bottom: Number(page.footerArea?.y ?? (page.height - page.marginBottom - page.marginFooter)) };
    pageBodies.set(pageIndex, body);
    return body;
  };
  if (declaredCellBreak && !tableProperties.hasCaption && tableProperties.cellSpacing === 0 &&
      tableProperties.restrictInPage && !tableProperties.allowOverlap && tableProperties.textWrap === 'TopAndBottom') {
    for (const [row, merged] of mergesByRow) {
      if (merged.length !== 1) continue;
      const owner = cells.find(cell => cell.index === merged[0]);
      const peers = cellsByRow.get(row);
      const boxes = boxesByRow.get(row);
      if (!owner || owner.row >= row || owner.row + owner.rowSpan !== row + 1 ||
          cellPages.get(owner.index)?.size !== 2 || peers?.length !== 2 || boxes?.length !== 2 ||
          peers.some(cell => cell.isHeader || cell.rowSpan !== 1 || cell.colSpan !== 1 || cell.height <= 0)) continue;
      const [head, tail] = boxes;
      const height = px(peers[0].height), paper = sourceBody(head.pageIndex);
      if (peers[0].height !== peers[1].height || height <= rows[row] + .1 || paper.columns !== 1 ||
          head.y < paper.top - .1 || head.y + height > paper.bottom + .1 ||
          head.pageIndex !== tail.pageIndex || Math.abs(head.y - tail.y) > .1 ||
          boxes.some(box => Math.abs(box.h - height) > .1)) continue;
      const fragment = pageBoxes.find(box => box.cellIdx === owner.index && box.pageIndex === head.pageIndex);
      // The authored height of an intact ordinary row is independent of
      // the short stored height on the earlier spanning cell.
      if (fragment && fragment.y < head.y && Math.abs(fragment.y + fragment.h - head.y - height) <= .1)
        rows[row] = height;
    }
  }
  const columnCellBox = dims.cellCount === 1 && dims.rowCount === 1 && dims.colCount === 1 && pageBoxes.length === 1 &&
    tableProperties.restrictInPage && !tableProperties.allowOverlap && !tableProperties.hasCaption &&
    !tableProperties.cellSpacing ? pageBoxes[0] : undefined;
  const columnCellBody = columnCellBox ? sourceBody(columnCellBox.pageIndex) : undefined;
  const intactColumnCell = columnCellBox && columnCellBody?.columns === 2 && columnCellBox.y >= columnCellBody.top - .1 &&
    columnCellBox.y + columnCellBox.h <= columnCellBody.bottom + .1 ? columnCellBox : undefined;
  const expandableCells = new Map<number, typeof pageBoxes[number]>();
  const tablePage = pageBoxes[0]?.pageIndex;
  const cellRasterFills = new Map<number, string>();
  const intactRasterGrid = () => {
    if (dims.rowCount === 1) return dims.colCount === dims.cellCount;
    // An inline form can paint independently only when every row has the
    // same measured outside edges and no cells run through row boundaries.
    if (!tableProperties.treatAsChar || cells.some(cell => cell.rowSpan !== 1) ||
        !pageBoxes.every(box => box.pageIndex === tablePage) || cellPages.size !== cells.length) return false;
    const close = (a: number, b: number) => Math.abs(a - b) <= sourceCellBoundsTolerance;
    let left = 0, right = 0, bottom = 0;
    for (let row = 0; row < dims.rowCount; row++) {
      const band = [...(boxesByRow.get(row) ?? [])].sort((a, b) => a.x - b.x);
      const first = band[0], last = band.at(-1);
      if (!first || !last || band.some((box, i) =>
        box.w <= 0 || box.h <= 0 || !close(first.y, box.y) || !close(first.h, box.h) ||
        i > 0 && !close(box.x, band[i - 1].x + band[i - 1].w))) return false;
      if (row === 0) {
        left = first.x;
        right = last.x + last.w;
      } else if (!close(first.x, left) || !close(last.x + last.w, right) || !close(first.y, bottom)) return false;
      bottom = first.y + first.h;
    }
    return true;
  };
  if (rasterFills && tablePage !== undefined && pageBoxes.length === cells.length &&
      !tableProperties.hasCaption && !tableProperties.allowOverlap && !tableProperties.cellSpacing && intactRasterGrid()) {
    const body = sourceBody(tablePage);
    for (const cell of cells) {
      const picture = rasterFills[cell.borderFillId], box = pageBoxes.find(box => box.cellIdx === cell.index);
      if (!picture || !box || box.pageIndex !== tablePage || box.y < body.top - sourceCellBoundsTolerance ||
          box.y + box.h > body.bottom + sourceCellBoundsTolerance || box.w < 1 || box.h < 1 || box.w >= 10000 || box.h >= 10000 ||
          cell.fillType !== 'none' || cell.isHeader || cell.rowSpan !== 1 || cell.textDirection ||
          cell.diagonalSlash || cell.diagonalBackSlash || Math.abs(box.w - cell.width / 75) > sourceCellBoundsTolerance) continue;
      const paragraphCount = doc.getCellParagraphCount(section, para, control, cell.index);
      if (paragraphCount !== 1) {
        // A horizontal merge can carry text over an intact picture fill.
        // Require each original text paragraph to paint inside this one
        // physical cell. Empty stored paragraphs keep their ordinary flow.
        if (paragraphCount < 2 || dims.rowCount < 2 || cell.colSpan < 2 ||
            box.h + sourceCellBoundsTolerance < cell.height / 75) continue;
        let hasText = false;
        const intactText = Array.from({ length: paragraphCount }, (_, index) => index).every(index => {
          const address = JSON.stringify([{ controlIndex: control, cellIndex: cell.index, cellParaIndex: index }]);
          const metrics = JSON.parse(doc.getParagraphLineMetrics(section, para, address));
          if (!metrics.stored || metrics.controlCount || !metrics.lines?.length) return false;
          const text = doc.getTextInCellByPath(section, para, address, 0, 0xffffffff);
          const ink = splitCellLines.get(`${section}:${para}:${control}:${cell.index}:${index}`);
          if (!text.trim()) return metrics.lines.length === 1 && !ink?.length;
          hasText = true;
          return !!ink && ink.length === metrics.lines.length && ink.every(line => line.page === tablePage &&
            line.y >= box.y - sourceCellBoundsTolerance && line.y + line.h <= box.y + box.h + sourceCellBoundsTolerance);
        });
        if (intactText && hasText)
          cellRasterFills.set(cell.index, encodeHwpCellRasterFill(box.w, box.h, picture.source, picture.brightness));
        continue;
      }
      const address = JSON.stringify([{ controlIndex: control, cellIndex: cell.index, cellParaIndex: 0 }]);
      const metrics = JSON.parse(doc.getParagraphLineMetrics(section, para, address));
      if (cell.colSpan !== 1 || Math.abs(box.h - cell.height / 75) > sourceCellBoundsTolerance) {
        // Empty horizontal frames may stretch to a neighbouring cell's
        // proven row height. The cell itself still contains no drawable text.
        if (dims.rowCount < 2 || !metrics.stored || box.h + sourceCellBoundsTolerance < cell.height / 75 ||
            doc.getTextInCellByPath(section, para, address, 0, 0xffffffff).trim()) continue;
      }
      if (!metrics.controlCount && metrics.lines?.length === 1)
        cellRasterFills.set(cell.index, encodeHwpCellRasterFill(box.w, box.h, picture.source, picture.brightness));
    }
  }
  if (hwpxFontSpace && tablePage !== undefined && tableProperties.treatAsChar && !tableProperties.allowOverlap &&
      !tableProperties.hasCaption && !tableProperties.cellSpacing && dims.colCount >= 2 && dims.colCount <= 6 &&
      pageBoxes.length === cells.length && pageBoxes.every(box => box.pageIndex === tablePage) &&
      sourceBody(tablePage).columns === 1) {
    const body = sourceBody(tablePage);
    for (const [row, rowCells] of cellsByRow) {
      const boxes = boxesByRow.get(row), first = boxes?.[0];
      if (!first || boxes.length !== rowCells.length || rowCells.length !== dims.colCount ||
          mergesByRow.has(row) || first.y < body.top - sourceCellBoundsTolerance ||
          first.y + first.h > body.bottom + sourceCellBoundsTolerance ||
          first.h <= Math.max(...rowCells.map(cell => cell.height / 75)) + sourceCellBoundsTolerance ||
          rowCells.some(cell => cell.rowSpan !== 1 || cell.colSpan !== 1 || cell.isHeader || cell.textDirection ||
            cell.diagonalSlash || cell.diagonalBackSlash ||
            !boxes.some(box => box.cellIdx === cell.index && Math.abs(box.y - first.y) <= .1 && Math.abs(box.h - first.h) <= .1))) continue;
      for (const box of boxes) {
        if (rowCells.find(cell => cell.index === box.cellIdx)?.verticalAlign === 1) expandableCells.set(box.cellIdx, box);
      }
    }
  }
  const completeExpandedCells = new Map<number, typeof pageBoxes[number]>();
  if (declaredCellBreak && tableProperties.repeatHeader && !tableProperties.hasCaption &&
      tableProperties.restrictInPage && !tableProperties.allowOverlap && !tableProperties.cellSpacing &&
      dims.colCount >= 2 && dims.colCount <= 6) {
    for (const [row, rowCells] of cellsByRow) {
      const boxes = boxesByRow.get(row), first = boxes?.[0];
      if (!first || rowCells.length !== dims.colCount || boxes.length !== rowCells.length || mergesByRow.has(row) ||
          rowCells.some(cell => cell.rowSpan !== 1 || cell.colSpan !== 1 || cell.isHeader || cell.textDirection ||
            cell.diagonalSlash || cell.diagonalBackSlash || !boxes.some(box => box.cellIdx === cell.index))) continue;
      const body = sourceBody(first.pageIndex);
      if (body.columns !== 1 || first.y < body.top - sourceCellBoundsTolerance ||
          first.y + first.h > body.bottom + sourceCellBoundsTolerance ||
          first.h <= Math.max(...rowCells.map(cell => cell.height / 75)) + sourceCellBoundsTolerance ||
          boxes.some(box => box.pageIndex !== first.pageIndex || Math.abs(box.y - first.y) > .1 ||
            Math.abs(box.h - first.h) > .1 || cellPages.get(box.cellIdx)?.size !== 1)) continue;
      for (const box of boxes)
        if (rowCells.find(cell => cell.index === box.cellIdx)?.verticalAlign === 1) completeExpandedCells.set(box.cellIdx, box);
    }
  }
  const automaticRows = new Set<number>();
  if (tableProperties.treatAsChar && !tableProperties.hasCaption && dims.colCount === 1 && dims.cellCount === dims.rowCount) {
    for (const cell of cells) {
      const boxes = boxesByRow.get(cell.row);
      if (cell.height !== 0 || cell.isHeader || cell.rowSpan !== 1 || cell.colSpan !== 1 ||
          cell.textDirection || cell.diagonalSlash || cell.diagonalBackSlash || boxes?.length !== 1) continue;
      const [box] = boxes, paper = sourceBody(box.pageIndex);
      if (box.cellIdx !== cell.index || paper.columns !== 1 || !(box.y >= paper.top - sourceCellBoundsTolerance &&
          box.y + box.h <= paper.bottom + sourceCellBoundsTolerance && rows[cell.row] - box.h > sourceCellBoundsTolerance)) continue;
      const width = box.w - (cell.paddingLeft + cell.paddingRight) / 75;
      const height = box.h - (cell.paddingTop + cell.paddingBottom) / 75;
      if (sourceBlankAutoCellFits(doc, [section, para, control, cell.index], width, height)) automaticRows.add(cell.row);
    }
  }
  // The source can finish a vertically merged pair at the page edge with a
  // smaller terminal minimum than the unpaginated grid. Require the complete
  // painted pair, its ordinary neighbour and the repeated header on the
  // following page before carrying this native minimum into the editable row.
  if (declaredCellBreak && tableProperties.repeatHeader && !tableProperties.hasCaption &&
      tableProperties.cellSpacing === 0 && tableProperties.textWrap === 'TopAndBottom' &&
      tableProperties.vertRelTo === 'Para' && tableProperties.vertAlign === 'Top' && !tableProperties.vertOffset &&
      tableProperties.horzRelTo === 'Para' && tableProperties.restrictInPage && !tableProperties.allowOverlap &&
      dims.colCount === 2 && cellsByRow.get(0)?.length === 2 && cellsByRow.get(0)?.every(cell => cell.isHeader)) {
    const close = (a: number, b: number, tolerance = sourceCellBoundsTolerance) => Math.abs(a - b) <= tolerance;
    for (const leaf of cells) {
      const row = leaf.row, top = cellsByRow.get(row - 1), next = cellsByRow.get(row + 1);
      if (row < 2 || leaf.rowSpan !== 1 || leaf.colSpan !== 1 || cellsByRow.get(row)?.length !== 1 ||
        top?.length !== 2 || next?.length !== 2 || next.some(cell => cell.rowSpan !== 1 || cell.colSpan !== 1)) continue;
      const owner = top.find(cell => cell.rowSpan === 2 && cell.col !== leaf.col);
      const pair = top.find(cell => cell.rowSpan === 1 && cell.col === leaf.col);
      if (!owner || !pair || [owner, pair, leaf].some(cell => cell.isHeader || cell.colSpan !== 1 ||
        cell.verticalAlign !== 1 || cell.textDirection || cell.diagonalSlash || cell.diagonalBackSlash ||
        cellPages.get(cell.index)?.size !== 1)) continue;
      const fragments = [owner, pair, leaf].map(cell => boxesByRow.get(cell.row)?.filter(box => box.cellIdx === cell.index));
      if (fragments.some(boxes => boxes?.length !== 1)) continue;
      const [span, head, tail] = fragments.map(boxes => boxes![0]);
      if (head.pageIndex !== tail.pageIndex || span.pageIndex !== tail.pageIndex ||
        !close(span.y, head.y) || !close(head.y + head.h, tail.y) || !close(span.y + span.h, tail.y + tail.h) ||
        !close(px(pair.height), head.h, .2) || !(tail.h > 0 && px(leaf.height) - tail.h > .5 &&
          px(leaf.height) - tail.h < tail.h * .2 && px(owner.height) - span.h > .5 && rows[row] > tail.h + .5)) continue;
      const repeat = boxesByRow.get(0)?.filter(box => box.pageIndex === tail.pageIndex + 1);
      const nextBoxes = boxesByRow.get(row + 1)?.filter(box => box.pageIndex === tail.pageIndex + 1);
      if (repeat?.length !== 2 || nextBoxes?.length !== 2 || !repeat.every(box =>
        nextBoxes.every(nextBox => close(box.y + box.h, nextBox.y)))) continue;
      const body = sourceBody(tail.pageIndex), following = sourceBody(tail.pageIndex + 1);
      const exit = body.bottom - tail.y - tail.h;
      if (body.columns !== 1 || following.columns !== 1 || span.y < body.top - sourceCellBoundsTolerance ||
        !(exit >= -sourceCellBoundsTolerance && exit <= 5)) continue;
      const complete = [owner, pair, leaf].every((cell, index) => {
        const { pageIndex, y, h } = [span, head, tail][index];
        const count = doc.getCellParagraphCount(section, para, control, cell.index);
        if (!count) return false;
        return Array.from({ length: count }, (_, p) => p).every(p => {
          const address = JSON.stringify([{ controlIndex: control, cellIndex: cell.index, cellParaIndex: p }]);
          const metrics = JSON.parse(doc.getParagraphLineMetrics(section, para, address));
          const ink = splitCellLines.get(`${section}:${para}:${control}:${cell.index}:${p}`);
          return metrics.stored && !metrics.tables?.length && ink?.length === metrics.lines?.length &&
            !!ink?.length && ink.every(line => line.page === pageIndex && line.y >= y - sourceCellBoundsTolerance &&
              line.y + line.h <= y + h + sourceCellBoundsTolerance);
        });
      });
      if (complete) rows[row] = Math.round(tail.h * 100) / 100;
    }
  }
  // A control can have a total authored cell height longer than the two
  // fragments the source actually painted before the following explicit break.
  let observedContainerHeight: number | undefined;
  const owner = cells[0];
  if (nextPageBreak && declaredCellBreak && !tableProperties.hasCaption && dims.rowCount === 1 && dims.colCount === 1 &&
      cells.length === 1 && owner.rowSpan === 1 && owner.colSpan === 1 && !owner.isHeader && owner.verticalAlign === 0 &&
      !owner.textDirection && !owner.diagonalSlash && !owner.diagonalBackSlash &&
      tableProperties.vertRelTo === 'Para' && tableProperties.vertAlign === 'Top' && !tableProperties.vertOffset &&
      tableProperties.horzRelTo === 'Column' && tableProperties.textWrap === 'TopAndBottom' &&
      pageBoxes.length > 1 && rows[0] <= 2000 && /<table\b[^>]*>[\s\S]*<table\b/i.test(html)) {
    const fragments = [...pageBoxes].sort((a, b) => a.pageIndex - b.pageIndex);
    const first = fragments[0], last = fragments.at(-1)!;
    const tolerance = sourceCellBoundsTolerance;
    if (pitches.get(`page:${section}:${para + 1}`) === last.pageIndex + 1 &&
        fragments.every((box, index) => {
          const body = sourceBody(box.pageIndex);
          return box.cellIdx === owner.index && box.pageIndex === first.pageIndex + index &&
            box.h > 0 && Math.abs(box.w - first.w) <= tolerance && body.columns === 1 &&
            box.y >= body.top - tolerance && box.y + box.h <= body.bottom + tolerance;
        }) && last.y + last.h < sourceBody(last.pageIndex).bottom - tolerance) {
      const used = Math.round(fragments.reduce((sum, box) => sum + box.h, 0) * 100) / 100;
      if (used > 0 && rows[0] - used > tolerance) observedContainerHeight = used;
    }
  }
  const inlineStrokeHeight = dims.rowCount === 1 && dims.colCount === 1 && cells.length === 1 &&
    owner.rowSpan === 1 && owner.colSpan === 1 && !owner.isHeader && !owner.textDirection &&
    !owner.diagonalSlash && !owner.diagonalBackSlash && owner.borderTop?.type === 1 && owner.borderBottom?.type === 1 &&
    rows[0] > 0 && rows[0] <= 2000 && /<table\b[^>]*>[\s\S]*<table\b/i.test(html)
    ? sourceInlineStrokeHeight(doc, section, para, control) : undefined;
  let centerBorders: boolean | 'inline' | 'continuation' = sourceCentersTableBorders(doc, section, para, control, tableProperties, cells, pageBoxes,
    page => sourceBody(page).columns === 1, splitCellLines);
  const intactContinuationGrid = !centerBorders && sourceCentersMultilineRowBorders(doc, section, para, control,
    tableProperties, cells, pageBoxes, page => sourceBody(page).columns === 1, splitCellLines);
  if (intactContinuationGrid) centerBorders = 'continuation';
  const firstTableFragment = declaredCellBreak && tableProperties.repeatHeader && !tableProperties.hasCaption &&
    tableProperties.horzRelTo === 'Column' && tableProperties.vertRelTo === 'Para' &&
    tableProperties.vertAlign === 'Top' && tableProperties.textWrap === 'TopAndBottom' &&
    dims.colCount >= 2 && dims.colCount <= 6
    ? sourceTableContinuation(doc, section, para, control, dims.rowCount, cells, pageBoxes) : undefined;
  const firstTableOverhang = !firstTableFragment && declaredCellBreak && tableProperties.repeatHeader &&
    tableProperties.restrictInPage && !tableProperties.allowOverlap && !tableProperties.hasCaption &&
    tableProperties.textWrap === 'TopAndBottom' && tableProperties.vertRelTo === 'Para' &&
    tableProperties.vertAlign === 'Top' && tableProperties.horzRelTo === 'Column' && tableProperties.horzAlign === 'Left' &&
    dims.colCount >= 2 && dims.colCount <= 6
    ? sourceFirstTableOverhang(doc, section, para, control, dims.rowCount, cells, pageBoxes) : undefined;
  let repeatTop: number | undefined;
  if (declaredCellBreak && tableProperties.repeatHeader && !tableProperties.hasCaption &&
    tableProperties.textWrap === 'TopAndBottom' && tableProperties.vertRelTo === 'Para' &&
    tableProperties.vertAlign === 'Top' && Number(tableProperties.vertOffset) < 0) {
    const firstCell = cells.find(cell => cell.row === 0 && cell.col === 0 && cell.isHeader);
    const headerBoxes = pageBoxes.filter(box => box.cellIdx === firstCell?.index);
    const pages = new Set(pageBoxes.map(box => box.pageIndex));
    if (pages.size > 1 && headerBoxes.length === pages.size && new Set(headerBoxes.map(box => box.pageIndex)).size === pages.size) {
      const sourcePages = new Map(headerBoxes.map(box => [box.pageIndex, sourceBody(box.pageIndex)]));
      if ([...sourcePages.values()].every(page => page.columns === 1)) {
        // Repeat the outer clearance if it is present above every native header.
        const outerTop = hwpTableTopGap(px(Number(tableProperties.outerTop)));
        if (outerTop && headerBoxes.every(box => Math.abs(box.y - sourcePages.get(box.pageIndex)!.top - outerTop) <= .1))
          repeatTop = outerTop;
        // A shorter first cell can pin the shared edge in the grid solver.
        // Keep the larger authored minimum where every ordinary cell occupies
        // that same intact native row within the body. Split and merged rows
        // keep the grid calculation.
        for (const [row, own] of cellsByRow) {
          if (own.length < 2 || mergesByRow.has(row) || own.some(cell => !(cell.height > 0) || cell.isHeader)) continue;
          const height = Math.max(...own.map(cell => px(cell.height)));
          if (height <= rows[row] + .1) continue;
          const boxes = boxesByRow.get(row);
          if (!boxes || boxes.length !== own.length) continue;
          const first = boxes[0], paper = sourcePages.get(first.pageIndex);
          if (!paper || !(first.y >= paper.top - .1 && first.y + first.h <= paper.bottom + .1)) continue;
          const byCell = new Map(boxes.map(box => [box.cellIdx, box]));
          if (own.every(cell => {
            const box = byCell.get(cell.index);
            return box && box.pageIndex === first.pageIndex &&
              Math.abs(box.y - first.y) <= .1 && Math.abs(box.h - height) <= .1;
          })) rows[row] = height;
        }
      }
    }
  }
  // Remember an intact centered row placed immediately after a native page
  // break inside this table. CELL still permits splitting after an edit.
  const intactCenteredRows = new Set<number>();
  const headerBoxes = boxesByRow.get(0);
  if (declaredCellBreak && tableProperties.repeatHeader && !tableProperties.hasCaption &&
    cellsByRow.get(0)?.some(cell => cell.isHeader) && headerBoxes) {
    for (const [row, own] of cellsByRow) {
      if (row < 2 || mergesByRow.has(row) || own.length < 2 || own.some(cell =>
        cell.isHeader || cell.colSpan !== 1 || cell.verticalAlign !== 1 || cell.textDirection ||
        cell.diagonalSlash || cell.diagonalBackSlash)) continue;
      const boxes = boxesByRow.get(row), preceding = boxesByRow.get(row - 1);
      if (boxes?.length !== own.length || !preceding?.length) continue;
      const first = boxes[0], page = sourceBody(first.pageIndex);
      const previousPage = Math.max(...preceding.map(box => box.pageIndex));
      if (first.pageIndex !== previousPage + 1 || !preceding.some(box => box.pageIndex < previousPage) ||
        page.columns !== 1 || first.y + first.h > page.bottom + .1 ||
        Math.abs(first.h - Math.max(...own.map(cell => px(cell.height)))) > .1) continue;
      const header = headerBoxes.find(box => box.pageIndex === first.pageIndex);
      if (!header || header.y < page.top - .1 || Math.abs(first.y - header.y - header.h) > .15) continue;
      const ids = new Set(boxes.map(box => box.cellIdx));
      if (own.every(cell => ids.has(cell.index)) && boxes.every(box =>
        box.pageIndex === first.pageIndex && Math.abs(box.y - first.y) <= .1 && Math.abs(box.h - first.h) <= .1))
        intactCenteredRows.add(row);
    }
  }
  const intactTopRows = new Map<number, HwpIntactRow>();
  const measuredRepeat = !firstTableFragment && declaredCellBreak && tableProperties.repeatHeader &&
    !tableProperties.hasCaption && tableProperties.textWrap === 'TopAndBottom' &&
    tableProperties.horzRelTo === 'Column' && tableProperties.vertRelTo === 'Para' && tableProperties.vertAlign === 'Top' &&
    dims.colCount >= 2 && dims.colCount <= 6 && cellsByRow.get(0)?.length === dims.colCount &&
    cellsByRow.get(0)?.every(cell => cell.isHeader && cell.rowSpan === 1 && cell.colSpan === 1);
  if ((firstTableFragment || measuredRepeat) && tableProperties.restrictInPage &&
      !tableProperties.allowOverlap && tableProperties.cellSpacing === 0) {
    const firstPage = pageBoxes[0]?.pageIndex;
    const leading = new Map<number, number>();
    let headerRows = 0;
    if (measuredRepeat) {
      // A repeated header can include a second, deliberately empty row.
      for (; headerRows < dims.rowCount; headerRows++) {
        const cells = cellsByRow.get(headerRows) ?? [];
        if (cells.length !== dims.colCount ||
            !cells.every(cell => cell.isHeader && cell.rowSpan === 1 && cell.colSpan === 1)) break;
      }
      for (const box of pageBoxes) if (box.row >= headerRows)
        leading.set(box.pageIndex, Math.min(box.y, leading.get(box.pageIndex) ?? Infinity));
    }
    const repeatedBoxes = measuredRepeat ? pageBoxes.filter(box => box.row < headerRows) : [];

    for (const [row, own] of cellsByRow) {
      const boxes = boxesByRow.get(row);
      if (row <= (firstTableFragment?.row ?? 2) || own.length !== dims.colCount || mergesByRow.has(row) ||
          !boxes?.length || boxes[0].pageIndex === firstPage) continue;
      const first = boxes[0], page = sourceBody(first.pageIndex);
      if (measuredRepeat) {
        const previousPage = first.pageIndex - 1;
        const prior = boxesByRow.get(row - 1), earlier = boxesByRow.get(row - 2);
        if (first.y !== leading.get(first.pageIndex) || !prior || prior.length !== dims.colCount ||
            mergesByRow.has(row - 1) || !earlier?.some(box => box.pageIndex < previousPage) ||
            !earlier.some(box => box.pageIndex === previousPage) ||
            prior.some(box => box.pageIndex !== previousPage || Math.abs(box.y - prior[0].y) > .1 ||
              Math.abs(box.h - prior[0].h) > .1)) continue;
        const preceding = sourceBody(previousPage), precedingStart = leading.get(previousPage);
        if (preceding.columns !== 1 || precedingStart === undefined ||
            Math.abs(first.y - page.top - (precedingStart - preceding.top)) > sourceCellBoundsTolerance) continue;

        // The prior page carried the tail of another row and stopped after
        // an intact row. The same repeated header immediately leads the
        // candidate, and the earlier sheet cannot house its first ink line.
        const headers = repeatedBoxes.filter(box => box.pageIndex === first.pageIndex);
        const oldHeaders = repeatedBoxes.filter(box => box.pageIndex === previousPage);
        const headerCells = headerRows * dims.colCount;
        if (headers.length !== headerCells || oldHeaders.length !== headerCells ||
            !headers.every(box => oldHeaders.some(old => old.cellIdx === box.cellIdx &&
              Math.abs(box.y - page.top - (old.y - preceding.top)) <= sourceCellBoundsTolerance &&
              Math.abs(box.h - old.h) <= .1 && Math.abs(box.w - old.w) <= .1 && box.y + box.h <= first.y + .1))) continue;
        let firstInkEnd = Infinity;
        for (const cell of own) {
          const count = doc.getCellParagraphCount(section, para, control, cell.index);
          for (let p = 0; p < Math.min(16, count); p++) {
            const bands = splitCellLines.get(`${section}:${para}:${control}:${cell.index}:${p}`);
            for (const band of bands ?? []) if (band.page === first.pageIndex)
              firstInkEnd = Math.min(firstInkEnd, band.y + band.h - first.y);
          }
        }
        const available = preceding.bottom - prior[0].y - prior[0].h;
        if (!Number.isFinite(firstInkEnd) || available < -sourceCellBoundsTolerance ||
            available >= firstInkEnd - sourceCellBoundsTolerance) continue;
      }
      const proof = sourceIntactTopRow(doc, section, para, control, own, boxes, rows[row], splitCellLines, page);
      if (proof) intactTopRows.set(row, proof);
    }
  }
  const compactFaceCells = new Map<number, { box: typeof pageBoxes[number]; position: CompactCellFacePosition }>();
  if (!tableProperties.treatAsChar && tableProperties.pageBreak === 1 && tableProperties.repeatHeader &&
      tableProperties.restrictInPage && !tableProperties.allowOverlap && !tableProperties.hasCaption &&
      !tableProperties.cellSpacing && tableProperties.textWrap === 'TopAndBottom' &&
      tableProperties.vertRelTo === 'Para' && tableProperties.vertAlign === 'Top' && !tableProperties.vertOffset &&
      dims.colCount >= 2 && dims.colCount <= 12) {
    const cellsById = new Map(cells.map(cell => [cell.index, cell]));
    const pages = new Map<number, typeof pageBoxes>();
    for (const box of pageBoxes) {
      const list = pages.get(box.pageIndex) ?? []; list.push(box); pages.set(box.pageIndex, list);
    }
    const firstPage = Math.min(...pages.keys());
    for (const [page, boxes] of pages) {
      if (page === firstPage || !boxes.some(box => box.row === 0 && cellsById.get(box.cellIdx)?.isHeader)) continue;
      const body = sourceBody(page);
      const own = boxes.filter(box => box.row > 0 && !cellsById.get(box.cellIdx)?.isHeader);
      if (!own.length || body.columns !== 1) continue;
      const header = boxes.filter(box => box.row === 0);
      if (page === firstPage + 1 && pages.size >= 3 && header.length === dims.colCount &&
          Math.abs(header[0].h - Math.max(...header.map(box => cellsById.get(box.cellIdx)!.height / 75))) <= sourceCellBoundsTolerance &&
          header.every(box => {
            const cell = cellsById.get(box.cellIdx)!;
            return cell.isHeader && cell.rowSpan === 1 && cell.colSpan === 1 &&
              cellPages.get(box.cellIdx)?.size === pages.size &&
              Math.abs(box.y - body.top) <= sourceCellBoundsTolerance &&
              Math.abs(box.h - header[0].h) <= sourceCellBoundsTolerance &&
              pageBoxes.filter(part => part.cellIdx === box.cellIdx).every(part =>
                Math.abs(part.w - box.w) <= sourceCellBoundsTolerance && Math.abs(part.h - box.h) <= sourceCellBoundsTolerance);
          })) {
        for (const box of header) {
          const cell = cellsById.get(box.cellIdx)!;
          if (cell.verticalAlign === 1 && !cell.textDirection && !cell.diagonalSlash && !cell.diagonalBackSlash)
            compactFaceCells.set(box.cellIdx, { box, position: 'header' });
        }
      }
      const ids = own.map(box => box.row);
      const firstRow = Math.min(...ids), lastRow = Math.max(...ids);
      for (const row of new Set(ids)) {
        const onPage = own.filter(box => box.row === row), first = onPage[0];
        if (onPage.length !== dims.colCount || mergesByRow.has(row) || first.y < body.top - sourceCellBoundsTolerance ||
            first.y + first.h > body.bottom + sourceCellBoundsTolerance || onPage.some(box => {
              const cell = cellsById.get(box.cellIdx);
              return !cell || cell.rowSpan !== 1 || cell.colSpan !== 1 || cellPages.get(box.cellIdx)?.size !== 1 ||
                Math.abs(box.y - first.y) > sourceCellBoundsTolerance || Math.abs(box.h - first.h) > sourceCellBoundsTolerance ||
                Math.abs(cell.height / 75 - box.h) > sourceCellBoundsTolerance;
            })) continue;
        for (const box of onPage) {
          const cell = cellsById.get(box.cellIdx)!;
          if (cell.verticalAlign === 1 && !cell.textDirection && !cell.diagonalSlash && !cell.diagonalBackSlash)
            compactFaceCells.set(box.cellIdx, { box, position: row === firstRow || row === lastRow ? 'edge' : 'interior' });
        }
      }
    }
  }
  const beforeByCell = new Map<number, number[]>();
  const flowOwnerCells = new Set<number>();
  if (declaredCellBreak && !tableProperties.hasCaption) for (const [row, own] of cellsByRow) {
    if (own.length < 2 || mergesByRow.has(row) || own.some(cell => !(cell.height > 0) || cell.colSpan !== 1 || cell.isHeader)) continue;
    const boxes = boxesByRow.get(row);
    if (!boxes) continue;
    if (own.length === 2 && boxes.length > 2 && own.every(cell =>
      cell.verticalAlign === 0 && !cell.textDirection && !cell.diagonalSlash && !cell.diagonalBackSlash)) {
      const tolerance = sourceCellBoundsTolerance;
      const firstCell = boxes.filter(box => box.cellIdx === own[0].index).sort((a, b) => a.pageIndex - b.pageIndex);
      const otherCell = boxes.filter(box => box.cellIdx === own[1].index).sort((a, b) => a.pageIndex - b.pageIndex);
      const usedHeight = firstCell.reduce((sum, box) => sum + box.h, 0);
      const storedHeight = Math.max(...own.map(cell => px(cell.height)));
      // Some cells store continuous line positions across two source sheets.
      // Confirm both repeated headers and each painted line before using the
      // original spacing when the two native fragments are below the minimum.
      const repeatedFragments = measuredRepeat && firstCell.length <= 3 &&
        firstCell.every(fragment => headerBoxes?.filter(box => box.pageIndex === fragment.pageIndex).length === dims.colCount);
      const repeatedPair = repeatedFragments && firstCell.length === 2;
      const measuredRepeatedFragments = repeatedFragments && usedHeight < storedHeight - tolerance;
      const measuredTwoFragments = repeatedPair && usedHeight < storedHeight - tolerance;
      const measuredStablePair = repeatedPair && Math.abs(usedHeight - storedHeight) <= tolerance;
      const alignedFragments = firstCell.length >= 2 && otherCell.length === firstCell.length &&
        firstCell.every((fragment, index) => {
          const other = otherCell[index], page = sourceBody(fragment.pageIndex);
          return other.pageIndex === fragment.pageIndex && fragment.h > 0 &&
            (!index || fragment.pageIndex === firstCell[index - 1].pageIndex + 1) &&
            Math.abs(other.y - fragment.y) <= tolerance && Math.abs(other.h - fragment.h) <= tolerance &&
            page.columns === 1 && fragment.y >= page.top - tolerance && fragment.y + fragment.h <= page.bottom + tolerance;
        });
      // This table already has a first-fragment plan. Later whole-paragraph
      // boundaries still need their own repeated headers and painted evidence.
      if (alignedFragments && firstTableFragment && row > firstTableFragment.row && firstCell.length <= 3 &&
          usedHeight < storedHeight - tolerance && own.length === dims.colCount && firstCell.every(fragment => {
            const headers = headerBoxes?.filter(box => box.pageIndex === fragment.pageIndex) ?? [];
            const top = sourceBody(fragment.pageIndex).top + firstTableFragment.tableTop;
            return headers.length === dims.colCount && headers.every(header =>
              header.h > 0 && Math.abs(header.y - top) <= tolerance && header.y + header.h <= fragment.y);
          })) {
        const profiles = own.map((cell, index) => sourcePaintedCellBefore(doc, [section, para, control, cell.index],
          splitCellLines, pitches, index ? otherCell : firstCell, px(cell.paddingTop), px(cell.paddingBottom)));
        const gap = profiles[0]?.before[0];
        if (gap && profiles.every(profile => profile?.before.every(value => value === gap))) {
          const filled = firstCell.map(fragment => own.map((cell, index) => {
            let lastInk = -Infinity;
            for (const [paragraph] of profiles[index]!.before.entries()) {
              const lines = splitCellLines.get(`${section}:${para}:${control}:${cell.index}:${paragraph}`) ?? [];
              for (const line of lines) if (line.page === fragment.pageIndex) lastInk = Math.max(lastInk, line.y + line.h);
            }
            return Math.abs(fragment.y + fragment.h - px(cell.paddingBottom) - lastInk) <= tolerance;
          }));
          if (sourceWholeParagraphPeerRow(profiles, filled)) {
            own.forEach((cell, index) => beforeByCell.set(cell.index, profiles[index]!.before));
            continue;
          }
        }
      }
      if (alignedFragments && (usedHeight > storedHeight + tolerance || measuredRepeatedFragments || measuredStablePair)) {
        if (measuredStablePair) {
          // The source row can keep all of its authored height even when its
          // longer story starts a new paragraph at the repeated header.
          const measured = own.map((cell, index) => sourcePaintedCellBefore(doc, [section, para, control, cell.index],
            splitCellLines, pitches, index ? otherCell : firstCell, px(cell.paddingTop), px(cell.paddingBottom)));
          const [longer, shorter] = [...measured].sort((a, b) => (b?.before.length ?? 0) - (a?.before.length ?? 0));
          if (longer?.fillsRow && shorter?.before.length === 1 && !shorter.cuts.length &&
              !longer.terminalBlank && !shorter.terminalBlank && longer.cuts.length === 1 &&
              !longer.cuts[0].continuedLine && longer.cuts[0].paragraph >= 2 &&
              longer.before.length - longer.cuts[0].paragraph >= 2 &&
              longer.before.every(gap => gap === shorter.before[0]))
            own.forEach((cell, index) => beforeByCell.set(cell.index, measured[index]!.before));
          continue;
        }
        const margins = own.map((cell, index) => measuredRepeatedFragments && !measuredTwoFragments ? undefined
          : sourceSplitCellBefore(doc, [section, para, control, cell.index],
            splitCellLines, index ? otherCell : firstCell, px(cell.paddingTop), px(cell.paddingBottom), measuredTwoFragments));
        const [first, tail] = firstCell;
        const observed = (cell: number, index: number) => splitCellLines.get(
          `${section}:${para}:${control}:${cell}:${index}`,
        ) ?? [];
        if (measuredTwoFragments) {
          // Here the actual line starts, rather than cached vertical offsets,
          // prove missing paragraph gaps. Keep the ordinary split policy.
          margins.forEach((profile, index) => {
            if (!profile?.fillsFragments) return;
            const cell = own[index].index, peer = own[1 - index].index;
            const main = profile.before.map((_, p) => observed(cell, p));
            if (!main.some(lines => lines[0]?.page === first.pageIndex && lines.at(-1)?.page === tail.pageIndex) ||
                main.filter(lines => lines[0]?.page === tail.pageIndex).length < 2) return;
            const count = doc.getCellParagraphCount(section, para, control, peer);
            const peerLines = Array.from({ length: count }, (_, p) => observed(peer, p)).flat();
            if (count < 2 || !peerLines.length || peerLines.some(line => line.page !== first.pageIndex)) return;
            const path = JSON.stringify([{ controlIndex: control, cellIndex: peer, cellParaIndex: count - 1 }]);
            const blank = JSON.parse(doc.getParagraphLineMetrics(section, para, path));
            if (!blank.stored || blank.controlCount || blank.lines?.length !== 1 ||
                doc.getTextInCellByPath(section, para, path, 0, 0xffffffff) !== '') return;
            beforeByCell.set(cell, profile.before);
          });
        } else if (margins.some(profile => profile?.fillsFragments)) {
          margins.forEach((profile, index) => {
            if (profile) beforeByCell.set(own[index].index, profile.before);
            if (profile?.fillsFragments) flowOwnerCells.add(own[index].index);
          });
        } else if (firstCell.length === 2) {
          // A final paragraph can be the only story crossing the boundary.
          // Confirm the peer is intact and the measured final lines fill both
          // sheets before giving back the gaps. The normal splitter stays in use.
          if (repeatedPair && own.some((cell, index) => {
            const peerGaps = margins[1 - index]?.before, peer = own[1 - index].index;
            if (!peerGaps || peerGaps.length < 2 || peerGaps.some((_, p) => {
              const lines = observed(peer, p);
              return !lines.length || lines.some(line => line.page !== first.pageIndex);
            })) return false;
            const path: [number, number, number, number] = [section, para, control, cell.index];
            const profile = sourceSplitCellBefore(doc, path, splitCellLines, index ? otherCell : firstCell,
              px(cell.paddingTop), px(cell.paddingBottom), true);
            if (!profile?.fillsFragments || profile.before.length < 3 ||
                ![...profile.before, ...peerGaps].every(gap => gap === profile.before[0])) return false;
            const last = profile.before.length - 1, lines = observed(cell.index, last);
            if (profile.before.slice(0, last).some((_, p) =>
                observed(cell.index, p).some(line => line.page !== first.pageIndex)) ||
                lines.filter(line => line.page === first.pageIndex).length < 2 ||
                lines.filter(line => line.page === tail.pageIndex).length < 2) return false;
            beforeByCell.set(cell.index, profile.before);
            return true;
          })) continue;
          // One cell may establish the first fragment and the other may end
          // with an intact inline table that occupies the continuation.
          const fragments = [firstCell, otherCell];
          own.forEach((cell, index) => {
            const peer = own[1 - index], [first, tail] = fragments[index];
            const top = px(cell.paddingTop), bottom = px(cell.paddingBottom);
            const [peerFirst] = fragments[1 - index];
            const leading = sourceIntactCellBefore(doc, [section, para, control, peer.index], pitches,
              peerFirst.pageIndex, peerFirst.h - px(peer.paddingTop + peer.paddingBottom));
            if (!leading) return;
            const before = sourceBeforeNestedCellTail(doc, [section, para, control, cell.index],
              pitches, first, tail, top, bottom);
            if (before) {
              beforeByCell.set(peer.index, leading);
              beforeByCell.set(cell.index, before);
            }
          });
        }
        if (repeatedFragments && own.length === dims.colCount && !own.some(cell => beforeByCell.has(cell.index))) {
          const measured = own.map((cell, index) => sourcePaintedCellBefore(doc, [section, para, control, cell.index],
            splitCellLines, pitches, index ? otherCell : firstCell, px(cell.paddingTop), px(cell.paddingBottom)));
          const gap = measured[0]?.before[0];
          if (gap && measured.every(profile => profile?.before.every(value => value === gap)) &&
              measured.some((profile, index) => {
                const peer = measured[1 - index];
                if (!profile?.fillsRow || !peer) return false;
                if (measuredRepeatedFragments && firstCell.length === 3)
                  return !profile.cuts[0].continuedLine && profile.cuts[1].continuedLine &&
                    peer.cuts.length === 1 && peer.cuts[0].continuedLine && peer.terminalBlank;
                if (firstCell.length !== 2 || !profile.cuts[0].continuedLine) return false;
                const continuationHasAnotherParagraph = profile.cuts[0].paragraph + 1 < profile.before.length;
                const grows = usedHeight > storedHeight + tolerance;
                // The shorter story either ends before the original break
                // or starts a fresh paragraph on the continuation.
                if (peer.cuts.length === 0)
                  return grows && peer.before.length === 1 && profile.before.length === 2 ||
                    measuredTwoFragments && peer.terminalBlank && continuationHasAnotherParagraph;
                if (!peer.cuts[0].continuedLine) return profile.cuts[0].paragraph === 0;
                // When both cells continue text, they begin the continuation
                // from different paragraph roles and each has further gaps
                // before a new paragraph on the second sheet.
                return grows && continuationHasAnotherParagraph && peer.cuts[0].paragraph + 1 < peer.before.length &&
                  profile.cuts[0].paragraph + peer.cuts[0].paragraph === 1;
              })) {
            own.forEach((cell, index) => beforeByCell.set(cell.index, measured[index]!.before));
          }
        }
        if (repeatedPair && usedHeight > storedHeight + tolerance && own.length === dims.colCount &&
            !own.some(cell => beforeByCell.has(cell.index))) {
          // Restore the intact story's own spacing when the other cell alone
          // fills the continuation with the final line of its last paragraph.
          const fragments = [firstCell, otherCell];
          own.some((cell, index) => {
            const continuing = sourcePaintedCellBefore(doc, [section, para, control, cell.index],
              splitCellLines, pitches, fragments[index], px(cell.paddingTop), px(cell.paddingBottom), true);
            const cut = continuing?.cuts[0];
            if (!continuing?.fillsRow || continuing.terminalBlank || continuing.cuts.length !== 1 ||
                !cut?.continuedLine || cut.paragraph < 2 || cut.paragraph !== continuing.before.length - 1) return false;
            const final = observed(cell.index, cut.paragraph);
            if (final.filter(line => line.page === tail.pageIndex).length !== 1 ||
                final.filter(line => line.page === first.pageIndex).length < 2) return false;
            const peer = own[1 - index];
            const intact = sourcePaintedCellBefore(doc, [section, para, control, peer.index],
              splitCellLines, pitches, fragments[1 - index], px(peer.paddingTop), px(peer.paddingBottom));
            const gap = continuing.before[0];
            if (!intact || intact.terminalBlank || intact.cuts.length || intact.before.length < 2 ||
                ![...continuing.before, ...intact.before].every(value => value === gap)) return false;
            beforeByCell.set(peer.index, intact.before);
            return true;
          });
        }
        if (repeatedPair && usedHeight > storedHeight + tolerance && own.length === dims.colCount &&
            !own.some(cell => beforeByCell.has(cell.index))) {
          const fragments = [firstCell, otherCell];
          own.some((peer, index) => {
            const short = sourcePaintedCellBefore(doc, [section, para, control, peer.index],
              splitCellLines, pitches, fragments[index], px(peer.paddingTop), px(peer.paddingBottom));
            if (short?.before.length !== 1 || short.cuts.length || short.terminalBlank) return false;
            const lines = observed(peer.index, 0), last = lines.at(-1);
            if (lines.length < 2 || !last || last.page !== first.pageIndex ||
                Math.abs(first.y + first.h - px(peer.paddingBottom) - last.y - last.h) > tolerance) return false;
            const main = own[1 - index];
            const continued = sourcePaintedCellBefore(doc, [section, para, control, main.index],
              splitCellLines, pitches, fragments[1 - index], px(main.paddingTop), px(main.paddingBottom), false, true);
            const cut = continued?.cuts[0];
            if (!continued?.fillsRow || continued.terminalBlank || continued.cuts.length !== 1 ||
                cut?.continuedLine !== false || cut.paragraph !== 1 || continued.before.length < 3 ||
                continued.before.some(gap => gap !== short.before[0])) return false;
            beforeByCell.set(peer.index, short.before);
            beforeByCell.set(main.index, continued.before);
            return true;
          });
        }
        if (measuredRepeatedFragments && firstCell.length === 3 && own.length === dims.colCount &&
            !own.some(cell => beforeByCell.has(cell.index))) {
          const fragments = [firstCell, otherCell];
          own.some((cell, index) => {
            const profile = sourcePaintedCellBefore(doc, [section, para, control, cell.index],
              splitCellLines, pitches, fragments[index], px(cell.paddingTop), px(cell.paddingBottom), true);
            if (!profile?.fillsRow || profile.terminalBlank || profile.cuts.length !== 2 ||
                profile.cuts.some(cut => cut.continuedLine) || profile.before.some(gap => gap !== profile.before[0])) return false;
            const [first, second] = profile.cuts.map(cut => cut.paragraph);
            if (Math.min(first, second - first, profile.before.length - second) < 2 ||
                !sourceInlinePeerInMiddleFragment(doc, [section, para, control, own[1 - index].index], fragments[1 - index][1])) return false;
            beforeByCell.set(cell.index, profile.before);
            return true;
          });
        }
      }
    }
    if (boxes.length !== own.length) continue;
    const first = boxes[0];
    // When the measured row grows beyond every stored cell minimum, its
    // height can independently confirm which text cells supplied that growth.
    if (!(first.h > Math.max(...own.map(cell => px(cell.height))) + .1)) continue;
    const paper = sourceBody(first.pageIndex);
    if (paper.columns !== 1 || !(first.y >= paper.top - .1 && first.y + first.h <= paper.bottom + .1)) continue;
    const byCell = new Map(boxes.map(box => [box.cellIdx, box]));
    if (!own.every(cell => {
      const box = byCell.get(cell.index);
      return box && box.pageIndex === first.pageIndex && Math.abs(box.y - first.y) <= .1 && Math.abs(box.h - first.h) <= .1;
    })) continue;
    for (const cell of own) {
      if (cell.textDirection || cell.diagonalSlash || cell.diagonalBackSlash) continue;
      const padding = (Number(cell.paddingTop) + Number(cell.paddingBottom)) / 75;
      const before = sourceIntactCellBefore(doc, [section, para, control, cell.index], pitches, first.pageIndex, first.h - padding);
      if (before) beforeByCell.set(cell.index, before);
    }
  }
  let alignedHeaderEnds: Map<number, number> | undefined;
  if (declaredCellBreak && tableProperties.repeatHeader && !tableProperties.hasCaption &&
      !tableProperties.allowOverlap && tableProperties.cellSpacing === 0 && tableProperties.textWrap === 'TopAndBottom') {
    for (const owner of cells) {
      const count = cellPages.get(owner.index)?.size ?? 0;
      if (owner.rowSpan < 3 || count < 3) continue;
      // This extended flow places the copied header at the body edge.
      // Tables that reserve an authored outside inset need their own plan.
      if (count > 3 && (tableProperties.outerTop !== 0 || tableProperties.outerBottom !== 0)) continue;
      const group = Array.from({ length: owner.rowSpan }, (_, i) => owner.row + i);
      if (group.some(row => mergesByRow.get(row)?.length !== 1)) continue;
      const neighbours = group.map(row => (cellsByRow.get(row) ?? []).filter(cell => cell !== owner));
      const boxes = group.flatMap(row => boxesByRow.get(row) ?? []);
      const cuts = sourceEmptyMergeCuts(doc, [section, para, control], owner, neighbours, boxes, sourceBody, splitCellLines);
      if (count > 3) {
        if (cuts.length < 2 || centerBorders && centerBorders !== 'continuation') continue;
        const headerEnds = alignedHeaderEnds ??= sourceAlignedHeaderEnds(cells, boxesByRow.get(0) ?? [], sourceBody);
        const pages = [...new Set(pageBoxes.map(box => box.pageIndex))];
        const first = Math.min(...pages);
        if (sourceBody(first).columns !== 1 || pages.some(page => page !== first && !headerEnds.has(page)) ||
            !centerBorders && !sourceCentersContinuedBorders(doc, section, para, control, cells, pageBoxes, splitCellLines)) continue;
        centerBorders = 'continuation';
      }
      for (const row of cuts)
        observedSplitRows.add(row);
    }
  }
  for (const c of cells) if ((!c.isHeader || !tableProperties.repeatHeader) && c.rowSpan === 1) {
    const pages = cellPages.get(c.index), merges = mergesByRow.get(c.row);
    if (declaredCellBreak || (pages?.size ?? 0) > 1) continuationRows.add(c.row);
    // The shared render split projects one cut into the continuous editable
    // grid. A spanning cell covering additional sheets needs a different
    // continuation plan, even if its neighbouring cell only crosses once.
    if (pages?.size === 2 && merges?.length &&
      pitches.has(`continues:${section}:${para}:${control}:${c.index}`) && merges.every(index => {
        const mergedPages = cellPages.get(index);
        return mergedPages?.size === 2 && [...pages].every(page => mergedPages.has(page));
      }))
      observedSplitRows.add(c.row);
    if (!pages || pages.size < 3 || merges?.length !== 1 || dims.colCount !== 2 || c.colSpan !== 1 || c.verticalAlign !== 0 ||
      !declaredCellBreak || tableProperties.repeatHeader || tableProperties.hasCaption ||
      tableProperties.textWrap !== 'TopAndBottom' || tableProperties.vertRelTo !== 'Para' || Number(tableProperties.vertOffset) !== 0) continue;
    const owner = cells.find(cell => cell.index === merges[0]);
    const ownerPages = cellPages.get(merges[0]);
    const sheetNumbers = [...pages].sort((a, b) => a - b);
    if (!owner || owner.rowSpan !== 2 || owner.row + 1 !== c.row || owner.colSpan !== 1 || owner.verticalAlign !== 0 ||
      ownerPages?.size !== pages.size || !sheetNumbers.every((page, index) => page === sheetNumbers[0] + index && ownerPages.has(page))) continue;
    const segments = pageBoxes.filter(box => box.cellIdx === c.index);
    const ownerSegments = pageBoxes.filter(box => box.cellIdx === owner.index);
    if (segments.length !== pages.size || ownerSegments.length !== pages.size || segments.some(box => {
      const body = sourceBody(box.pageIndex);
      if (body.columns !== 1 || !(box.y >= body.top - .1 && box.y + box.h <= body.bottom + .1)) return true;
      const ownerBox = ownerSegments.find(item => item.pageIndex === box.pageIndex);
      return !(box.h > 0) || !ownerBox || Math.abs(box.y + box.h - ownerBox.y - ownerBox.h) > .2;
    })) continue;
    const measured = Math.round(segments.reduce((sum, box) => sum + box.h, 0) * 100) / 100;
    if (!(measured > 0 && measured < rows[c.row] - .5 && rows[c.row] <= 2000)) continue;
    // A tall merged box can be empty on one of its pages. Require painted
    // text in the ordinary cell before using the sum as a render-only minimum.
    const painted = sheetNumbers.every(page => (JSON.parse(doc.getPageTextLayout(page)).runs ?? []).some((run: any) => {
      const step = run.cellPath?.[0];
      return run.secIdx === section && run.parentParaIdx === para && (run.cellPath?.length ?? 0) <= 1 &&
        (step?.controlIndex ?? run.controlIdx) === control && (step?.cellIndex ?? run.cellIdx) === c.index && run.text?.trim();
    }));
    if (painted) rowReserves.set(c.row, measured);
  }
  const continuedReserves = new Map<number, NonNullable<ReturnType<typeof sourceContinuedRowReserve>>>();
  if (centerBorders === 'continuation' && alignedHeaderEnds) {
    const headerEnds = alignedHeaderEnds;
    for (const row of observedSplitRows) {
      if (mergesByRow.get(row)?.length !== 1 || rowReserves.has(row)) continue;
      const reserve = sourceContinuedRowReserve(cellsByRow.get(row) ?? [], boxesByRow.get(row) ?? [], rows[row],
        sourceBody, page => headerEnds.get(page));
      if (reserve) continuedReserves.set(row, reserve);
    }
  }
  const centeredLabels = new Map<number, string>();
  const sourceScaledCells = new Map<number, 2 | 3>();
  if (declaredCellBreak && tableProperties.repeatHeader && !tableProperties.hasCaption &&
      tableProperties.cellSpacing === 0 && tableProperties.restrictInPage && !tableProperties.allowOverlap) {
    for (const [row, own] of cellsByRow) {
      const boxes = boxesByRow.get(row);
      const merged = mergesByRow.get(row);
      if (own.length < 2 || own.some(cell => cell.isHeader || cell.colSpan !== 1) ||
          !boxes || boxes.length <= own.length || merged && (merged.length !== 1 || !observedSplitRows.has(row))) continue;
      const firstPage = Math.min(...boxes.map(box => box.pageIndex));
      const head = boxes.filter(box => box.pageIndex === firstPage), body = sourceBody(firstPage);
      if (head.length !== own.length || body.columns !== 1 || head.some(box => !(box.h > 0) ||
          Math.abs(box.y - head[0].y) > .1 || Math.abs(box.h - head[0].h) > .1)) continue;
      if (!own.some(cell => pitches.has(`continues:${section}:${para}:${control}:${cell.index}`))) continue;
      if (merged) {
        // Match the same cut of the already witnessed vertical merge. Only
        // the ordinary cells that are fully present on both sheets qualify.
        const owner = cells.find(cell => cell.index === merged[0]);
        const ownerBoxes = pageBoxes.filter(box => box.cellIdx === owner?.index);
        const firstMerge = ownerBoxes.find(box => box.pageIndex === firstPage);
        const nextMerge = ownerBoxes.find(box => box.pageIndex === firstPage + 1);
        if (owner?.row === row && owner.col === 0 && owner.colSpan === 1 && own.length === dims.colCount &&
            own.length >= 4 && ownerBoxes.length === 4 && firstMerge && nextMerge &&
            ownerBoxes.every((box, index) => box.pageIndex === firstPage + index &&
              Math.abs(box.w - firstMerge.w) <= sourceCellBoundsTolerance && sourceBody(box.pageIndex).columns === 1)) {
          // A vertical merge begins in this row, and the other cells have
          // three complete fragment boxes inside that merge. If their short
          // labels were painted in the first box, retain that one line there.
          const ordinary = own.filter(cell => cell !== owner);
          const physical = ordinary.map(cell => boxes.filter(box => box.cellIdx === cell.index)
            .sort((a, b) => a.pageIndex - b.pageIndex));
          const close = (a: number, b: number) => Math.abs(a - b) <= sourceCellBoundsTolerance;
          if (physical.every(parts => parts.length === 3 && parts.every((box, index) => {
            const peer = physical[0][index], mergeBox = ownerBoxes[index];
            return box.pageIndex === firstPage + index && box.h > 0 && close(box.y, peer.y) &&
              close(box.h, peer.h) && close(box.w, parts[0].w) && close(box.y, mergeBox.y) &&
              box.y + box.h <= mergeBox.y + mergeBox.h + sourceCellBoundsTolerance;
          }))) {
            const labels = ordinary.flatMap((cell, index) => {
              if (cell.rowSpan !== 1 || cell.verticalAlign !== 1 || cell.textDirection ||
                  cell.diagonalSlash || cell.diagonalBackSlash ||
                  doc.getCellParagraphCount(section, para, control, cell.index) !== 1) return [];
              const proof = sourceCenteredCellLabels(doc, section, para, control, cell.index, physical[index][0], splitCellLines);
              return proof ? [[cell.index, proof + ' first'] as const] : [];
            });
            const continued = ordinary.filter(cell => !labels.some(([index]) => index === cell.index) &&
              pitches.has(`continues:${section}:${para}:${control}:${cell.index}`));
            if (labels.length >= 2 && continued.length) {
              for (const [cell, proof] of labels) centeredLabels.set(cell, proof);
              for (const cell of continued) sourceScaledCells.set(cell.index, 3);
            }
          }
          continue;
        }
        const tail = boxes.filter(box => box.pageIndex === firstPage + 1);
        if (!owner || owner.colSpan !== 1 || owner.row >= row || !firstMerge || !nextMerge || ownerBoxes.length !== 2 ||
            Math.abs(firstMerge.w - nextMerge.w) > .1 || tail.length !== own.length ||
            boxes.length !== head.length + tail.length || sourceBody(firstPage + 1).columns !== 1 ||
            head[0].y < firstMerge.y - .1 || head[0].y + head[0].h > firstMerge.y + firstMerge.h + .1 ||
            tail.some(box => {
              const start = head.find(first => first.cellIdx === box.cellIdx);
              return !start || Math.abs(start.w - box.w) > .1 || Math.abs(box.y - nextMerge.y) > .1 ||
                !(box.h > 0) || Math.abs(box.h - tail[0].h) > .1 || box.y + box.h > nextMerge.y + nextMerge.h + .1;
            })) continue;
      }
      for (const cell of own) {
        const first = head.find(box => box.cellIdx === cell.index);
        if (!first || !Number.isFinite(authoredCellWidths.get(cell.index)) || (cellPages.get(cell.index)?.size ?? 0) < 2 ||
            cell.verticalAlign !== 1 || cell.textDirection || cell.diagonalSlash || cell.diagonalBackSlash) continue;
        const proof = sourceCenteredCellLabels(doc, section, para, control, cell.index, first, splitCellLines, merged !== undefined);
        if (!proof) continue;
        const peer = own.length === 2 && merged && own.find(item => item.index !== cell.index);
        const headPeer = peer && head.find(box => box.cellIdx === peer.index);
        const tailPeer = peer && boxes.find(box => box.cellIdx === peer.index && box.pageIndex === firstPage + 1);
        const compact = headPeer && tailPeer && first.h < rows[row] - sourceCellBoundsTolerance &&
          sourceCompactContinuedCell(doc, section, para, control, headPeer, tailPeer);
        centeredLabels.set(cell.index, compact ? proof + ' first ' + compact : proof);
      }
    }
  }
  const style = (tag: string, css: string) => tag.includes('style="') ? tag.replace(/style="([^"]*)"/, (_, old) => `style="${old}${css}"`) : tag.replace(/>$/, ` style="${css}">`);
  if (centerBorders === 'continuation' && declaredCellBreak && tableProperties.repeatHeader &&
      !tableProperties.hasCaption && !tableProperties.cellSpacing && tableProperties.restrictInPage && !tableProperties.allowOverlap) {
    for (const [row, cellsInRow] of cellsByRow) {
      const boxes = boxesByRow.get(row), first = boxes?.[0];
      if (!first || cellsInRow.length < 3 || boxes.length !== cellsInRow.length || mergesByRow.has(row) ||
          cellsInRow.every(cell => cell.colSpan === 1) || cellsInRow.some(cell => cell.isHeader || cell.rowSpan !== 1 ||
            cell.textDirection || cell.diagonalSlash || cell.diagonalBackSlash) ||
          cellsInRow.reduce((sum, cell) => sum + cell.colSpan, 0) !== dims.colCount ||
          first.h <= Math.max(...cellsInRow.map(cell => cell.height / 75)) + sourceCellBoundsTolerance ||
          boxes.some(box => box.pageIndex !== first.pageIndex || Math.abs(box.y - first.y) > .1 ||
            Math.abs(box.h - first.h) > .1 || cellPages.get(box.cellIdx)?.size !== 1)) continue;
      const body = sourceBody(first.pageIndex);
      if (body.columns !== 1 || first.y < body.top - .16 || first.y + first.h > body.bottom + .16) continue;
      const rowWidth = boxes.reduce((sum, box) => sum + box.w, 0);
      for (const cell of cellsInRow) {
        const box = boxes.find(box => box.cellIdx === cell.index);
        if (!box || box.w <= rowWidth - box.w || cell.colSpan !== 1 || cell.verticalAlign !== 1 ||
            doc.getCellParagraphCount(section, para, control, cell.index) !== 3) continue;
        const hasThreeIntactParagraphs = [0, 1, 2].every(cellParaIndex => {
          const key = JSON.stringify([{ controlIndex: control, cellIndex: cell.index, cellParaIndex }]);
          const metric = JSON.parse(doc.getParagraphLineMetrics(section, para, key));
          const ink = splitCellLines.get(`${section}:${para}:${control}:${cell.index}:${cellParaIndex}`);
          return metric.stored && !metric.controlCount && metric.lines?.length === 2 && ink?.length === 2 &&
            ink.every(line => line.page === first.pageIndex && line.y >= box.y - .16 && line.y + line.h <= box.y + box.h + .16);
        });
        if (hasThreeIntactParagraphs) sourceScaledCells.set(cell.index, 2);
      }
    }
  }
  let depth = 0, cellIndex = 0, row = 0, cellParagraph = 0, cellStart = 0, header = false, paragraphStyle = '', paragraphProperties: Record<string, any> = {};
  const paintedPages = new Map<number, NativeCellInkRun[]>();
  const readCellPaint = (page: number) => {
    if (!paintedPages.has(page)) paintedPages.set(page, JSON.parse(doc.getPageTextLayout(page)).runs ?? []);
    return paintedPages.get(page)!;
  };
  const pairedWidthProofs = new Map<number, HwpSourceWrap[] | undefined>();
  const pairedWidths = (c: typeof cells[number], box: typeof pageBoxes[number]) => {
    if (pairedWidthProofs.has(c.index)) return pairedWidthProofs.get(c.index);
    pairedWidthProofs.set(c.index, undefined);
    const own = cellsByRow.get(c.row), boxes = boxesByRow.get(c.row);
    if (!own || !boxes || own.length !== boxes.length || c.colSpan !== 1 ||
        doc.getCellParagraphCount(section, para, control, c.index) !== 2 ||
        boxes.some(part => part.pageIndex !== box.pageIndex || Math.abs(part.y - box.y) > sourceCellBoundsTolerance ||
          Math.abs(part.h - box.h) > sourceCellBoundsTolerance)) return;
    const [first, second] = [0, 1].map(p => splitCellLines.get(`${section}:${para}:${control}:${c.index}:${p}`));
    if (!first?.length || !second?.length || Math.max(first.length, second.length) <= 2 ||
        first.at(-1)!.y >= second[0].y || [...first, ...second].some(line => line.page !== box.pageIndex)) return;
    const proofs = [0, 1].map(index => sourceHwpCompactCellFace(doc, { section, paragraph: para, control, cell: c.index, index },
      box, 2, JSON.parse(doc.getCellParaPropertiesAt(section, para, control, c.index, index)),
      { intactColumns: c.colSpan, paragraphs: 2 }, () => readCellPaint(box.pageIndex)));
    if (!proofs.every(proof => proof !== undefined)) return;
    pairedWidthProofs.set(c.index, proofs);
    return proofs;
  };
  const hasOnlyInlineTable = (content: string) => {
    const cell = cells[cellIndex - 1];
    if (!cell || cell.rowSpan !== 1 || !/^<table\b/.test(content)) return false;
    const [table, extra] = elementRanges(content, 'table');
    if (!table || extra || table.start !== 0 || table.end !== content.length) return false;
    try {
      if (doc.getCellParagraphCount(section, para, control, cell.index) !== 1) return false;
      const path = JSON.stringify([{ controlIndex: control, cellIndex: cell.index, cellParaIndex: 0 }]);
      if (doc.getTextInCellByPath(section, para, path, 0, 0xffffffff) !== '') return false;
      const metrics = JSON.parse(doc.getParagraphLineMetrics(section, para, path));
      return metrics.stored === true && metrics.controlCount === 1 && metrics.lines.length === 1 &&
        metrics.tables.length === 1 && metrics.tables[0].inline === true;
    } catch { return false; }
  };
  const runRepairs: string[] = [];
  const nativeGrid = tableProperties.treatAsChar && hwpNativeGrid(tableProperties.tableWidth, columns, cells);
  let result = html.replace(/<\/?(?:table|tr|td|p)\b[^>]*>/g, (tag, offset: number) => {
    if (/^<table\b/.test(tag)) {
      if (++depth !== 1) return tag;
      const firstPlacement = firstTableFragment ?? firstTableOverhang;
      if (firstPlacement)
        tag = tag.replace(/>$/, ` data-hwp-table-continuation="${escape(JSON.stringify(firstPlacement))}">`);
      if (nestedLayout) tag = tag.replace(/>$/, ` ${hwpNestedTableLayoutAttribute}="${escape(nestedLayout)}">`);
      if (repeatTop) tag = tag.replace(/>$/, ` data-hwp-table-top-gap="${repeatTop}">`);
      if (centerBorders) tag = tag.replace(/>$/, ` data-hwp-centered-borders="${centerBorders}">`);
      return style(tag, `width:${columns.reduce((sum, n) => sum + n, 0)}px;margin-bottom:${px(Math.max(0, Number(tableProperties.outerBottom) || 0))}px;`);
    }
    if (tag === '</table>') { depth--; return tag; }
    if (depth !== 1) return tag;
    if (/^<p\b/.test(tag)) {
      try {
        const sourceIndex = tag.match(/\bdata-hwp-cell-paragraph="(\d+)"/)?.[1];
        if (sourceIndex !== undefined) cellParagraph = Number(sourceIndex);
        const c = cells[cellIndex - 1], index = c.index;
        const properties = JSON.parse(doc.getCellParaPropertiesAt(section, para, control, index, cellParagraph));
        const char = JSON.parse(doc.getCellCharPropertiesAt(section, para, control, index, cellParagraph, 0));
        const key = `${section}:${para}:${control}:${index}:${cellParagraph}`;
        const pitch = pitches.get(key);
        if (properties.lineSpacingType === 'Percent' && Number(properties.lineSpacing) === 0 &&
          cellParagraph + 1 < doc.getCellParagraphCount(section, para, control, index)) {
          const next = cellParagraph + 1, nextKey = `${section}:${para}:${control}:${index}:${next}`;
          const firstY = pitches.get(`top:${key}`), nextY = pitches.get(`top:${nextKey}`), height = pitches.get(`height:${key}`);
          if (height && firstY !== undefined && nextY !== undefined && Math.abs(firstY - nextY) < .5 &&
            pitches.get(`page:${key}`) === pitches.get(`page:${nextKey}`) &&
            JSON.parse(doc.getLineInfoInCell(section, para, control, index, cellParagraph, 0)).lineCount === 1 &&
            JSON.parse(doc.getLineInfoInCell(section, para, control, index, next, 0)).lineCount === 1) {
            const following = fragment(doc.exportSelectionInCellHtml(section, para, control, index, next, 0, next, 0xffffffff));
            const reserve = physicalLineHeight(JSON.parse(doc.getCellParaPropertiesAt(section, para, control, index, next)),
              JSON.parse(doc.getCellCharPropertiesAt(section, para, control, index, next, 0)));
            if (!following.replace(/<[^>]*>/g, '').trim() && !/<(?:img|table|br)\b|<!--\s*rhwp:/i.test(following) && reserve >= height)
              tag = tag.replace(/>$/, ` data-hwp-zero-line="${height},${Math.round((reserve - height) * 100) / 100}">`);
          }
        }
        cellParagraph++;
        // A one-line source cell can be shorter than CSS's em * percentage.
        // Preserve its measured content pitch, without enforcing a fixed row.
        let cap: number | undefined;
        if (c.rowSpan === 1 && doc.getCellParagraphCount(section, para, control, index) === 1 &&
          JSON.parse(doc.getLineInfoInCell(section, para, control, index, 0, 0)).lineCount === 1)
          cap = px(c.height - c.paddingTop - c.paddingBottom);
        const text = html.slice(offset + tag.length, html.indexOf('</p>', offset));
        const final = cellParagraph === doc.getCellParagraphCount(section, para, control, index);
        if (tableProperties.treatAsChar && !tableProperties.hasCaption && dims.cellCount === 1 &&
            pageBoxes.length === 1 && sourceBody(pageBoxes[0].pageIndex).columns === 2 &&
            cellParagraph === 1 && final && !c.paddingTop && !c.paddingBottom &&
            !properties.spacingBefore && !properties.spacingAfter) {
          const images = text.match(/<img\b[^>]*>/g) ?? [];
          if (images.length === 1 && !/data-image-floating|<br\b|data-(?:note-reference|page-number|docx-tab)/i.test(text) &&
              !entities(text.replace(/<[^>]*>/g, '')).trim()) {
            const height = Number(images[0].match(/\bheight="([\d.]+)"/)?.[1]);
            if (height > 0 && Math.abs(pageBoxes[0].h - height) <= 1) {
              const path = JSON.stringify([{ controlIndex: control, cellIndex: index, cellParaIndex: 0 }]);
              const metrics = JSON.parse(doc.getParagraphLineMetrics(section, para, path));
              // The stored object line and physical cell both match the
              // picture. Its source height already includes the descent.
              if (metrics.stored && metrics.controlCount === 1 && metrics.lines?.length === 1 &&
                  Math.abs(metrics.lines[0].height - height) <= 1 && doc.getTextInCellByPath(section, para, path, 0, 0xffffffff) === '')
                tag = tag.replace(/>$/, ' data-word-text-alignment="top">');
            }
          }
        }
        const terminalBox = terminalBoxes?.get(index);
        const terminalInk = terminalBox && splitCellLines.get(key);
        const centeredTerminal = c.verticalAlign === 1 && !mergesByRow.has(c.row) && tableProperties.outerTop > 0 &&
          terminalBox && terminalInk?.length === 1 &&
          Math.abs(terminalInk[0].y + terminalInk[0].h / 2 - terminalBox.y - terminalBox.h / 2) <= sourceCellBoundsTolerance;
        const authoredTerminal = terminalBox && terminalInk && cellParagraph === 1 && final && !header &&
          c.rowSpan === 1 && (c.verticalAlign === 0 || centeredTerminal) && !c.textDirection && !c.diagonalSlash && !c.diagonalBackSlash &&
          !properties.spacingBefore && !properties.spacingAfter && !properties.indent && !properties.marginLeft &&
          !properties.marginRight && !/<(?:a|br|img|svg|table|sup|sub)\b|\bdata-(?:docx-tab|note-reference|document-field)/i.test(text)
          ? sourceCellTerminal(doc, section, para, control, index, c, terminalBox, terminalInk, !!centeredTerminal) : undefined;
        const cellLineOptions = { pitch, cap, authoredTerminal,
          before: beforeByCell.get(index)?.[cellParagraph - 1],
          inkHeight: pitches.get(`height:${key}`),
          final,
          lines: JSON.parse(doc.getLineInfoInCell(section, para, control, index, cellParagraph - 1, 0)).lineCount,
          simple: !/<(?:img|table)\b|<!--\s*rhwp:/i.test(text),
        };
        let paragraphCssValue = cellParagraphCss(properties, char, cellLineOptions);
        if (text.includes('\t') && !tag.includes('data-tab-stops')) tag = tag.replace(/>$/, `${tabs?.(index, cellParagraph - 1) ?? ''}>`);
        const tableSpaces = tag.includes('data-hwp-inline-table-spaces="true"');
        const mixed = !final && !tableSpaces && dims.cellCount === 1 && tableProperties.treatAsChar
          ? sourceMixedCellLeading(doc, section, para, [{ controlIndex: control, cellIndex: index, cellParaIndex: cellParagraph - 1 }],
            properties, pitch ?? physicalLineHeight(properties, char), text)
          : '';
        const wrap = new RegExp(` ${hwpSourceWrapAttribute}="([^"]*)"`).exec(tag);
        const reference = wrap ? parseHwpSourceWrap(entities(wrap[1])) : undefined;
        const outer = authoredCellWidths.get(index);
        if (wrap && reference && outer && (cellPages.get(index)?.size ?? 0) > 0) {
          const text = outer - px(c.paddingLeft + c.paddingRight) - properties.marginLeft - properties.marginRight + Math.min(0, properties.indent);
          if (text > 0 && text <= outer && outer <= 10000) {
            reference.cellWidth = { outer, text: Math.round(text * 1000) / 1000 };
            if (flowOwnerCells.has(index) && reference.ends.length > 1 && !reference.origin && !reference.nestedHanging)
              reference.splitCell = true;
            tag = tag.replace(wrap[0], ` ${hwpSourceWrapAttribute}="${escape(JSON.stringify(reference))}"`);
          }
        }
        const expanded = expandableCells.get(index);
        if (!wrap && expanded && pitch && !tableSpaces && !header && !properties.spacingBefore && !properties.spacingAfter &&
            !/<(?:a|br|img|svg|table|sup|sub)\b|\bdata-(?:docx-tab|note-reference|document-field)/i.test(text)) {
          const p = cellParagraph - 1;
          const path = JSON.stringify([{ controlIndex: control, cellIndex: index, cellParaIndex: p }]);
          const literal = doc.getTextInCellByPath(section, para, path, 0, 0xffffffff);
          const ink = splitCellLines.get(key);
          if (literal.length <= 2000 && ink && ink.length >= 2 &&
              ink.every(line => line.page === expanded.pageIndex && line.y >= expanded.y - sourceCellBoundsTolerance &&
                line.y + line.h <= expanded.y + expanded.h + sourceCellBoundsTolerance)) {
            const runs = sourceFontRuns(doc.getCharShapeRunsInCellByPath(section, para, path, 0, [...literal].length), 0, literal,
              at => JSON.parse(doc.getCellCharPropertiesAtByPath(section, para, path, at)));
            const readLines = () => doc.getParagraphLineMetrics(section, para, path);
            const source = preserveHwpSourceSpaces('<p>' + text + '</p>', literal, runs, hwpxFontSpace, readLines);
            const marked = source.match(new RegExp(` ${hwpSourceWrapAttribute}="([^"]*)"`));
            const proof = marked && parseHwpSourceWrap(entities(marked[1]));
            const width = expanded.w - px(c.paddingLeft + c.paddingRight) - properties.marginLeft -
              properties.marginRight + Math.min(0, properties.indent);
            if (proof && width > 0 && width <= expanded.w && expanded.w <= 10000 &&
                proof.ends.length === ink.length &&
                ink.slice(1).every((line, i) => line.charStart !== undefined &&
                  normalizeHwpWrapText(literal.slice(0, line.charStart)).length === proof.ends[i])) {
              const value = { ...proof, cellWidth: { outer: expanded.w, text: Math.round(width * 1000) / 1000 },
                expandedRowHeight: expanded.h };
              tag = tag.replace(/>$/, ` ${hwpSourceWrapAttribute}="${escape(JSON.stringify(value))}">`);
            }
          }
        }
        let compactFaceCell = compactFaceCells.get(index);
        if (!compactFaceCell && centerBorders === 'continuation' && declaredCellBreak && tableProperties.repeatHeader &&
            !header && c.colSpan === 1 && c.rowSpan === 1 && !c.textDirection && !c.diagonalSlash && !c.diagonalBackSlash &&
            properties.indent < 0 && splitCellLines.get(key)?.length) {
          const page = pitches.get(`page:${key}`);
          const box = pageBoxes.find(part => part.pageIndex === page && part.cellIdx === index);
          if (box && sourceBody(box.pageIndex).columns === 1)
            compactFaceCell = { box, position: 'continued' };
        }
        if (compactFaceCell && !tableSpaces &&
            (wrap ? compactFaceCell.position === 'continued' : !tag.includes(hwpSourceWrapAttribute)) &&
            !/<(?:a|img|svg|table|sup|sub)\b|\bdata-(?:docx-tab|note-reference|document-field)/i.test(text)) {
          const proof = sourceHwpCompactCellFace(doc, { section, paragraph: para, control, cell: index, index: cellParagraph - 1 },
            compactFaceCell.box, doc.getCellParagraphCount(section, para, control, index), properties, compactFaceCell.position, () => {
              const page = compactFaceCell.box.pageIndex;
              if (!paintedPages.has(page)) paintedPages.set(page, JSON.parse(doc.getPageTextLayout(page)).runs ?? []);
              return paintedPages.get(page)!;
            }, sourceScaledCells.get(index));
          if (proof && !wrap) tag = tag.replace(/>$/, ` ${hwpSourceWrapAttribute}="${escape(JSON.stringify(proof))}">`);
          else if (proof && wrap && reference) {
            // Keep a richer reference, including a split-cell policy added above.
            if (Object.keys(reference).every(key => ['fingerprint', 'ends', 'cellWidth'].includes(key)) &&
                reference.fingerprint === proof.fingerprint && reference.ends.length === proof.ends.length &&
                reference.ends.every((value, i) => value === proof.ends[i]))
              tag = tag.replace(new RegExp(` ${hwpSourceWrapAttribute}="[^"]*"`),
                ` ${hwpSourceWrapAttribute}="${escape(JSON.stringify(proof))}"`);
          }
        }
        const completeRow = completeExpandedCells.get(index);
        const hanging = -properties.indent;
        if (completeRow && !tableSpaces && !header && pitch && hanging > 0 && properties.alignment === 'justify' &&
            !properties.spacingBefore && !properties.spacingAfter && !properties.marginLeft && !properties.marginRight &&
            tag.includes('&quot;authoredCellFace&quot;') &&
            !/<(?:a|br|img|svg|table|sup|sub)\b|\bdata-(?:docx-tab|note-reference|document-field)/i.test(text)) {
          const path = JSON.stringify([{ controlIndex: control, cellIndex: index, cellParaIndex: cellParagraph - 1 }]);
          const ink = splitCellLines.get(key);
          const width = completeRow.w - px(c.paddingLeft + c.paddingRight);
          if (ink?.length === 2 && ink.every(line => line.page === completeRow.pageIndex &&
              line.y >= completeRow.y - sourceCellBoundsTolerance && line.y + line.h <= completeRow.y + completeRow.h + sourceCellBoundsTolerance) &&
              Math.abs(ink[1].y - ink[0].y - pitch) <= sourceCellBoundsTolerance && hanging < width / 4) {
            const literal = doc.getTextInCellByPath(section, para, path, 0, 0xffffffff);
            const metrics = doc.getParagraphLineMetrics(section, para, path);
            const source = preserveHwpColumnLineEnds('<p>' + text + '</p>', literal, () => metrics);
            const marked = source.match(new RegExp(` ${hwpSourceWrapAttribute}="([^"]*)"`));
            const proof = marked && parseHwpSourceWrap(entities(marked[1]));
            const runs = sourceFontRuns(doc.getCharShapeRunsInCellByPath(section, para, path, 0, [...literal].length), 0,
              literal, at => JSON.parse(doc.getCellCharPropertiesAtByPath(section, para, path, at)));
            const authoredScale = runs[0]?.widths?.[0];
            const previous = tag.match(new RegExp(` ${hwpSourceWrapAttribute}="([^"]*)"`));
            const prior = previous && parseHwpSourceWrap(entities(previous[1]));
            if (previous && prior?.authoredCellFace?.continued && proof?.columnWidth && proof.fingerprint === prior.fingerprint &&
                prior.ends.length === 2 && prior.ends.every((end, i) => end === proof.ends[i]) &&
                Math.abs(prior.authoredCellFace.rowHeight - completeRow.h) <= .1 &&
                Math.abs(prior.authoredCellFace.hanging! - hanging) <= .1 && ink[1].charStart !== undefined &&
                proof.ends.length === 2 && proof.ends[0] === normalizeHwpWrapText(literal.slice(0, ink[1].charStart)).length &&
                proof.ends[0] >= 20 && proof.ends[1] - proof.ends[0] <= 3 &&
                authoredScale && authoredScale >= 90 && authoredScale < 100 &&
                runs.every(run => run.widths?.length === 7 && run.widths.every(scale => scale === authoredScale)) &&
                Math.abs(width - proof.columnWidth) <= sourceCellBoundsTolerance) {
              const value = { ...prior,
                expandedRowHeight: completeRow.h, expandedRowLine: { hanging, pitch, scale: authoredScale } };
              const attribute = ` ${hwpSourceWrapAttribute}="${escape(JSON.stringify(value))}"`;
              tag = tag.replace(previous[0], attribute);
            }
          }
        }
        // The whole-row witness also verifies merged columns within the row.
        const intactWidthRow = intactContinuationGrid && !mergesByRow.has(c.row);
        if (centerBorders === 'continuation' && (declaredCellBreak || intactWidthRow) && tableProperties.repeatHeader &&
            !tableSpaces && !header && (c.colSpan === 1 || intactWidthRow) && c.rowSpan === 1 && c.verticalAlign === 1 &&
            !c.textDirection && !c.diagonalSlash && !c.diagonalBackSlash && properties.indent <= 0 &&
            text.includes(characterScaleProperty) &&
            !/<(?:a|img|svg|table|sup|sub)\b|\bdata-(?:docx-tab|note-reference|document-field)/i.test(text)) {
          const lines = splitCellLines.get(key), page = lines?.[0]?.page;
          const box = page === undefined || !lines?.every(line => line.page === page) ? undefined :
            pageBoxes.find(part => part.pageIndex === page && part.cellIdx === index);
          if (box && sourceBody(box.pageIndex).columns === 1) {
            const pair = intactWidthRow ? pairedWidths(c, box) : undefined;
            const proof = pair?.[cellParagraph - 1] ?? sourceHwpCompactCellFace(doc,
              { section, paragraph: para, control, cell: index, index: cellParagraph - 1 },
              box, doc.getCellParagraphCount(section,para,control,index), properties,
              intactWidthRow ? { intactColumns: c.colSpan } : 'scaledContinuation', () => readCellPaint(box.pageIndex));
            const marked = tag.match(new RegExp(` ${hwpSourceWrapAttribute}="([^"]*)"`));
            const prior = marked && parseHwpSourceWrap(entities(marked[1]));
            if (proof && (!marked || prior?.fingerprint === proof.fingerprint &&
                prior.ends.length === proof.ends.length && prior.ends.every((end,i) => end === proof.ends[i]))) {
              const value = { ...proof, ...prior, cellWidthFallback: proof.cellWidthFallback };
              const serialized = JSON.stringify(value);
              if (parseHwpSourceWrap(serialized)) {
                const attribute = ` ${hwpSourceWrapAttribute}="${escape(serialized)}"`;
                tag = marked ? tag.replace(marked[0],attribute) : tag.replace(/>$/,attribute + '>');
              }
            }
          }
        }
        if (!wrap && !header && intactColumnCell && intactColumnCell.cellIdx === index &&
            !c.textDirection && !c.diagonalSlash && !c.diagonalBackSlash && properties.alignment === 'justify' &&
            !/<(?:a|br|img|svg|table|sup|sub)\b|\bdata-(?:docx-tab|note-reference|document-field)/i.test(text)) {
          const path = [{ controlIndex: control, cellIndex: index, cellParaIndex: cellParagraph - 1 }];
          const visible = splitCellLines.get(key);
          const stories = doc.getCellParagraphCount(section, para, control, index);
          const singleStory = cellParagraph === 1 && final && stories === 1;
          let observedStories = singleStory ? 1 : 0;
          let tracked: ReturnType<typeof restoreUniformLatinMetrics>;
          const indentedLead = cellParagraph === 1 && stories === 2 && properties.indent > 0;
          if (visible && visible.length > 2 && (singleStory || stories >= 2 && stories <= 8 && !final) &&
              tableProperties.treatAsChar && (!properties.indent || indentedLead) && !properties.marginLeft && !properties.marginRight &&
              !properties.spacingBefore && !/<\/?(?!span\b)[a-z]/i.test(text)) {
            const address = JSON.stringify(path);
            const literal = doc.getTextInCellByPath(section, para, address, 0, 0xffffffff);
            if (literal.length <= 1024 && /^[ -~\u00a0\u2007\u2010-\u201f\u2026]+$/.test(literal)) {
              const runs = sourceFontRuns(doc.getCharShapeRunsInCellByPath(section, para, address, 0, literal.length),
                0, literal, at => JSON.parse(doc.getCellCharPropertiesAtByPath(section, para, address, at)), undefined, true);
              if (runs.length === 1) {
                tracked = restoreUniformLatinMetrics('<p>' + text + '</p>', literal, runs[0]);
                if (tracked) observedStories = stories;
              }
            }
          }
          const reference = sourceColumnCellWrap(doc, section, para, path, visible, intactColumnCell, observedStories);
          if (reference) {
            if (!indentedLead)
              tag = tag.replace(/>$/, ' ' + hwpSourceWrapAttribute + '="' + escape(JSON.stringify(reference)) + '">');
            if (tracked) {
              tag = tag.replace(/>$/, ' data-hwp-run-repair="' + (runRepairs.push(tracked.html.slice(3, -4)) - 1) + '">');
              // The complete source story verifies the physical advance
              // at the authored font. Future edits can still rescale it.
              if (singleStory && reference.columnCellPitch)
                paragraphCssValue = cellParagraphCss(properties, char, {
                  ...cellLineOptions, pitch: reference.columnCellPitch, linePolicyFont: tracked.size * 4 / 3,
                });
              // A verified gap in the same intact cell is the distance
              // beyond the last line's advance, before its short sibling.
              const last = visible?.at(-1), advance = reference.columnCellPitch;
              if (indentedLead && last && advance && properties.spacingAfter > 0) {
                const next = splitCellLines.get([section, para, control, index, 1].join(':'));
                const sibling = JSON.parse(doc.getCellParaPropertiesAt(section, para, control, index, 1));
                const authored = properties.spacingAfter / 2;
                const observed = next?.[0] && next[0].y - last.y - advance;
                if (next?.length === 1 && next[0].page === last.page && !sibling.spacingBefore &&
                    authored <= advance && observed !== undefined && Math.abs(observed - authored) <= sourceCellBoundsTolerance)
                  paragraphCssValue += 'margin-bottom:' + Math.round(authored * 15) / 15 + 'px;';
              }
            }
          }
        }
        if (!wrap && outer && !header && tableProperties.treatAsChar && tableProperties.pageBreak === 0 &&
          !tableProperties.hasCaption && c.rowSpan === 1 && c.colSpan === 1 && cellParagraph === 1 && final &&
          !c.textDirection && !c.diagonalSlash && !c.diagonalBackSlash && !properties.indent &&
          !properties.marginLeft && !properties.marginRight && cellPages.get(index)?.size === 1) {
          const ink = splitCellLines.get(key);
          const width = outer - px(c.paddingLeft + c.paddingRight);
          if (ink?.length === 1 && ink[0].shortLabelInk && width >= ink[0].shortLabelInk - .05 &&
              width - ink[0].shortLabelInk <= 1) {
            const address = JSON.stringify([{ controlIndex: control, cellIndex: index, cellParaIndex: 0 }]);
            const lines = JSON.parse(doc.getParagraphLineMetrics(section, para, address));
            const word = doc.getTextInCellByPath(section, para, address, 0, 0xffffffff);
            if (lines.stored && !lines.controlCount && lines.lines?.length === 1 && hwpShortCellLabel(word) &&
                Math.abs(lines.lines[0].width - width) <= .12) {
              const shape = { fingerprint: hwpWrapFingerprint(word), ends: [word.length],
                cellWidth: { outer, text: Math.round(width * 1000) / 1000 } };
              tag = tag.replace(/>$/, ` ${hwpSourceWrapAttribute}="${escape(JSON.stringify(shape))}">`);
            }
          }
        }
        if (!tag.includes(hwpSourceWrapAttribute) && header && c.isHeader && c.row === 0 && tableProperties.repeatHeader &&
            tableProperties.pageBreak === 1 && !tableProperties.treatAsChar && !tableProperties.allowOverlap &&
            tableProperties.textWrap === 'TopAndBottom' &&
            !tableProperties.cellSpacing && !tableProperties.hasCaption && c.verticalAlign === 1 &&
            c.rowSpan === 1 && c.colSpan === 1 && !c.textDirection && !c.diagonalSlash && !c.diagonalBackSlash &&
            cellParagraph === 1 && final && properties.alignment === 'center' &&
            !properties.indent && !properties.marginLeft && !properties.marginRight) {
          const boxes = pageBoxes.filter(box => box.cellIdx === index), first = boxes[0];
          if (boxes.length > 1 && boxes.every(box => Math.abs(box.w - first.w) <= sourceCellBoundsTolerance &&
              Math.abs(box.h - first.h) <= sourceCellBoundsTolerance)) {
            let proof: HwpSourceWrap | undefined;
            const sameHeader = boxes.every(box => {
              const body = sourceBody(box.pageIndex);
              if (body.columns !== 1 || box.y < body.top - .1 || box.y + box.h > body.bottom + .1) return false;
              if (!paintedPages.has(box.pageIndex)) paintedPages.set(box.pageIndex, JSON.parse(doc.getPageTextLayout(box.pageIndex)).runs ?? []);
              const line = sourceCellCaptionLine(doc,section,para,control,index,box,px(c.paddingLeft),px(c.paddingRight),
                paintedPages.get(box.pageIndex)!,true);
              if (!line || proof && proof.fingerprint !== line.fingerprint) return false;
              proof ??= line;
              return true;
            });
            if (sameHeader && proof) tag = tag.replace(/>$/, ` ${hwpSourceWrapAttribute}="${escape(JSON.stringify(proof))}">`);
          }
        }
        if (!tag.includes(hwpSourceWrapAttribute) && outer && !header && tableProperties.treatAsChar &&
            tableProperties.pageBreak === 0 && !tableProperties.hasCaption && !tableProperties.allowOverlap &&
            c.rowSpan === 1 && c.colSpan === 1 && !c.textDirection && !c.diagonalSlash && !c.diagonalBackSlash &&
            cellParagraph === 2 && final && !properties.indent && !properties.marginLeft && !properties.marginRight &&
            cellPages.get(index)?.size === 1) {
          try {
            const box = pageBoxes.find(box => box.cellIdx === index);
            const ink = splitCellLines.get(key);
            if (box && ink?.length === 1 && ink[0].page === box.pageIndex) {
              const page = box.pageIndex;
              if (!paintedPages.has(page)) paintedPages.set(page, JSON.parse(doc.getPageTextLayout(page)).runs ?? []);
              const proof = sourceCellCaptionLine(doc, section, para, control, index, box,
                px(c.paddingLeft), px(c.paddingRight), paintedPages.get(page)!);
              if (proof) tag = tag.replace(/>$/, ` ${hwpSourceWrapAttribute}="${escape(JSON.stringify(proof))}">`);
            }
          } catch { /* Invalid optional ink evidence keeps the normal cell layout. */ }
        }
        if (!tag.includes(hwpSourceWrapAttribute) && text.includes(characterScaleProperty) && outer && !header &&
          tableProperties.treatAsChar && !tableProperties.hasCaption && c.rowSpan === 1 && c.colSpan === 1 &&
          cellParagraph === 1 && final && !c.textDirection && !c.diagonalSlash && !c.diagonalBackSlash &&
          properties.indent <= 0 && !properties.marginLeft && !properties.marginRight && cellPages.get(index)?.size === 1) {
          const line = splitCellLines.get(key);
          if (line?.length === 1) {
            const page = line[0].page;
            if (!paintedPages.has(page)) paintedPages.set(page, JSON.parse(doc.getPageTextLayout(page)).runs ?? []);
            const reference = sourceScaledCellRun(doc, section, para, control, index, outer,
              outer - px(c.paddingLeft + c.paddingRight), line[0], paintedPages.get(page)!);
            if (reference) tag = tag.replace(/>$/, ` ${hwpSourceWrapAttribute}="${escape(JSON.stringify(reference))}">`);
          }
        }
        if (!tag.includes(hwpSourceWrapAttribute) && outer && !header && char.fontFamily === hwpScaledCellFace &&
            text.includes(characterScaleProperty) && dims.rowCount === 1 && dims.cellCount === dims.colCount &&
            tableProperties.treatAsChar && !tableProperties.hasCaption && !tableProperties.allowOverlap && !tableProperties.cellSpacing &&
            c.rowSpan === 1 && c.colSpan === 1 && cellParagraph === 1 && final && !c.textDirection &&
            !c.diagonalSlash && !c.diagonalBackSlash && !properties.indent &&
            !properties.marginLeft && !properties.marginRight && cellPages.get(index)?.size === 1) {
          try {
            const box = pageBoxes.find(box => box.cellIdx === index);
            if (box) {
              const page = box.pageIndex;
              if (!paintedPages.has(page)) paintedPages.set(page, JSON.parse(doc.getPageTextLayout(page)).runs ?? []);
              const proof = sourceScaledCellLine(doc, section, para, control, index, box,
                px(c.paddingLeft), px(c.paddingRight), px(c.height), paintedPages.get(page)!);
              if (proof) tag = tag.replace(/>$/, ` ${hwpSourceWrapAttribute}="${escape(JSON.stringify(proof))}">`);
            }
          } catch { /* Invalid optional ink evidence keeps the normal cell layout. */ }
        }
        return paragraphAutoSpace(cssStyle(tag, paragraphCssValue + (tableSpaces ? 'line-height:0.1pt;' : mixed)), properties);
      } catch { return tag; }
    }
    if (tag === '</td>') {
      // The inline table already occupies the sole native cell line. Its own
      // leading is on the table; DOCX supplies the required terminal carrier.
      // An actual empty source paragraph is handled by the paragraph path.
      const occupied = cellParagraph || hasOnlyInlineTable(html.slice(cellStart, offset).trim());
      return `${occupied ? '' : paragraphAutoSpace(`<p style="${paragraphStyle}">`, paragraphProperties) + '</p>'}</${header ? 'th' : 'td'}>`;
    }
    if (/^<tr\b/.test(tag)) {
      if (continuationRows.has(row)) tag = tag.replace(/>$/, ' data-row-split="allow">');
      if (observedSplitRows.has(row)) tag = tag.replace(/>$/, ' data-hwp-row-continued="true">');
      if (intactCenteredRows.has(row)) tag = tag.replace(/>$/, ' data-hwp-row-intact="true">');
      const intactTop = intactTopRows.get(row);
      if (intactTop) tag = tag.replace(/>$/, ` ${hwpIntactRowAttribute}="${escape(JSON.stringify(intactTop))}">`);
      const reserve = rowReserves.get(row);
      if (reserve !== undefined) tag = tag.replace(/>$/, ` data-hwp-row-reserve="1 ${rows[row]} ${reserve}">`);
      const continued = continuedReserves.get(row);
      if (continued) tag = tag.replace(/>$/, ` data-hwp-row-reserve="4 ${rows[row]} ${continued.observed} ${continued.headHeight}">`);
      if (row === 0 && observedContainerHeight !== undefined)
        tag = tag.replace(/>$/, ` data-hwp-row-reserve="2 ${rows[row]} ${observedContainerHeight}">`);
      else if (row === 0 && inlineStrokeHeight !== undefined)
        tag = tag.replace(/>$/, ` data-hwp-row-reserve="3 ${rows[row]} ${inlineStrokeHeight}">`);
      tag = tag.replace(/>$/, ' data-hwp-grid-row="true">');
      const height = automaticRows.has(row) ? 'auto' : `${rows[row] ?? 1}px`;
      row++;
      return style(tag, `height:${height};`);
    }
    if (!/^<td\b/.test(tag)) return tag;
    const c = cells[cellIndex++]; if (!c) return tag;
    cellParagraph = 0; cellStart = offset + tag.length; header = !!c.isHeader && !!tableProperties.repeatHeader;
    // CELL permits a continuing vertical merge after preceding edits move
    // a formerly single-page cell to a new boundary. Other policies need
    // observed source continuation before crossing that merged boundary.
    if (!header && c.rowSpan > 1 && (declaredCellBreak || (cellPages.get(c.index)?.size ?? 0) > 1))
      tag = tag.replace(/>$/, ' data-cell-split="allow">');
    try { paragraphProperties = JSON.parse(doc.getCellParaPropertiesAt(section, para, control, c.index, 0)); paragraphStyle = paragraphCss(paragraphProperties, JSON.parse(doc.getCellCharPropertiesAt(section, para, control, c.index, 0, 0)),
      c.rowSpan === 1 ? px(c.height - c.paddingTop - c.paddingBottom) : undefined, false); }
    catch { paragraphStyle = 'line-height:100%;font-size:10pt;'; paragraphProperties = {}; }
    const centered = centeredLabels.get(c.index);
    if (centered) tag = tag.replace(/>$/, ` ${hwpCenteredCellLabelsAttribute}="${centered}">`);
    const picture = cellRasterFills.get(c.index);
    if (picture) tag = tag.replace(/>$/, ` ${hwpCellRasterFillAttribute}="${picture}">`);
    const width = columns.slice(c.col, c.col + c.colSpan).map(n => Math.max(1, Math.round(n)));
    const borders = hwpCellBorders(c);
    let noWrap = false;
    if (doc.getCellParagraphCount(section, para, control, c.index) === 1) {
      const text = doc.exportSelectionInCellHtml(section, para, control, c.index, 0, 0, 0, 0xffffffff).replace(/<[^>]*>|<!--.*?-->/g, '').trim();
      noWrap = /^\d+[.)]?$/.test(text);
    }
    // TH has app defaults, but resetting them must not overwrite a fill or
    // text style already supplied by the HWP cell. Put defaults first.
    if (header) tag = tag.includes('style="')
      ? tag.replace('style="', 'style="background-color:transparent;color:#000;font-weight:normal;')
      : style(tag, 'background-color:transparent;color:#000;font-weight:normal;');
    return style(tag.replace(/^<td/, header ? '<th' : '<td').replace(/>$/, ` colwidth="${width.join(',')}">`), `width:${px(c.width)}px;height:${px(c.height)}px;padding:${px(c.paddingTop)}px ${px(c.paddingRight)}px ${px(c.paddingBottom)}px ${px(c.paddingLeft)}px;vertical-align:${['top', 'middle', 'bottom'][c.verticalAlign] ?? 'top'};${noWrap ? 'white-space:nowrap;' : ''}${borders}`);
  });
  if (nativeGrid) result = result.replace(/^<table\b/, `<table ${hwpNativeGridAttribute}="${nativeGrid}"`);
  return runRepairs.length ? result.replace(/(<p\b[^>]*?) data-hwp-run-repair="(\d+)"([^>]*>)[\s\S]*?(<\/p>)/g,
    (_match, start, index, tail, end) => start + tail + runRepairs[Number(index)] + end) : result;
}

function pageSettings(doc: HwpDocument, section: number, warnings: Set<string>): PageSettings {
  const p: PageDef = JSON.parse(doc.getPageDef(section));
  const letter = Math.abs(p.width - 61200) < 150 && Math.abs(p.height - 79200) < 150;
  const custom = !letter && (Math.abs(p.width - 59528) > 150 || Math.abs(p.height - 84188) > 150);
  const column = JSON.parse(doc.getColumnDef(section));
  const hideEmptyLines = JSON.parse(doc.getSectionDef(section)).hideEmptyLine === true;
  if (hideEmptyLines) warnings.add(emptyLineCompatibilityWarning);
  const page: PageSettings = { size: custom ? 'Custom' : letter ? 'Letter' : 'A4', ...(custom ? { width: mm(p.width), height: mm(p.height) } : {}), orientation: p.landscape ? 'landscape' : 'portrait',
    ...(hideEmptyLines ? { hideEmptyLines: true } : {}),
    ...(column.columnCount > 1 ? { columns: { count: column.columnCount, space: px(column.spacing), balance: column.columnType === 1 } } : {}),
    margin: mm(p.marginLeft), margins: { top: mm(p.marginTop + p.marginHeader), right: mm(p.marginRight), bottom: mm(p.marginBottom + p.marginFooter), left: mm(p.marginLeft) },
    header: '', footer: '', pageNumbers: false, headerDistance: mm(p.marginTop), footerDistance: mm(p.marginBottom), gutter: mm(p.marginGutter) };
  const originalMargins = parseHwpHorizontalMargins([p.marginLeft, p.marginRight]);
  if (originalMargins?.some(value => value !== null && hu(mm(value)) !== value)) page.hwpHorizontalMargins = originalMargins;
  for (const isHeader of [true, false]) {
    const hf = JSON.parse(doc.getHeaderFooter(section, isHeader, 0));
    if (hf.exists) {
      page[isHeader ? 'header' : 'footer'] = hf.text ?? '';
    }
  }
  return page;
}

/** The engine's selection HTML intentionally omits controls. Interleave their
 * HTML at the raw paragraph offsets instead of losing tables and pictures. */
export function editableHwp(doc: HwpDocument, onProgress?: HwpProgressReporter, metadata?: HwpLayoutMetadata, noteStory = false, captionDepth = 0): EditableHwp {
  if (metadata?.conversionBlock) throw new Error(metadata.conversionBlock);
  onProgress?.({ label: '문서의 표·그림 구조를 확인하는 중…' });
  const warnings = new Set<string>(metadata?.warnings);
  const controls: Control[] = JSON.parse(doc.getControls());
  const splitCellLines = new Map<string, HwpPaintedLine[]>();
  const bodyLines = new Map<string, HwpPaintedLine[]>();
  const pitches = sourceLinePitches(doc, splitCellLines, bodyLines);
  let equations: ReturnType<typeof cellContents> | undefined;
  let parts: ReturnType<typeof serializedParts> | undefined;
  const serialized = () => {
    if (parts) return parts;
    try { return parts = serializedParts(doc); }
    catch { warnings.add('일부 개체의 보조 서식 정보를 읽지 못했습니다. 원본 보기와 비교하세요.'); return parts = {}; }
  };
  const equationMetadata = () => {
    if (!equations) {
      equations = cellContents(doc, serialized());
      if ([...equations.values()].some(value => value.controls.some(c => c.kind === 'fieldBegin' && c.fieldType === 'BOOKMARK')))
        warnings.add('범위 책갈피는 시작 위치의 책갈피로 가져옵니다.');
    }
    return equations;
  };
  const captionDefinitions = tableCaptionDefinitions(serialized());
  if (captionDefinitions.length && captionDepth >= 8) throw Error('표 캡션이 너무 깊게 중첩되어 안전하게 변환하지 못했습니다.');
  const captionKeys = new Set(captionDefinitions.map(c => c.key));
  const noteStories = noteStory ? { notes: [], warnings: [] } : importHwpNotes(doc, serialized(), noteDoc => editableHwp(noteDoc, undefined, undefined, true));
  for (const warning of noteStories.warnings) warnings.add(warning);
  const noteReference = (control: { noteId?: string }) => {
    const note = noteStories.notes.find(n => n.id === control.noteId);
    if (!note) return;
    const mark = parseHwpFootnoteLayout(note.layout)?.mark;
    return `<span data-note-reference="${note.id}" data-note-kind="${note.kind}" data-note-label="${escape(note.label)}"${mark ? ` ${hwpFootnoteMarkAttribute}="${mark}"` : ''}>${escape(mark ?? note.label)}</span>`;
  };
  const equationHtml = (equation: { script: string; fontSize: number; color: number; width: number; height: number }) => {
    const svg = doc.renderEquationPreview(equation.script, equation.fontSize || 1200, equation.color || 0);
    const width = equation.width > 0 ? px(equation.width) : Number(svg.match(/\bwidth="([\d.]+)"/)?.[1]);
    const height = equation.height > 0 ? px(equation.height) : Number(svg.match(/\bheight="([\d.]+)"/)?.[1]);
    if (!(width > 0 && height > 0)) throw new Error('Empty equation preview');
    warnings.add('수식은 원래 모양의 그림으로 가져옵니다. 수식 명령어는 그림 설명에 보존하지만 수식 자체는 편집할 수 없습니다.');
    return `<img data-inline-image="true" data-hwp-svg="${escape(svg)}" src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQIHWP4z8DwHwAFgAI/ScLbtAAAAABJRU5ErkJggg==" width="${width}" height="${height}" alt="${escape('수식: ' + equation.script)}" />`;
  };
  let controlPositions: Map<string, any> | undefined;
  const positionKey = (section: number, para: number, control: number, path: CellPath = []) => `${section}:${para}:${control}:${JSON.stringify(path)}`;
  const controlPosition = (section: number, para: number, control: number, path: CellPath = []) => {
    if (!controlPositions) {
      controlPositions = new Map();
      for (let page = 0; page < doc.pageCount(); page++) for (const c of JSON.parse(doc.getPageControlLayout(page)).controls ?? []) {
        // Header/footer paragraph indexes are local to their own story. They
        // can equal body indexes and must never replace the body's anchor.
        if (c.headerFooter || (c.parentParaIdx !== undefined && !c.cellPath?.length)) continue;
        const key = positionKey(c.secIdx, c.paraIdx, c.controlIdx, c.cellPath ?? []);
        if (!controlPositions.has(key)) controlPositions.set(key, { ...c, page });
      }
    }
    return controlPositions.get(positionKey(section, para, control, path));
  };
  const floatingGeometry = (position: any, props: any) => ({
    horizontalPosition: { relative: 'page', offset: Math.round(position.x * 9525) },
    verticalPosition: { relative: 'page', offset: Math.round(position.y * 9525) },
    wrap: { type: ['BehindText', 'InFrontOfText'].includes(props.textWrap) ? 0 : props.textWrap === 'TopAndBottom' ? 3 : 1 },
    // The native layout API already resolved these coordinates against the
    // physical paper. Word's layoutInCell=true would add the cell origin again.
    behindDocument: props.textWrap === 'BehindText', allowOverlap: props.allowOverlap, layoutInCell: false,
    margins: Object.fromEntries(['left', 'right', 'top', 'bottom'].map(side => [side, Math.round((props['outerMargin' + side[0].toUpperCase() + side.slice(1)] || 0) * 127)])),
  });
  const overlays = new Map<number, any[]>();
  const metafilePreviews = new Map<string, string>();
  const croppedPicture = (html: string, props: any) => {
    html = metafilePreview(html, doc, metafilePreviews);
    const crop = [props.cropLeft / props.originalWidth, props.cropTop / props.originalHeight,
      props.cropRight / props.originalWidth, props.cropBottom / props.originalHeight];
    if (crop.every(n => Number.isFinite(n) && n >= 0 && n < 1) && crop.some(n => n > 0) && crop[0] + crop[2] < 1 && crop[1] + crop[3] < 1)
      return html.replace(/<img\b/, `<img data-hwp-image-crop="${crop.join(',')}"`);
    return html;
  };
  const pictureHtml = (section: number, para: number, control: Control): string | undefined => {
    let props: any;
    try { props = JSON.parse(doc.getPictureProperties(section, para, control.controlIndex)); } catch { return; }
    if (props.treatAsChar) return croppedPicture(fragment(doc.exportControlHtml(section, para, '[]', control.controlIndex)), props);
    const position = controlPosition(section, para, control.controlIndex);
    if (!position) return;
    let images = overlays.get(position.page);
    if (!images) {
      const page = JSON.parse(doc.getPageOverlayImages(position.page));
      images = [...page.behind ?? [], ...page.front ?? []]; overlays.set(position.page, images);
    }
    const image = images.find(i => Math.abs(i.bbox.x - position.x) < 1 && Math.abs(i.bbox.y - position.y) < 1 && Math.abs(i.bbox.width - position.w) < 1);
    let proof: ReturnType<typeof sourceBodyPictureFlow>;
    if (props.textWrap === 'Square' && props.vertRelTo === 'Para' && !props.outerMarginTop && !props.outerMarginBottom &&
        !props.outerMarginLeft && !props.outerMarginRight) {
      try { proof = sourceBodyPictureFlow(doc, section, para, position, pitches); }
      catch { /* Missing source metrics must not suppress the picture. */ }
    }
    let paperDrawing: { width: number; height: number } | undefined;
    if (['BehindText', 'InFrontOfText'].includes(props.textWrap) && props.horzRelTo === 'Paper' && props.vertRelTo === 'Paper' &&
        position.x === 0 && position.y === 0 && !props.outerMarginTop && !props.outerMarginBottom &&
        !props.outerMarginLeft && !props.outerMarginRight) {
      try {
        const paper = JSON.parse(doc.getPageInfo(position.page));
        const text = doc.getTextRange(section, para, 0, 0xffffffff);
        if (Math.abs(position.w - paper.width) <= .16 && Math.abs(position.h - paper.height) <= .16 &&
            !text.replace(/[\s\p{Cc}\ufffc]/gu, ''))
          paperDrawing = { width: position.w, height: position.h };
      } catch { /* Optional paper geometry must not suppress a valid image. */ }
    }
    const geometry = { ...floatingGeometry(position, props), ...(proof ? { hwpBodyFlow: proof } : {}),
      ...(paperDrawing ? { hwpPaperDrawing: paperDrawing } : {}) };
    if (image && /^image\/(png|jpeg|bmp|gif|webp)$/.test(image.mime))
      return `<img data-inline-image="true" data-image-floating="${escape(JSON.stringify(geometry))}" src="data:${image.mime};base64,${image.base64}" width="${position.w}" height="${position.h}" alt="${escape(props.description || '그림')}" />`;
    // Square/top-and-bottom pictures are part of text flow, not overlay
    // planes. Their asset is still available from the individual control.
    const raw = fragment(doc.exportControlHtml(section, para, '[]', control.controlIndex)).match(/<img\b[^>]*>/)?.[0];
    if (!raw) return;
    return croppedPicture(raw.replace(/\s(?:width|height|alt|data-inline-image)="[^"]*"/g, '').replace(/\s*\/?>$/, ` data-inline-image="true" data-image-floating="${escape(JSON.stringify(geometry))}" width="${position.w}" height="${position.h}" alt="${escape(props.description || '그림')}" />`), props);
  };
  let layerPage = -1, layerRoot: any;
  const pageLayer = (page: number) => {
    if (layerPage !== page) {
      const tree = JSON.parse(doc.getPageLayerTreeWithProfile(page, 'print', true, true));
      if (tree.schemaVersion !== 1 || tree.unit !== 'px' || tree.coordinateSystem !== 'page-top-left-y-down') return;
      layerPage = page; layerRoot = tree.root;
    }
    return layerRoot;
  };
  const cellShapePosition = (section: number, para: number, path: CellPath, index: number, props: any) => {
    const position = controlPosition(section, para, index, path);
    if (position || !props.treatAsChar) return position;
    // The control-layout API omits inline shapes in nested tables. A unique
    // own-size subtree enclosed by the addressed source cell is unambiguous;
    // never crop the cell or take a similarly sized object elsewhere.
    const boxes = JSON.parse(doc.getTableCellBboxesByPath(section, para, JSON.stringify(path)));
    for (const box of boxes.filter((b: any) => b.cellIdx === path.at(-1)!.cellIndex)) {
      const matches: any[] = [];
      const find = (node: any, depth = 0) => {
        if (!node || depth > 40) return;
        const b = node.bounds;
        if (b && Math.abs(b.width - px(props.width)) < 1 && Math.abs(b.height - px(props.height)) < 1 &&
          b.x >= box.x - 1 && b.y >= box.y - 1 && b.x + b.width <= box.x + box.w + 1 && b.y + b.height <= box.y + box.h + 1) { matches.push(node); return; }
        node.children?.forEach((n: any) => find(n, depth + 1)); if (node.child) find(node.child, depth + 1);
      };
      find(pageLayer(box.pageIndex));
      if (matches.length === 1) { const b = matches[0].bounds; return { type: 'shape', page: box.pageIndex, x: b.x, y: b.y, w: b.width, h: b.height }; }
    }
  };
  let originalObjectImageSize: ReturnType<typeof objectImageSizes> | undefined;
  const objectImageBytes = new Map<string, Uint8Array>();
  let cachedObjectImageBytes = 0;
  const objectLayer = (position: any, cell: boolean, subtree?: any) => {
    if (!position || !['shape', 'line', 'ellipse', 'polygon', 'group'].includes(position.type)) return;
    if (!pageLayer(position.page)) return;
    // Select one object subtree, never crop a page: a crop would also copy
    // unrelated text and background pictures underneath the shape.
    const matches: any[] = subtree ? [subtree] : [];
    const find = (node: any, depth = 0) => {
      if (!node || depth > 40) return;
      const b = node.bounds;
      if ((cell || node.layer?.zOrder === position.zOrder) && b &&
        [b.x - position.x, b.y - position.y, b.width - position.w, b.height - position.h].every(n => Math.abs(n) < 1)) {
        matches.push(node); return;
      }
      node.children?.forEach((child: any) => find(child, depth + 1));
      if (node.child) find(node.child, depth + 1);
    };
    if (!subtree) find(layerRoot);
    return matches.length === 1 ? matches[0] : undefined;
  };
  const objectHtml = (position: any, props: any, label: string, cell = false, subtree?: any): string | undefined => {
    try {
      const layer = objectLayer(position, cell, subtree);
      if (!layer) return;
      const node = structuredClone(layer), text: string[] = [];
      let count = 0, stroke = 0, needsBleed = false;
      const supported = (node: any, depth = 0): boolean => {
        if (++count > 10000 || depth > 40) return false;
        if (node.kind === 'group') return Array.isArray(node.children) && node.children.every((child: any) => supported(child, depth + 1));
        if (node.kind === 'clipRect') return supported(node.child, depth + 1);
        return node.kind === 'leaf' && Array.isArray(node.ops) && node.ops.every((op: any) => {
          if (++count > 10000) return false;
          stroke = Math.max(stroke, Number(op.style?.strokeWidth ?? op.style?.width) || 0);
          if (op.type === 'textRun') {
            const value = op.displayText ?? op.text ?? ''; text.push(value);
            if (value.trim()) needsBleed = true;
          } else if (!['rectangle', 'image'].includes(op.type)) needsBleed = true;
          if (op.transform?.rotation) needsBleed = true;
          if (op.type === 'image') {
            if (!/^image\/(?:png|jpeg|gif|webp|bmp)$/.test(op.mime) || op.effect && op.effect !== 'realPic' || op.brightness || op.contrast) return false;
            let bytes: Uint8Array;
            if (typeof op.base64 === 'string') {
              if (op.base64.length > 4 * 1024 * 1024 || !/^[A-Za-z0-9+/]+=*$/.test(op.base64)) return false;
              bytes = Uint8Array.from(atob(op.base64), ch => ch.charCodeAt(0));
            } else if (typeof op.sourceImageKey === 'string') {
              bytes = objectImageBytes.get(op.sourceImageKey) ?? doc.getSourceImageBytes(op.sourceImageKey);
              if (bytes.length > 3 * 1024 * 1024) return false;
              if (!objectImageBytes.has(op.sourceImageKey) && cachedObjectImageBytes + bytes.length <= 4 * 1024 * 1024) {
                objectImageBytes.set(op.sourceImageKey, bytes); cachedObjectImageBytes += bytes.length;
              }
              let binary = '';
              for (let i = 0; i < bytes.length; i += 8192) binary += String.fromCharCode(...bytes.subarray(i, i + 8192));
              op.base64 = btoa(binary);
            } else return false;
            if (op.crop && !op.originalSizeHu) {
              originalObjectImageSize ??= objectImageSizes(doc);
              op.originalSizeHu = originalObjectImageSize(bytes);
              if (!op.originalSizeHu) return false;
            }
            delete op.imageBytesOmitted;
            return true;
          }
          if (op.type === 'tabLeader') return op.leadersComplete === true && !op.rotation && !op.isVertical &&
            Array.isArray(op.leaders) && op.leaders.length <= 1000 && op.leaders.every((leader: any) => leader.fillType === 3);
          return ['rectangle', 'ellipse', 'line', 'polygon', 'path', 'textRun'].includes(op.type);
        });
      };
      if (!supported(node)) return;
      // Inline image-only groups have no ink beyond their declared bounds.
      // Inflating them by the generic preview padding can create a logo-only
      // page even though the original object fits its remaining line budget.
      const padding = cell && props.treatAsChar && !needsBleed && stroke === 0 ? 0 : Math.max(2, Math.ceil(stroke / 2 + 1));
      if (!Number.isFinite(padding) || padding > 1000) return;
      const payload = JSON.stringify({ ...node, previewPadding: padding });
      if (payload.length > 4 * 1024 * 1024) return;
      const b = node.bounds;
      const paperDrawing = !cell && !props.treatAsChar && props.horzRelTo === 'Paper' &&
        props.vertRelTo === 'Paper' && props.textWrap === 'TopAndBottom'
        ? { width: b.width + padding * 2, height: b.height + padding * 2 } : undefined;
      const geometry = cell ? { ...floatingGeometry(position, props), layoutInCell: true,
        horizontalPosition: { relative: 'column', offset: Math.round((px(props.horzOffset || 0) - padding) * 9525) },
        verticalPosition: { relative: 'paragraph', offset: Math.round((px(props.vertOffset || 0) - padding) * 9525) } }
        : { ...floatingGeometry({ x: b.x - padding, y: b.y - padding }, props),
          ...(paperDrawing ? { hwpPaperDrawing: paperDrawing } : {}) };
      warnings.add('도형·글상자는 원래 위치와 모양의 그림으로 가져옵니다. 글상자 내용은 그림 설명에 보존하지만 개체 내부는 편집할 수 없습니다.');
      return `<img data-inline-image="true" data-hwp-object="${escape(payload)}"${cell && props.treatAsChar ? '' : ` data-image-floating="${escape(JSON.stringify(geometry))}"`} src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQIHWP4z8DwHwAFgAI/ScLbtAAAAABJRU5ErkJggg==" width="${b.width + padding * 2}" height="${b.height + padding * 2}" alt="${escape([label || '도형', text.join('')].filter(Boolean).join(': '))}" />`;
    } catch { return; }
  };
  const shapeHtml = (section: number, para: number, control: Control, allowEditable?: 'both' | 'nested'): string | undefined => {
    try {
      const props = JSON.parse(doc.getShapeProperties(section, para, control.controlIndex));
      const position = controlPosition(section, para, control.controlIndex);
      if (allowEditable && position?.type === 'shape' && (!props.treatAsChar || allowEditable === 'both')) {
        try {
          const layer = objectLayer(position, !!props.treatAsChar);
          const convert = props.treatAsChar ? sourceInlineTextBox : sourceFloatingNestedTextBox;
          const editable = layer && convert(doc, section, para, control.controlIndex, position.page, layer, props);
          if (editable) return editable;
        } catch { /* The original raster remains the fallback for an unsupported text box. */ }
      }
      return objectHtml(position, props, control.userDesc, !!props.treatAsChar);
    } catch { return; }
  };
  const storyChanges = new Map<string, Partial<PageSettings>>(), restoredHeaders = new Set<string>();
  const definitions = headerDefinitions(serialized());
  const editableStories = editableHeaderStories(doc, definitions, isolated => editableHwp(isolated, undefined, undefined, true).html);
  // The raw selection exporter omits rich header stories. Reuse the isolated
  // header layer (never a page crop), with PAGE controls kept as live fields.
  for (const definition of definitions) {
    const key = `${definition.section}:${definition.para}`, page = pageSettings(doc, definition.section, warnings);
    const names = definition.apply === 0 ? [definition.kind, definition.kind === 'header' ? 'evenHeader' : 'evenFooter']
      : [definition.apply === 1 ? definition.kind === 'header' ? 'evenHeader' : 'evenFooter' : definition.kind];
    const text = [...definition.xml.matchAll(/<hp:t>([^<]*)<\/hp:t>/g)].map(m => m[1]).join('');
    const rich = editableStories.has(definition) || /<hp:(?:rect|line|pic|ellipse|polygon|container|tbl|equation|ole|autoNum)\b/.test(definition.xml);
    let html = editableStories.get(definition) ?? (text ? `<p>${text}</p>` : ''), preserved = editableStories.has(definition) || !rich;
    if (rich && !preserved) {
      const index = headerPreviewPage(doc, definition);
      const original = index === undefined ? undefined : headerLayer(pageLayer(index), definition.kind);
      if (original) {
        const layer = JSON.parse(JSON.stringify(original)), fields: any[] = [];
        const fieldCount = [...definition.xml.matchAll(/<hp:autoNum\b[^>]*numType="PAGE"/g)].length;
        const visit = (node: any) => {
          for (const op of node.ops ?? []) if (op.type === 'textRun' && op.displayText && op.displayText !== op.text) {
            const original = Array.from(op.text as string), display = Array.from(op.displayText as string);
            let from = 0, suffix = 0;
            while (from < Math.min(original.length, display.length) && original[from] === display[from]) from++;
            while (suffix < Math.min(original.length, display.length) - from && original.at(-suffix - 1) === display.at(-suffix - 1)) suffix++;
            const end = display.length - suffix, positions = op.displayPositions ?? op.positions;
            if (!original.slice(from, original.length - suffix).join('').trim() && /^\d+$/.test(display.slice(from, end).join('')) && positions?.length === display.length + 1) {
              fields.push({ op, from, end, bbox: { ...op.bbox, x: op.bbox.x + positions[from], width: positions[end] - positions[from] } });
            }
          }
          node.children?.forEach(visit); if (node.child) visit(node.child);
        };
        visit(layer);
        // Require an exact count; symbol projections must not be mistaken for
        // page fields in a story with unrelated special characters.
        if (fields.length === fieldCount) {
          const imageCarrierHeight = 1;
          const footerPaper = definition.kind === 'footer' && index !== undefined && fields.length === 1
            ? JSON.parse(doc.getPageInfo(index)) : undefined;
          const dynamic = fields.map(field => {
            const op = field.op, s = op.paintStyle ?? op.style, b = field.bbox;
            const left = b.x - (page.margins?.left ?? page.margin) * 96 / 25.4;
            const top = b.y - (page.headerDistance ?? 0) * 96 / 25.4;
            const footerBottom = footerPaper ? footerPaper.height - (page.footerDistance ?? 0) * 96 / 25.4 : 0;
            const below = footerBottom - b.y - b.height - imageCarrierHeight;
            // A footer is anchored from the paper's bottom. A single field
            // confirmed inside that story can keep its position with the
            // remaining space below it, including the following image carrier.
            const position = footerPaper && b.y >= footerPaper.footerArea?.y && below >= 0
              ? `margin-top:0;margin-bottom:${below}px` : `margin-top:${top}px;margin-bottom:0`;
            // This paragraph contains only the live field. Its paragraph
            // default also supplies the field's font after DOCX reimport.
            return `<p style="margin-left:${left}px;${position};line-height:${b.height}px;font-size:${s.fontSize}px;font-family:${escape(JSON.stringify(s.fontFamily))};color:${s.color};${s.bold ? 'font-weight:bold;' : ''}"><span data-page-number="PAGE">1</span></p>`;
          }).join('');
          for (const field of fields) {
            const glyphs = Array.from(field.op.displayText as string);
            glyphs.fill(' ', field.from, field.end); field.op.displayText = glyphs.join('');
          }
          const b = layer.bounds;
          const image = objectHtml({ type: 'group', page: index, x: b.x, y: b.y, w: b.width, h: b.height }, { textWrap: 'InFrontOfText' }, definition.kind === 'header' ? '머리말' : '꼬리말', false, layer);
          if (image) { html = dynamic + `<p style="line-height:${imageCarrierHeight}px;margin:0;font-size:${imageCarrierHeight}px">${image}</p>`; preserved = true; }
        }
      }
    }
    if (!preserved) { warnings.add('일부 머리말·꼬리말의 개체를 복원하지 못했습니다. 원본 보기와 비교하세요.'); continue; }
    const patch = storyChanges.get(key) ?? { differentFirstPage: false };
    if (definition.apply) patch.differentOddEven = true;
    for (const name of names) { (patch as any)[name] = entities(text); (patch as any)[name + 'Html'] = rich ? html : undefined; }
    storyChanges.set(key, patch); restoredHeaders.add(key);
  }
  for (const hidden of hiddenStories(serialized())) {
    const key = `${hidden.section}:${hidden.para}`;
    const patch: Partial<PageSettings> = { ...storyChanges.get(key), differentFirstPage: true };
    if (hidden.header) { patch.firstHeader = ''; patch.firstHeaderHtml = ''; }
    if (hidden.footer) { patch.firstFooter = ''; patch.firstFooterHtml = ''; patch.firstPageNumbers = false; }
    storyChanges.set(key, patch);
  }
  const byParagraph = new Map<number, Control[]>();
  for (const c of controls) {
    if (c.list === 0) { const list = byParagraph.get(c.para) ?? []; list.push(c); byParagraph.set(c.para, list); }
  }
  const pieces: string[] = [];
  let globalPara = 0, size = 0;
  const append = (value: string, before?: number) => {
    size += hwpHtmlBytes(value) + (pieces.length ? 1 : 0); checkHwpHtmlSize(size);
    if (before === undefined) pieces.push(value); else pieces.splice(before, 0, value);
  };
  const firstPage = pageSettings(doc, 0, warnings);
  Object.assign(firstPage, storyChanges.get('0:0'));
  const initialLayout = metadata?.paragraphs['0:0'];
  const pageNumberStory = !noteStory && initialLayout && !initialLayout.hideFirstPageNumber
    ? initialPageNumberStory(doc, controls, serialized(), () => headerLayer(pageLayer(0), 'footer')) : undefined;
  const restorePageNumber = (page: PageSettings) => {
    if (!pageNumberStory) return;
    page.pageNumbers = true; page.footerHtml = pageNumberStory;
    if (page.differentFirstPage) { page.firstPageNumbers = true; page.firstFooterHtml = pageNumberStory; }
    if (page.differentOddEven) { page.evenPageNumbers = true; page.evenFooterHtml = pageNumberStory; }
  };
  restorePageNumber(firstPage);
  if (firstPage.headerHtml === undefined) {
    const blankHeader = initialBlankHeaderStory(doc, definitions, editableStories, firstPage.header, () => !!headerLayer(pageLayer(0), 'header'));
    if (blankHeader) { firstPage.header = ''; firstPage.headerHtml = blankHeader; }
  }
  const total = Array.from({ length: doc.getSectionCount() }, (_, section) => doc.getParagraphCount(section)).reduce((sum, count) => sum + count, 0);
  const fontRuns = new Map<number, HwpScriptTypography>();
  onProgress?.({ label: '본문·표·그림을 변환하는 중…', completed: 0, total, unit: '문단' });
  for (let section = 0; section < doc.getSectionCount(); section++) {
    let currentPage = section ? { ...pageSettings(doc, section, warnings), ...storyChanges.get(`${section}:0`) } : firstPage;
    if (section) append(`<div data-section-page="${escape(JSON.stringify(currentPage))}"></div>`);
    const column = JSON.parse(doc.getColumnDef(section));
    let columnCount = column.columnCount ?? column.count ?? 1, columnSpace = px(column.spacing ?? 0), columnBalance = column.columnType === 1;
    const hasColumns = columnCount > 1 || Object.entries(metadata?.paragraphs ?? {}).some(([key, p]) => key.startsWith(section + ':') && (p.columns?.count ?? 1) > 1);
    let columnOpen = false;
    const openColumns = () => { if (hasColumns && !columnOpen) { append(`<div data-column-count="${columnCount}" data-column-space="${columnSpace}" data-column-balance="${columnBalance}">`); columnOpen = true; } };
    const closeColumns = () => { if (columnOpen) { append('</div>'); columnOpen = false; } };
    openColumns();
    let lastTopParagraph = -2, blankAfterFloatingTable = -1;
    const paragraphCount = doc.getParagraphCount(section), firstGlobal = globalPara;
    let borderGroup = { end: -1, attribute: '' };
    const continuesBorder = (index: number) => {
      const key = `${section}:${index}`, flow = metadata?.paragraphs[key];
      return !storyChanges.has(key) && !flow?.pageBreak && !flow?.columnBreak && !flow?.columns;
    };
    const supportsBorder = (index: number) => !byParagraph.has(firstGlobal + index) &&
      !!doc.getTextRange(section, index, 0, 0xffffffff).trim();
    for (let para = 0; para < paragraphCount; para++, globalPara++) {
      const properties = JSON.parse(doc.getParaPropertiesAt(section, para));
      const structure = metadata?.paragraphs[`${section}:${para}`];
      const startsNewFlow = properties.pageBreakBefore || structure?.pageBreak || structure?.columnBreak || structure?.columns;
      if (structure?.columns) { closeColumns(); columnCount = structure.columns.count; columnSpace = structure.columns.space; columnBalance = !!structure.columns.balance; openColumns(); }
      if (properties.pageBreakBefore || structure?.pageBreak) { closeColumns(); append('<div data-page-break="true"></div>'); openColumns(); }
      else if (structure?.columnBreak) append('<div data-column-break="true"></div>');
      const stories = storyChanges.get(`${section}:${para}`);
      if (para && stories) {
        closeColumns(); currentPage = { ...currentPage, ...stories, breakType: 'continuous' };
        restorePageNumber(currentPage);
        append(`<div data-section-page="${escape(JSON.stringify(currentPage))}"></div>`); openColumns();
      }
      // Some forms are editable text tables over a paper/background picture
      // in the same anchor paragraph. Stacking both blocks doubles their
      // height. Preserve an unambiguous one-cell overlay as a local backdrop.
      const tableBackgrounds = new Map<number, { html: string; original: string }>(), backgroundControls = new Set<number>();
      const paragraphControls = byParagraph.get(globalPara) ?? [];
      let relatedTableLayout: string | undefined;
      if (para > 0 && !noteStory && !hasColumns && !startsNewFlow && !stories &&
        paragraphControls.length === 2 && paragraphControls.every(c => c.ctrlId === 'tbl') &&
        !doc.getTextRange(section, para, 0, 0xffffffff).trim()) {
        const previous = byParagraph.get(globalPara - 1);
        const [left, right] = paragraphControls.map(c => controlPosition(section, para, c.controlIndex));
        if (previous?.length === 1 && previous[0].ctrlId === 'tbl' && left && right &&
          !doc.getTextRange(section, para - 1, 0, 0xffffffff).trim()) {
          relatedTableLayout = sourceRelatedTableLayout(doc, section, para, previous[0].controlIndex,
            [paragraphControls[0].controlIndex, paragraphControls[1].controlIndex], [left, right], splitCellLines);
        }
      }
      for (const table of paragraphControls.filter(c => c.ctrlId === 'tbl')) {
        try {
          const props = JSON.parse(doc.getTableProperties(section, para, table.controlIndex));
          const dims = JSON.parse(doc.getTableDimensions(section, para, table.controlIndex));
          if (props.treatAsChar || props.textWrap !== 'InFrontOfText' || dims.cellCount !== 1) continue;
          const cell = JSON.parse(doc.getCellProperties(section, para, table.controlIndex, 0));
          const position = controlPosition(section, para, table.controlIndex); if (!position) continue;
          const paragraph = JSON.parse(doc.getCellParaPropertiesAt(section, para, table.controlIndex, 0, 0));
          // A paragraph-relative picture follows the text, not the cell edge.
          // Middle/bottom alignment is safe only when the source text already
          // fills the cell: otherwise its alignment reserve moves the backdrop.
          if (cell.verticalAlign !== 0) {
            const key = `${section}:${para}:${table.controlIndex}:0:0`;
            const top = pitches.get(`top:${key}`), ink = pitches.get(`height:${key}`);
            const char = JSON.parse(doc.getCellCharPropertiesAt(section, para, table.controlIndex, 0, 0, 0));
            const line = pitches.get(key) ?? physicalLineHeight(paragraph, char);
            const reserve = top === undefined || ink === undefined ? NaN
              : top - position.y - px(cell.paddingTop) - (paragraph.spacingBefore || 0) / 2 - Math.max(0, line - ink) / 2;
            if (!Number.isFinite(reserve) || Math.abs(reserve) > 2 || pitches.get(`page:${key}`) !== position.page) continue;
          }
          const candidates = paragraphControls.filter(c => c.ctrlId === 'gso' && !backgroundControls.has(c.controlIndex)).flatMap(c => {
            try {
              const picture = JSON.parse(doc.getPictureProperties(section, para, c.controlIndex)), box = controlPosition(section, para, c.controlIndex);
              if (!picture.treatAsChar || !box || box.page !== position.page || box.type !== 'image') return [];
              const overlap = Math.max(0, Math.min(box.x + box.w, position.x + position.w) - Math.max(box.x, position.x)) *
                Math.max(0, Math.min(box.y + box.h, position.y + position.h) - Math.max(box.y, position.y));
              if (overlap < Math.max(box.w * box.h, position.w * position.h) * .9) return [];
              const image = fragment(doc.exportControlHtml(section, para, '[]', c.controlIndex)).match(/<img\b[^>]*>/)?.[0];
              return image ? [{ c, box, image }] : [];
            } catch { return []; }
          });
          if (candidates.length !== 1) continue;
          const { c, box, image } = candidates[0];
          const geometry = { horizontalPosition: { relative: 'column', offset: Math.round((box.x - position.x - px(cell.paddingLeft)) * 9525) },
            verticalPosition: { relative: 'paragraph', offset: Math.round((box.y - position.y - px(cell.paddingTop) - (paragraph.spacingBefore || 0) / 2) * 9525) },
            wrap: { type: 0 }, behindDocument: true, layoutInCell: true, allowOverlap: true };
          tableBackgrounds.set(table.controlIndex, { original: image, html: image.replace(/\s(?:width|height|data-inline-image|data-image-floating)="[^"]*"/g, '').replace(/\s*\/?>$/, ` data-inline-image="true" data-image-floating="${escape(JSON.stringify(geometry))}" width="${box.w}" height="${box.h}" />`) });
          backgroundControls.add(c.controlIndex);
        } catch { /* Ambiguous overlays keep both original controls. */ }
      }
      const positions: number[] = JSON.parse(doc.getControlTextPositions(section, para));
      const textLength = doc.getParagraphLength(section, para);
      const wrapKey = `${section}:${para}`;
      const wrapPage = pitches.get(`page:${wrapKey}`), wrapX = pitches.get(`x:${wrapKey}`), wrapY = pitches.get(`top:${wrapKey}`);
      const wrapOrigin = columnCount === 1 && properties.indent < 0 && !pitches.has(`cross-page:${wrapKey}`) &&
        wrapPage !== undefined && wrapX !== undefined && wrapY !== undefined && wrapX >= 0 && wrapY >= 0
        ? { page: wrapPage + 1, x: wrapX, y: wrapY } : undefined;
      const recordBodyWidth = columnCount === 1 && properties.alignment === 'justify' && properties.indent <= 0 &&
        !pitches.has(`cross-page:${wrapKey}`);
      const following = para + 1 < paragraphCount ? byParagraph.get(globalPara + 1) : undefined;
      const next = following?.length === 1 ? following[0] : undefined;
      // A stored heading immediately before an inline table can keep that
      // table on the authored sheet, where an extra fallback line cannot.
      const singleBodyLine = !noteStory && !paragraphControls.length && !startsNewFlow && recordBodyWidth && wrapOrigin &&
        pitches.get(`lines:${wrapKey}`) === 1 && next?.ctrlId === 'tbl' && (() => {
          try {
            const table = JSON.parse(doc.getTableProperties(section, para + 1, next.controlIndex));
            const grid = JSON.parse(doc.getTableDimensions(section, para + 1, next.controlIndex));
            const box = controlPosition(section, para + 1, next.controlIndex);
            return table.treatAsChar && !table.hasCaption && grid.cellCount === 1 && box &&
              box.page + 1 === wrapOrigin.page && box.y > wrapOrigin.y;
          } catch { return false; }
        })();
      const range = (start: number, end: number) => {
        const html = fragment(doc.exportSelectionHtml(section, para, start, para, end));
        // Control positions can extend past a control-only paragraph's text.
        // getTextRange takes a count, whereas selection HTML takes an end.
        if (start >= textLength || end <= start) return html;
        const text = doc.getTextRange(section, para, start, Math.min(end, textLength) - start);
        const withBreaks = restoreHardLineBreaks(html, [...text.matchAll(/\n/g)].map(m => m.index!));
        const runs = sourceFontRuns(doc.getCharShapeRuns(section, para, start, Math.min(end, textLength)), start, text,
          offset => JSON.parse(doc.getCharPropertiesAt(section, para, offset)), fontRuns, true);
        const scripts = restoreScriptFonts(withBreaks, text, runs, columnCount > 1, metadata?.latinFaces);
        const formatted = start === 0 && end >= textLength && positions.length === 0
          ? preserveHwpSourceSpaces(scripts, text, runs, metadata?.fontSpace,
            () => doc.getParagraphLineMetrics(section, para, '[]'), wrapOrigin, recordBodyWidth, !!singleBodyLine) : scripts;
        let result = formatted;
        if (text.includes('\u2007')) {
          const prefix = start ? doc.getTextRange(section, para, 0, start).length : 0;
          const fixed = (equationMetadata().get(`${section}:${para}`)?.fixedSpaces ?? [])
            .filter(offset => offset >= prefix && offset < prefix + text.length).map(offset => offset - prefix);
          result = restoreHwpFixedSpaces(formatted, text, fixed);
        }
        return columnCount === 2 && properties.alignment === 'justify' && start === 0 && end >= textLength &&
            positions.length === 0 && !pitches.has(`cross-page:${wrapKey}`)
          ? preserveHwpColumnLineEnds(result, text, () => doc.getParagraphLineMetrics(section, para, '[]'),
            properties.indent < 0 && properties.lineSpacingType === 'Percent' ? bodyLines.get(wrapKey) : undefined) : result;
      };
      const whole = range(0, 0xffffffff);
      let open = whole.match(/^<p\b[^>]*>/)?.[0] ?? '<p style="">';
      const painted = bodyLines.get(wrapKey), bodyX = wrapX ?? pitches.get(`leading-x:${wrapKey}`);
      const tailRightMargin = (() => {
        if (!recordBodyWidth || !properties.marginRight || properties.marginLeft || painted?.length !== 2 ||
            wrapPage === undefined || pitches.get(`page:${section}:${para + 1}`) !== wrapPage + 1 ||
            para + 1 >= paragraphCount || positions.length || properties.lineSpacingType !== 'Percent') return;
        const last = painted[1], paper = JSON.parse(doc.getPageInfo(wrapPage));
        const clearance = paper.footerArea?.y - last.y - last.h;
        const leading = properties.lineSpacing * last.h / 100 - last.h;
        if (Number.isFinite(clearance) && clearance >= 0 && clearance <= leading + .15) return Number(properties.marginRight);
      })();
      if (recordBodyWidth && positions.length === 0 && properties.indent < 0 && painted?.length &&
          wrapPage !== undefined && bodyX !== undefined && wrapY !== undefined && bodyX >= 0 && wrapY >= 0 &&
          painted.every(line => line.page === wrapPage) && !properties.marginLeft && (!properties.marginRight || tailRightMargin) &&
          !properties.keepWithNext && !properties.keepLines && !properties.pageBreakBefore &&
          properties.lineSpacingType === 'Percent' && !properties.fontLineHeight) {
        const text = doc.getTextRange(section, para, 0, textLength);
        const nativeSingle = painted.length === 1 && properties.alignment === 'justify' &&
          text.length >= 12 && text.length <= 160 && /^[ \u00a0]/.test(text)
          ? nativeLineTypography(doc, section, para, textLength) : undefined;
        if (painted.length >= 2 || nativeSingle) {
          const witnessed = preserveHwpBodyLineEnds(whole, text, () => doc.getParagraphLineMetrics(section, para, '[]'), {
            origin: { page: wrapPage + 1, x: bodyX, y: wrapY }, hanging: -properties.indent, painted, tailRightMargin, nativeSingle,
          });
          // Keep the print witness separate from the original fixed-space input.
          open = witnessed.match(/^<p\b[^>]*>/)?.[0] ?? open;
        }
      }
      const afterSpacer = `${section}:${para + 2}`, nextHeadY = pitches.get(`top:${afterSpacer}`);
      if (wrapPage !== undefined && wrapPage > 0 && wrapY !== undefined && nextHeadY !== undefined &&
          para > 0 && para + 2 < paragraphCount && columnCount === 1 && !startsNewFlow && !positions.length &&
          textLength >= 8 && textLength <= 120 && !properties.indent && !properties.marginLeft && !properties.marginRight &&
          pitches.get(`lines:${wrapKey}`) === 1 && pitches.get(`page:${section}:${para - 1}`) === wrapPage - 1 &&
          pitches.get(`page:${afterSpacer}`) === wrapPage && !open.includes(hwpSourceWrapAttribute) &&
          !metadata?.paragraphs[`${section}:${para + 1}`]?.pageBreak && !metadata?.paragraphs[afterSpacer]?.pageBreak) {
        const proof = sourceHwpPageHead(doc, section, para, wrapPage, wrapY, nextHeadY,
          doc.getTextRange(section, para, 0, textLength));
        if (proof) open = open.replace(/>$/, ` ${hwpSourceWrapAttribute}="${escape(JSON.stringify(proof))}">`);
      }
      const blank = !textLength && !paragraphControls.length;
      if (whole.includes('\t')) open = open.replace(/>$/, `${hwpTabAttribute(equationMetadata().get(`${section}:${para}`)?.tabStops)}>`);
      else if (blank && properties.tabStops?.length)
        open = open.replace(/>$/, `${hwpTabAttribute(hwpExplicitTabStops(properties.tabStops))}>`);
      const paragraphChar = JSON.parse(doc.getCharPropertiesAt(section, para, 0));
      open = paragraphAutoSpace(cssStyle(open, paragraphCss(properties, paragraphChar)), properties);
      if (blank && /^#[0-9a-f]{6}$/i.test(paragraphChar.textColor))
        open = cssStyle(open, `color:${paragraphChar.textColor};`);
      if (blankAfterFloatingTable === para && !textLength && !paragraphControls.length && !startsNewFlow && !stories)
        open = cssStyle(open, 'margin-top:0px;');
      if (para > borderGroup.end) borderGroup = connectedHwpBorderGroup(doc, section, para, paragraphCount, properties, continuesBorder, supportsBorder);
      if (borderGroup.attribute) open = open.replace(/>$/, borderGroup.attribute + '>');
      const pitch = pitches.get(`${section}:${para}`);
      if (pitch) open = cssStyle(open, `line-height:${pitch}px;`);
      const sizes = [...whole.matchAll(/font-size:\s*([\d.]+)(pt|px)/gi)].map(m => Number(m[1]) * (m[2].toLowerCase() === 'pt' ? 4 / 3 : 1));
      const largest = sizes.reduce((maximum, size) => Math.max(maximum, size), 0);
      const smallest = sizes.reduce((minimum, size) => Math.min(minimum, size), Infinity);
      const base = pitch ?? physicalLineHeight(properties, paragraphChar);
      const percent = Number(properties.lineSpacing), font = Number(paragraphChar.fontSize) / 75;
      if (properties.lineSpacingType === 'Percent' && !properties.fontLineHeight && percent > 0 &&
          Number.isFinite(base) && base > 0 && font > 0 && sizes.length) {
        const sourceTextLine = percent > 100 && base > font && paragraphControls.length === 0 &&
          Math.abs(largest - font) < .06 && !/(?:\t|<(?:br|img|svg|u|s|del|sub|sup|mark)\b)/i.test(whole);
        const continuation = !startsNewFlow && !stories && lastTopParagraph === para - 1;
        const atTop = sourceTextLine && columnCount === 1 && hwpBodyTextAtLineTop(doc, section, para, font, base,
          continuation ? 'continued' : 'body');
        if (atTop && textLength > 0 && entities(whole.replace(/<[^>]*>/g, '')).trim()) lastTopParagraph = para;
        if (largest - smallest > .1) {
          const witnessed = sourceTextLine && !atTop && columnCount === 1 && !startsNewFlow && !stories;
          const storedSingleLine = witnessed && sourceStoredMixedLine(doc, section, para, pitches, font, base);
          const mixedLinePair = witnessed && !storedSingleLine && !properties.spacingBefore && !properties.spacingAfter && whole.includes(hwpSourceWrapAttribute)
            ? sourceMixedBodyLinePair(doc, section, para, pitches, font, base, percent) : undefined;
          const bodyEndFit = sourceTextLine && atTop && !startsNewFlow && !stories && !properties.spacingBefore &&
            !properties.spacingAfter && sourceTerminalTextAtLineTop(doc, section, para, pitches, font, base)
            ? sourceHwpBodyEndFit(doc, section, para, pitches, pageLayer) : undefined;
          const storedUniformPair = atTop && !startsNewFlow && !stories && whole.includes(hwpSourceWrapAttribute) &&
            sourceUniformMixedLinePair(doc, section, para, pitches, font, base, properties);
          const maxFont = Math.max(largest, font);
          open = cssStyle(open, `--hwp-body-line:${encodeHwpBodyLine({ percent, font: maxFont, base,
            ...(storedSingleLine ? { storedSingleLine: true } : { align: atTop ? 'mixed-top' : undefined, mixedLinePair }),
            ...(bodyEndFit ? { bodyEndFit, maxFont } : {}), ...(storedUniformPair ? { storedUniformPair: true } : {}) })};`);
        } else if (atTop || sourceTextLine && columnCount === 1 && !startsNewFlow && !stories &&
          !Number(properties.spacingBefore) && !Number(properties.spacingAfter) &&
          (continuesBorder(para + 1) && sourceTerminalTextAtLineTop(doc, section, para, pitches, font, base) ||
            !properties.keepLines && !properties.keepWithNext && !properties.indent && !properties.marginLeft && !properties.marginRight &&
            sourceTextAfterInlineTableFits(doc, section, para, pitches, font, base))) {
          open = cssStyle(open, `--hwp-body-line:${encodeHwpBodyLine({ percent, font, base, align: 'top' })};`);
        } else if (sourceTextLine && columnCount > 1) {
          const next = metadata?.paragraphs[`${section}:${para + 1}`];
          if ((next?.pageBreak || next?.columnBreak) && hwpBodyTextAtLineTop(doc, section, para, font, base, 'column-end')) {
            // The source em defines the line origin. Script-relative sizes
            // belong to the original text and are not a later resize.
            const scriptSizes = [...whole.matchAll(/data-hwp-script-size="([\d.]+)pt"/g)]
              .map(m => Number(m[1]) * 4 / 3);
            const maxFont = Math.max(font, Math.round(largest * 1.5) / 1.5, ...scriptSizes);
            const columnFit = columnCount === 2 ? sourceHwpColumnEndFit(doc, section, para, pitches, pageLayer) : undefined;
            open = cssStyle(open, `--hwp-body-line:${encodeHwpBodyLine({ percent, font, base, maxFont, align: 'column-end', columnFit })};`);
          } else if (columnCount === 2 && pitch !== undefined && sourcePreciseColumnPitch(doc, section, para, pitches, font, base, percent)) {
            open = cssStyle(open, `--hwp-body-line:${encodeHwpBodyLine({ percent, font, base, precisePitch: true })};`);
          }
        }
      }
      if (!startsNewFlow && !stories && !noteStory && columnCount === 1 && !positions.length &&
          metadata?.paragraphs[`${section}:${para + 1}`]?.pageBreak && !open.includes(hwpSourceWrapAttribute) &&
          !/(?:\t|<(?:a|br|img|svg|u|s|del|sub|sup|mark)\b)/i.test(whole)) {
        const source = sourceHwpForcedBodyEnd(doc, section, para, pitches, pageLayer);
        if (source) open = open.replace(/>$/, ` ${hwpSourceWrapAttribute}="${escape(JSON.stringify(source))}">`);
      }
      if (!startsNewFlow && !stories && !noteStory && columnCount === 1 && positions.length === 2 &&
          paragraphControls.length === 2 && paragraphControls.every(control => control.ctrlId === 'fn') &&
          !open.includes(hwpSourceWrapAttribute) && !borderGroup.attribute &&
          !/(?:\t|<(?:a|br|img|svg|u|s|del|sub|sup|mark)\b)/i.test(whole)) {
        const noteControls = equationMetadata().get(wrapKey)?.controls.filter(control => control.kind === 'footNote') ?? [];
        const terminalNotes = noteControls.flatMap(control => {
          const note = noteStories.notes.find(n => n.id === control.noteId);
          const layout = parseHwpFootnoteLayout(note?.layout);
          const literal = note?.html && normalizeHwpWrapText(entities(note.html.replace(/<[^>]*>/g, '')));
          return layout && literal ? [{ mark: layout.mark, fingerprint: hwpWrapFingerprint(literal) }] : [];
        });
        if (terminalNotes.length === 2) {
          const source = sourceFootnoteTerminal(doc, section, para, pitches, positions, [terminalNotes[0], terminalNotes[1]]);
          if (source) open = open.replace(/>$/, ` ${hwpSourceWrapAttribute}="${escape(JSON.stringify(source))}">`);
        }
      }
      const inline = (html: string) => html.replace(/^<p\b[^>]*>\s*/, '').replace(/<\/p>\s*$/, '');
      const onlyFloating = (html: string) => html.includes('data-image-floating') && !html.replace(/<[^>]*>/g, '').trim() &&
        [...html.matchAll(/<img\b[^>]*>/g)].every(m => m[0].includes('data-image-floating'));
      const onlyWhitespace = (html: string) => !/<(?:br|img|table|svg)\b|data-(?:note-reference|page-number|symbol-font)/i.test(html) &&
        !entities(html.replace(/<[^>]*>/g, '')).trim();
      let cursor = 0, content = '', contentBytes = 0, lastInlineTable = false, firstBodyControl = true;
      let pictureTableSpacesShareSourceLine = false;
      let tableFollowingHost: number | undefined;
      let hostOnPriorPage: ReturnType<typeof sourceHostOnPriorPage>;
      let inlineTail: { piece: number; line: number; width: number; left: number; available: number; wrapping: boolean; right: number; bottom: number } | undefined;
      let replacedTextBox = false;
      const appendInline = (value: string) => { contentBytes += hwpHtmlBytes(value); checkHwpHtmlSize(size + contentBytes); content += value; };
      for (const c of byParagraph.get(globalPara) ?? []) {
        if (backgroundControls.has(c.controlIndex)) continue;
        if (['secd', 'cold', 'head', 'foot', 'pgnp', 'pgct', 'pghd', 'nwno'].includes(c.ctrlId)) continue;
        const startsWithControl = firstBodyControl;
        firstBodyControl = false;
        const pos = Math.max(cursor, positions[c.controlIndex] ?? cursor);
        appendInline(inline(range(cursor, pos))); cursor = pos;
        // Field starts have no native ctrlId. Resolve only an exact source
        // type/position/count match; an unknown drawing stays explicit.
        if (!c.ctrlId) {
          const unknown = paragraphControls.filter(control => !control.ctrlId && positions[control.controlIndex] === pos);
          const fields = equationMetadata().get(`${section}:${para}`)?.controls.filter(control => control.kind === 'fieldBegin' && control.offset === pos) ?? [];
          const field = unknown.length === fields.length ? fields[unknown.indexOf(c)] : undefined;
          const bookmark = field?.fieldType === 'BOOKMARK' ? hwpBookmark(field.bookmarkName) : undefined;
          if (bookmark) { appendInline(bookmark); continue; }
        }
        if (c.ctrlId === 'bokm') {
          const ordinal = paragraphControls.filter(c => c.ctrlId === 'bokm').indexOf(c);
          const name = equationMetadata().get(`${section}:${para}`)?.controls.filter(c => c.kind === 'bookmark')[ordinal]?.bookmarkName;
          const bookmark = hwpBookmark(name);
          if (bookmark) appendInline(bookmark);
          else warnings.add('일부 책갈피 이름을 가져오지 못했습니다. 본문 내용은 유지합니다.');
          continue;
        }
        onProgress?.({ label: c.ctrlId === 'tbl' ? '표를 변환하는 중…' : c.ctrlId === 'gso' ? '그림·도형을 변환하는 중…' : '문서 개체를 변환하는 중…' });
        let controlHtml = fragment(doc.exportControlHtml(section, para, '[]', c.controlIndex));
        if (!c.ctrlId && /^<!--\s*rhwp:\s*셀 안 CharOverlap\b/.test(controlHtml)) {
          const entries = equationMetadata().get(wrapKey)?.controls;
          const source = entries?.length === positions.length ? entries[c.controlIndex] : undefined;
          if (source?.offset === pos && (source.boxedDigit || source.circledNumber)) {
            const char = JSON.parse(doc.getCharPropertiesAt(section, para, pos));
            controlHtml = source.circledNumber ? circledCompositionHtml(source.circledNumber, char.fontSize)
              : boxedDigitHtml(source.boxedDigit, char.fontSize) ?? controlHtml;
          }
        }
        if (c.ctrlId === 'fn' || c.ctrlId === 'en') {
          const kind = c.ctrlId === 'fn' ? 'footNote' : 'endNote';
          const ordinal = paragraphControls.filter(x => x.ctrlId === c.ctrlId).indexOf(c);
          const note = equationMetadata().get(`${section}:${para}`)?.controls.filter(x => x.kind === kind)[ordinal];
          if (note) controlHtml = noteReference(note) ?? controlHtml;
        }
        if (c.ctrlId === 'atno') {
          const ordinal = paragraphControls.filter(c => c.ctrlId === 'atno').indexOf(c);
          const type = equationMetadata().get(`${section}:${para}`)?.controls.filter(c => c.kind === 'autoNum')[ordinal]?.numberType;
          if (type === 'PAGE' || type === 'TOTAL_PAGE') controlHtml = `<span data-page-number="${type === 'PAGE' ? 'PAGE' : 'NUMPAGES'}">1</span>`;
        }
        if (c.ctrlId === 'gso') {
          const editableBox = !hasColumns && !stories && !properties.spacingBefore && !properties.spacingAfter &&
            paragraphControls.length <= 3 && paragraphControls.every(control => control.ctrlId === 'gso') &&
            (!content.trim() || onlyFloating(content)) && !doc.getTextRange(section, para, 0, 0xffffffff);
          const shapeMode = editableBox ? paragraphControls.length === 1 && !content.trim() ? 'both' : 'nested' : undefined;
          controlHtml = pictureHtml(section, para, c) ?? shapeHtml(section, para, c, shapeMode) ?? controlHtml;
          if (controlHtml.startsWith('<div data-document-textbox=')) {
            append(controlHtml);
            replacedTextBox = true;
            continue;
          }
          if (c.userDesc === 'OLE' && c.props.TreatAsChar) {
            const ordinal = paragraphControls.filter(c => c.userDesc === 'OLE').indexOf(c);
            const info = equationMetadata().get(`${section}:${para}`)?.controls.filter(c => c.kind === 'ole')[ordinal];
            const bytes = info?.binaryItem && serialized()[`BinData/${info.binaryItem}.OLE`];
            const preview = bytes ? olePreview(bytes) : undefined;
            const width = px((c.props as any).Width), height = px((c.props as any).Height);
            if (preview && width > 0 && height > 0 && width <= 10000 && height <= 10000) {
              let binary = ''; for (let i = 0; i < preview.bytes.length; i += 8192) binary += String.fromCharCode(...preview.bytes.subarray(i, i + 8192));
              controlHtml = `<img data-inline-image="true" src="data:image/bmp;base64,${btoa(binary)}" width="${width}" height="${height}" alt="${escape('OLE 미리보기' + (preview.description ? ': ' + preview.description : ''))}" />`;
              warnings.add('OLE 개체는 파일에 저장된 미리보기 그림으로 가져옵니다. 원본 OLE 데이터는 원본 한글 파일에 남아 있으며 편집본에서는 개체 내부를 편집할 수 없습니다.');
            }
          }
        }
        if (c.ctrlId === 'tbl') {
          const ordinal = (byParagraph.get(globalPara) ?? []).filter(c => c.ctrlId === 'tbl').indexOf(c);
          controlHtml = restoreCellContent(doc, section, para, c.controlIndex, ordinal, controlHtml, equationHtml, (path, index) => {
            try {
              const props = JSON.parse(doc.getCellShapePropertiesByPath(section, para, JSON.stringify(path), index));
              return objectHtml(cellShapePosition(section, para, path, index, props), props, props.description || '도형', true);
            } catch { return; }
          }, warnings, equationMetadata, restoredHeaders.has(`${section}:${para}`), (path, index, html) => {
            try {
              const props = JSON.parse(doc.getCellPicturePropertiesByPath(section, para, JSON.stringify(path), index));
              let picture = html.replace(/\s(?:alt|data-image-floating)="[^"]*"/g, '')
                .replace(/\s*\/?>$/, ` alt="${escape(props.description || '')}" />`);
              if (!props.treatAsChar) {
                const position = controlPosition(section, para, index, path);
                if (position) picture = picture.replace(/<img\b/, `<img data-image-floating="${escape(JSON.stringify(floatingGeometry(position, props)))}"`);
                else warnings.add('셀 안의 떠 있는 그림 위치를 확인하지 못해 일반 그림으로 가져왔습니다. 원본 배치를 확인하세요.');
              }
              return croppedPicture(picture, props);
            }
            catch { return html; }
          }, noteReference, (path, html, policy) => nestedTableGeometry(doc, section, para, path, html, pitches, policy),
            (key, html) => captionKeys.has(key) ? html.replace(/^<table\b/, `<table data-hwp-caption-key="${key}"`) : html,
            metadata?.fontSpace, metadata?.latinFaces);
          const hasNextPageBreak = !hasColumns && paragraphControls.length === 1 &&
            !doc.getTextRange(section, para, 0, 0xffffffff).trim() &&
            metadata?.paragraphs[`${section}:${para + 1}`]?.pageBreak === true;
          controlHtml = tableGeometry(doc, section, para, c.controlIndex, controlHtml, pitches, splitCellLines, hasNextPageBreak,
            (cell, p) => hwpTabAttribute(equationMetadata().get(`${section}:${para}/t${ordinal}/c${cell}/p${p}`)?.tabStops), metadata?.hwpxCellFontSpace, metadata?.hwpxCellRasterFills);
          if (relatedTableLayout && c === paragraphControls[0])
            controlHtml = controlHtml.replace(/^<table\b/, `<table ${hwpRelatedTableLayoutAttribute}="${escape(relatedTableLayout)}"`);
          const background = tableBackgrounds.get(c.controlIndex);
          if (background) {
            const target = /(<t[dh]\b[^>]*>\s*<p\b[^>]*>)/;
            if (target.test(controlHtml)) {
              controlHtml = controlHtml.replace(target, '$1' + background.html);
              warnings.add('겹친 표와 배경 그림은 편집 가능한 표와 셀 배경 그림으로 함께 가져옵니다.');
            } else {
              controlHtml = `<p>${background.original}</p>` + controlHtml;
              warnings.add('겹친 표의 배경 그림을 별도 그림으로 보존합니다. 원본 배치를 확인하세요.');
            }
          }
        }
        if (c.ctrlId === 'eqed') {
          const equation = JSON.parse(doc.getEquationProperties(section, para, c.controlIndex, -1, -1));
          const script = equation.script || '';
          try {
            controlHtml = equationHtml({ ...equation, script });
          } catch { controlHtml = ''; }
        }
        if (!controlHtml || /^<!--/.test(controlHtml)) {
          let detail = '';
          if (c.ctrlId === 'eqed') detail = JSON.parse(doc.getEquationProperties(section, para, c.controlIndex, -1, -1)).script ?? '';
          controlHtml = `<span>[${escape(c.userDesc || c.ctrlId)}${detail ? ': ' + escape(detail) : ' — 원본 보기에서 확인'}]</span>`;
          if (c.ctrlId === 'gso') {
            try {
              const count = doc.getCellParagraphCount(section, para, c.controlIndex, 0);
              if (count) controlHtml = fragment(doc.exportSelectionInCellHtml(section, para, c.controlIndex, 0, 0, 0, count - 1, 0xffffffff));
            } catch { /* Non-textbox shapes retain a visible placeholder. */ }
          }
          warnings.add(`${c.userDesc || c.ctrlId} 개체는 편집 가능한 원본 개체로 변환하지 못해 위치에 텍스트 표시를 남깁니다.`);
        }
        if (/^<(?:table|p)\b/.test(controlHtml)) {
          const inlineProperties = c.ctrlId === 'tbl' ? JSON.parse(doc.getTableProperties(section, para, c.controlIndex)) : undefined;
          const inlineTable = !!inlineProperties?.treatAsChar;
          if (inlineTable && inlineProperties.pageBreak === 1 && !hasColumns && !startsNewFlow && paragraphControls.length === 1) {
            const previous = byParagraph.get(globalPara - 1), earlier = byParagraph.get(globalPara - 2);
            const blank = metadata?.paragraphs[`${section}:${para - 1}`];
            const distance = previous?.length === 1 && previous[0].ctrlId === 'tbl' ? 1
              : !previous?.length && para >= 2 && earlier?.length === 1 && earlier[0].ctrlId === 'tbl' &&
                !blank?.pageBreak && !blank?.columnBreak ? 2 : undefined;
            const nativePage = distance && sourceKeptInlineTable(doc, section, para, c.controlIndex, distance, pitches);
            if (nativePage) controlHtml = controlHtml.replace(/^<table\b/, '<table data-hwp-kept-table-page="' + nativePage + '"');
          }
          const precedingBlank = metadata?.paragraphs[`${section}:${para - 1}`];
          if (inlineTable && inlineProperties.pageBreak === 2 && !hasColumns && !startsNewFlow && paragraphControls.length === 1 &&
              !precedingBlank?.pageBreak && !precedingBlank?.columnBreak) {
            const nativePage = sourceKeptInlineTableAfterBody(doc, section, para, c.controlIndex);
            if (nativePage) controlHtml = controlHtml.replace(/^<table\b/, '<table data-hwp-kept-table-page="' + nativePage + '"');
          }
          if (inlineTable && !doc.getTextRange(section, para, 0, 0xffffffff).trim()) {
            const lineIndex = JSON.parse(doc.getLineInfo(section, para, pos)).lineIndex;
            const char = JSON.parse(doc.getCharPropertiesAt(section, para, pos));
            const leading = structure?.lineSpacings?.[lineIndex] ??
              (['Percent', 'SpaceOnly'].includes(properties.lineSpacingType) ? physicalLineHeight(properties, char) - Number(char.fontSize) / 75 : 0);
            // An inline table occupies the line's object height, followed by
            // the source line's extra leading. Removing the placeholder text
            // line must not also remove that independent gap.
            if (Number.isFinite(leading) && leading > 0) {
              const props = JSON.parse(doc.getTableProperties(section, para, c.controlIndex));
              controlHtml = controlHtml.replace(/^<table\b[^>]*>/, tag => cssStyle(tag, `margin-bottom:${Math.round((px(Math.max(0, Number(props.outerBottom) || 0)) + leading) * 100) / 100}px;`));
            }
          }
          // Spaces before an inline table share the table's line. Keep their
          // text, but do not turn them into an additional full-height line.
          // A positive subpixel pitch also survives DOCX exact-line spacing.
          if (content.trim()) {
            const spacesOnObjectLine = inlineTable && paragraphControls.length === 1 && onlyWhitespace(content) &&
              /^ +$/.test(entities(content.replace(/<[^>]*>/g, ''))) &&
              JSON.parse(doc.getLineInfo(section, para, pos)).lineCount === 1;
            const host = spacesOnObjectLine ? cssStyle(open, 'line-height:0.1pt;margin-bottom:0px;')
              .replace(/>$/, ' data-hwp-inline-table-spaces="true">')
              : inlineTable && onlyWhitespace(content) ? cssStyle(open, 'line-height:0.1px;margin-bottom:0px;')
              : onlyFloating(content) ? cssStyle(open, 'line-height:1px;margin-top:0;margin-bottom:0;') : open;
            append(host + content + '</p>');
          }
          const tableWidth = Number(controlHtml.match(/^<table[^>]*style="[^"]*width:([\d.]+)px/)?.[1]);
          const pd: PageDef = JSON.parse(doc.getPageDef(section));
          if (c.ctrlId === 'tbl' && !inlineTable) {
            const props = JSON.parse(doc.getTableProperties(section, para, c.controlIndex));
            if (props.textWrap === 'InFrontOfText' && ['Page', 'Paper'].includes(props.vertRelTo) && props.vertAlign === 'Bottom' && !props.hasCaption) {
              const position = controlPosition(section, para, c.controlIndex);
              const pages = new Set(JSON.parse(doc.getTableCellBboxes(section, para, c.controlIndex)).map((b: any) => b.pageIndex));
              const height = px(pd.landscape ? pd.width : pd.height);
              const edge = height - (props.vertRelTo === 'Page' ? px(pd.marginBottom + pd.marginFooter) : 0);
              const gap = position && edge - position.y - position.h;
              if (position && pages.size === 1 && position.x >= 0 && Number.isFinite(gap) && gap >= -.5 && gap <= 1000)
                controlHtml = controlHtml.replace(/^<table\b/, `<table data-table-anchor="${escape(JSON.stringify({ x: position.x, y: Math.max(0, gap), bottom: 0,
                  overlay: props.vertRelTo === 'Page' ? 'page-bottom' : 'paper-bottom' }))}"`);
            }
          }
          if (inlineTable && columnCount === 1) {
            const position = controlPosition(section, para, c.controlIndex);
            if (position) controlHtml = controlHtml.replace(/^<table\b[^>]*>/, tag => cssStyle(tag, `margin-left:${Math.round((position.x - px(pd.marginLeft)) * 100) / 100}px;`));
          } else if (inlineTable && columnCount === 2 && paragraphControls.length === 1 && properties.marginLeft > 0 &&
              !properties.marginRight && !content.trim() && !doc.getTextRange(section, para, 0, 0xffffffff) &&
              inlineProperties.horzRelTo === 'Para' && inlineProperties.horzAlign === 'Left' &&
              inlineProperties.textWrap === 'TopAndBottom' && !inlineProperties.horzOffset && !inlineProperties.outerLeft &&
              !inlineProperties.hasCaption && inlineProperties.restrictInPage && !inlineProperties.allowOverlap) {
            const dimensions = JSON.parse(doc.getTableDimensions(section, para, c.controlIndex));
            if (dimensions.cellCount === 1 && dimensions.rowCount === 1 && dimensions.colCount === 1) {
              const position = controlPosition(section, para, c.controlIndex);
              const indent = position && sourceInlineColumnTableIndent(doc, section, para, c.controlIndex, position, properties.marginLeft);
              if (indent) controlHtml = controlHtml.replace(/^<table\b[^>]*>/, tag => cssStyle(tag, `margin-left:${indent}px;`));
            }
          }
          const fullWidth = columnCount > 1 && tableWidth > px((pd.landscape ? pd.height : pd.width) - pd.marginLeft - pd.marginRight) / columnCount * 1.35;
          if (para === 0 && /^<table\b/.test(controlHtml) && tableWidth >= px((pd.landscape ? pd.height : pd.width) - pd.marginLeft - pd.marginRight) * .8) {
            const props = JSON.parse(doc.getTableProperties(section, para, c.controlIndex)), position = controlPosition(section, para, c.controlIndex);
            const pages = new Set(JSON.parse(doc.getTableCellBboxes(section, para, c.controlIndex)).map((b: any) => b.pageIndex));
            if (position && pages.size === 1) {
              const anchor = !props.treatAsChar && props.textWrap === 'TopAndBottom' && props.vertRelTo === 'Paper' && props.horzRelTo === 'Paper'
                ? { x: position.x, y: position.y, bottom: px(props.outerBottom || 0) }
                : startsWithControl && onlyWhitespace(content) ? hwpMarginCenteredTableAnchor(doc, props, position) : undefined;
              if (anchor) controlHtml = controlHtml.replace(/^<table\b/, `<table data-table-anchor="${escape(JSON.stringify(anchor))}"`);
            }
          }
          // A standalone empty host paragraph is the table's live flow
          // anchor, not an extra text line. Preserve its nonnegative offset
          // instead of pinning to cached page y. The native renderer applies
          // outerTop only for some relocated hosts; retain that distance only
          // when the source line and actual table independently confirm it.
          // Visible host text, negative offsets, captions and columns need
          // their own exclusion layout and retain the explicit fallback.
          if (!inlineTable && !hasColumns && /^<table\b/.test(controlHtml) && !/^<table\b[^>]*\bdata-table-anchor=/.test(controlHtml) &&
            paragraphControls.length === 1 && !doc.getTextRange(section, para, 0, 0xffffffff).trim() &&
            tableWidth >= px((pd.landscape ? pd.height : pd.width) - pd.marginLeft - pd.marginRight) * .8) {
            const props = JSON.parse(doc.getTableProperties(section, para, c.controlIndex)), position = controlPosition(section, para, c.controlIndex);
            const offset = px(Number(props.vertOffset)), outerTop = px(Number(props.outerTop));
            const firstLine = JSON.parse(doc.getParagraphLineMetrics(section, para, '[]')).lines?.[0];
            const sourceDistance = position && firstLine ? position.y - px(pd.marginTop + pd.marginHeader) - Number(firstLine.verticalPosition) - offset : NaN;
            const y = offset + (outerTop > 0 && Math.abs(sourceDistance - outerTop) <= .5 ? outerTop : 0);
            const x = position && (props.horzRelTo === 'Paper' ? position.x : position.x - px(pd.marginLeft));
            if (props.textWrap === 'TopAndBottom' && props.vertRelTo === 'Para' && props.vertAlign === 'Top' && !props.hasCaption &&
              ['Column', 'Para', 'Paper'].includes(props.horzRelTo) && Number(props.vertOffset) >= 0 && y >= 0 && x !== undefined && x >= 0)
              controlHtml = controlHtml.replace(/^<table\b/, `<table data-table-anchor="${escape(JSON.stringify({ x, y, bottom: px(props.outerBottom || 0), reference: 'paragraph', ...(props.horzRelTo !== 'Paper' ? { horizontal: 'margin' } : {}) }))}"`);
          }
          if (c.ctrlId === 'tbl' && !inlineTable && !/^<table\b[^>]*\bdata-table-anchor=/.test(controlHtml))
            warnings.add('일부 떠 있는 표는 본문 흐름 안으로 옮깁니다. 음수 위치·호스트 본문·단·겹침·캡션은 원본과 다를 수 있습니다.');
          if (section === 0 && para === 0 && inlineTable && !hasColumns && !content.trim() &&
            paragraphControls.length > 1 && paragraphControls.at(-1) === c &&
            paragraphControls.slice(0, -1).every(control => ['secd', 'cold', 'pgnp'].includes(control.ctrlId) && control.pos < c.pos) &&
            !doc.getTextRange(section, para, 0, 0xffffffff) && !properties.pageBreakBefore && !structure?.pageBreak && !structure?.columnBreak &&
            !properties.spacingBefore && !properties.spacingAfter && properties.lineSpacingType === 'Percent' && !properties.fontLineHeight) {
            const props = JSON.parse(doc.getTableProperties(section, para, c.controlIndex));
            const position = controlPosition(section, para, c.controlIndex);
            const boxes = JSON.parse(doc.getTableCellBboxes(section, para, c.controlIndex));
            const lead = !props.hasCaption && Number(props.vertOffset) === 0 && position?.page === 0 && boxes.length &&
              boxes.every((box: any) => box.pageIndex === 0) && sourceInitialTableLead(doc,
                { index: c.controlIndex, textStart: c.pos, y: position.y, height: position.h }, font, percent, (pd.marginTop + pd.marginHeader) / 75);
            // The object remains inline. An ordinary exact-height paragraph
            // represents the source control line, with the table's outer top
            // reserved as following space in both Studio and standard DOCX.
            if (lead) append(`<p style="font-size:${font}px;line-height:${lead.advance}px;margin:0 0 ${lead.outerTop}px 0" data-word-line-rule="exact"></p>`);
          }
          if (fullWidth) { closeColumns(); append('<div data-column-count="1" data-column-space="0">'); }
          const tablePiece = pieces.length;
          if (c.ctrlId === 'tbl' && !inlineTable && !hasColumns && startsWithControl && !pos && c.pos === 0 &&
            paragraphControls.length === 1 && !content && textLength &&
            tableWidth >= px((pd.landscape ? pd.height : pd.width) - pd.marginLeft - pd.marginRight) * .8) {
            if (!/^<table\b[^>]*\bdata-table-anchor=/.test(controlHtml))
              hostOnPriorPage = sourceHostOnPriorPage(doc, section, para, c.controlIndex, pitches);
            if (hostOnPriorPage || sourceHostBeforeFloatingTable(doc, section, para, c.controlIndex, pitches)) tableFollowingHost = tablePiece;
          }
          if (inlineTable && !hasColumns && !inlineProperties.hasCaption &&
              paragraphControls.every(control => control === c || ['secd','cold','pgnp','head','foot'].includes(control.ctrlId)) &&
              !doc.getTextRange(section, para, 0, 0xffffffff).trim() && !/^<table\b[^>]*\bdata-table-anchor=/.test(controlHtml)) {
            const top = hwpInlineTableTop(inlineProperties.outerTop);
            if (top !== undefined)
              controlHtml = controlHtml.replace(/^<table\b/, `<table ${hwpInlineTableTopAttribute}="${top}"`);
          }
          append(controlHtml);
          inlineTail = undefined;
          if (inlineTable && !fullWidth && columnCount === 1 && tableWidth > 0) {
            const position = controlPosition(section, para, c.controlIndex);
            const pages = new Set(JSON.parse(doc.getTableCellBboxes(section, para, c.controlIndex)).map((b: any) => b.pageIndex));
            const left = position ? position.x - px(pd.marginLeft) : 0;
            const available = px((pd.landscape ? pd.height : pd.width) - pd.marginLeft - pd.marginRight) - left;
            if (pages.size === 1 && available - tableWidth >= 1) {
              const props = JSON.parse(doc.getTableProperties(section, para, c.controlIndex));
              inlineTail = { piece: tablePiece, line: JSON.parse(doc.getLineInfo(section, para, pos)).lineIndex, width: tableWidth, left, available,
                wrapping: pos === 0 && paragraphControls.length === 1 && !props.hasCaption && !Number(properties.spacingBefore) && left >= 0 && tableWidth <= available * .65,
                right: px(Math.max(0, Number(props.outerRight) || 0)), bottom: px(Math.max(0, Number(props.outerBottom) || 0)) };
            }
          }
          if (fullWidth) { append('</div>'); openColumns(); }
          content = ''; contentBytes = 0;
          lastInlineTable = inlineTable;
          pictureTableSpacesShareSourceLine = inlineTable && paragraphControls.length === 1 &&
            controlHtml.includes('data-word-text-alignment="top"') &&
            /^ +$/.test(doc.getTextRange(section, para, 0, 0xffffffff)) &&
            JSON.parse(doc.getLineInfo(section, para, 0)).lineCount === 1;
        } else {
          appendInline(controlHtml.replace(/<img /g, '<img data-inline-image="true" '));
          lastInlineTable = false;
          pictureTableSpacesShareSourceLine = false;
        }
      }
      appendInline(inline(range(cursor, 0xffffffff)));
      if (inlineTail && !onlyWhitespace(content) && !onlyFloating(content) &&
        !/<(?:br|table)\b/.test(content) && JSON.parse(doc.getLineInfo(section, para, cursor)).lineIndex === inlineTail.line &&
        JSON.parse(doc.getLineInfo(section, para, 0)).lineCount === inlineTail.line + 1) {
        // Inline tables and following text can share the same source line.
        // A borderless layout row preserves this editable relationship in
        // standard DOCX too, instead of inventing a paragraph below the table.
        // Restrict this to a source table on one page and its final text line.
        const { piece, width, left, available } = inlineTail;
        const table = pieces[piece];
        const gap = [...(table.match(/^<table\b[^>]*>/)?.[0] ?? '').matchAll(/margin-bottom:([\d.]+)px/g)].at(-1)?.[1] ?? '0';
        const inner = table.replace(/^<table\b[^>]*>/, tag => cssStyle(tag, 'margin-left:0;margin-right:0;margin-bottom:0;'));
        const cell = 'padding:0;border:none;vertical-align:bottom;';
        const tailWidth = Math.max(1, available - width);
        const tail = cssStyle(open, 'margin:0;text-indent:0;text-align:left;') + content + '</p>';
        const wrapped = `<table style="width:${available}px;margin-left:${left}px;margin-bottom:${gap}px;border:none;"><tr><td colwidth="${Math.round(width)}" style="${cell}width:${width}px;">${inner}</td><td colwidth="${Math.round(tailWidth)}" style="${cell}width:${tailWidth}px;">${tail}</td></tr></table>`;
        size += hwpHtmlBytes(wrapped) - hwpHtmlBytes(table); checkHwpHtmlSize(size);
        pieces[piece] = wrapped; content = ''; contentBytes = 0;
      } else if (inlineTail?.wrapping && inlineTail.line === 0 && !onlyWhitespace(content) &&
        !/<(?:br|table|img|svg)\b|data-(?:note-reference|page-number|symbol-font)/.test(content) &&
        JSON.parse(doc.getLineInfo(section, para, 0)).lineCount > 1 && properties.lineSpacingType === 'Percent' && !properties.fontLineHeight) {
        // A leading inline label is an object on the first line, not a block
        // followed by a second paragraph. A paragraph-relative wrapping table
        // lets later text regain the full line width, including after edits.
        const { piece, left, right, bottom } = inlineTail;
        const anchor = { x: left, y: 0, bottom, reference: 'paragraph', horizontal: 'margin', wrap: 'right', right };
        const previous = pieces[piece];
        pieces[piece] = previous.replace(/^<table\b/, `<table data-table-anchor="${escape(JSON.stringify(anchor))}"`);
        size += hwpHtmlBytes(pieces[piece]) - hwpHtmlBytes(previous); checkHwpHtmlSize(size);
        // Cached first-line pitch includes the removed object's height. The
        // remaining paragraph must retain its text's own percentage leading.
        const sizes = [...content.matchAll(/font-size:\s*([\d.]+)(pt|px)/gi)].map(m => Number(m[1]) * (m[2].toLowerCase() === 'pt' ? 4 / 3 : 1));
        const font = Math.max(Number(paragraphChar.fontSize) / 75, ...sizes), base = font * Number(properties.lineSpacing) / 100;
        if (font > 0 && base > 0) open = cssStyle(open, `text-indent:0;line-height:${base}px;--hwp-body-line:${encodeHwpBodyLine({ percent: Number(properties.lineSpacing), font, base })};`);
      }
      if (!content && !textLength && !paragraphControls.length && columnCount === 1 && !startsNewFlow &&
          properties.lineSpacingType === 'Percent' && !properties.fontLineHeight &&
          !Number(properties.spacingBefore) && !Number(properties.spacingAfter)) {
        const line = sourceTerminalEmptyLine(doc, section, para, pitches, font, base, percent) ??
          (metadata?.paragraphs[String(section) + ':' + (para + 1)]?.pageBreak
            ? sourceTerminalEmptyLineAfterTable(doc, section, para, pitches, font, base, percent) : null);
        if (line) open = cssStyle(open, `line-height:${line.base}px;--hwp-body-line:${encodeHwpBodyLine(line)};`);
      }
      // Keep real blank paragraphs, but avoid an extra empty paragraph after a table.
      if (content.trim() || !replacedTextBox && !paragraphControls.some(c => c.ctrlId === 'tbl')) {
        if (lastInlineTable && onlyWhitespace(content)) {
          open = pictureTableSpacesShareSourceLine && /^ +$/.test(entities(content.replace(/<[^>]*>/g, '')))
            ? cssStyle(open, 'line-height:0.1pt;margin-top:0px;margin-bottom:0px;')
              .replace(/>$/, ' data-hwp-inline-table-spaces="true">')
            : cssStyle(open, 'line-height:0.1px;margin-top:0px;');
        }
        else if (onlyFloating(content))
          open = cssStyle(open, 'line-height:1px;margin-top:0;margin-bottom:0;');
        if (para === 0 && columnCount === 2 && properties.lineSpacingType === 'Percent' && !properties.fontLineHeight &&
            !properties.spacingBefore && !properties.spacingAfter && !/<(?:br|svg|table|sup|sub)\b|\bdata-(?:docx-tab|note-reference|document-field)/i.test(content)) {
          const art = content.match(/<img\b[^>]*>/gi) ?? [];
          const actualPage = pitches.get(`page:${section}:${para}`);
          if (actualPage !== undefined && art.length && art.every(image => image.includes('hwpPaperDrawing'))) {
            const text = entities(content.replace(/<[^>]*>/g, ''));
            const line = sourcePaperHeadingLine(doc, section, para, actualPage, text, percent, font, base);
            if (line) open = cssStyle(open, `--hwp-body-line:${encodeHwpBodyLine(line)};`);
          }
        }
        const imageControl = paragraphControls.find(c => c.ctrlId === 'gso');
        if (imageControl && paragraphControls.every(c => c === imageControl || c.ctrlId === 'pghd') &&
          (content.match(/<img\b/g) ?? []).length === 1 && !content.includes('data-image-floating=') &&
          (content.includes('data-hwp-object=') || columnCount === 2) &&
          !doc.getTextRange(section, para, 0, 0xffffffff).trim()) {
          // The native object-only line includes the picture's descent inside
          // its measured height. Baseline-aligned CSS images add the font's
          // descent again. Keep the native top alignment as standard OOXML.
          const metrics = JSON.parse(doc.getParagraphLineMetrics(section, para, '[]'));
          const position = controlPosition(section, para, imageControl.controlIndex);
          if (position && (content.includes('data-hwp-object=')
            ? metrics.lines?.length === 1 && Math.abs(Number(metrics.lines[0].height) - position.h) <= 1
            : sourceTightInlinePicture(doc, section, para, imageControl.controlIndex, position)))
            open = open.replace(/>$/, ' data-word-text-alignment="top">');
        }
        const before = tableFollowingHost !== undefined && !onlyWhitespace(content) &&
          !/<(?:br|img|table|svg)\b|data-(?:note-reference|page-number|symbol-font)/i.test(content) && !content.includes('\t')
          ? tableFollowingHost : undefined;
        if (before !== undefined && hostOnPriorPage) {
          const { x, top, bottom, blankUsesTableGap } = hostOnPriorPage;
          const anchor = { x, y: top, bottom, reference: 'paragraph', horizontal: 'margin' };
          const original = pieces[before];
          pieces[before] = original.replace(/^<table\b/, `<table data-table-anchor="${escape(JSON.stringify(anchor))}"`);
          size += hwpHtmlBytes(pieces[before]) - hwpHtmlBytes(original); checkHwpHtmlSize(size);
          // Retain the real blank line. Its native top already includes the
          // table's outer gap, and the following text proves the full pitch.
          if (blankUsesTableGap) blankAfterFloatingTable = para + 1;
        }
        append(open + content + '</p>', before);
      }
      onProgress?.({ label: '본문·표·그림을 변환하는 중…', completed: globalPara + 1, total, unit: '문단' });
    }
    closeColumns();
  }
  for (const note of noteStories.notes) {
    const layout = note.layout ? ` ${hwpFootnoteLayoutAttribute}="${escape(note.layout)}"` : '';
    append(`<div data-document-note="${note.id}" data-note-kind="${note.kind}" data-note-label="${escape(note.label)}"${layout}>${note.html}</div>`);
  }
  onProgress?.({ label: '변환한 본문과 그림을 정리하는 중…' });
  let html = pieces.join('\n').replace(/\sdata-hwp-cell-paragraph="\d+"/g, '').replace(/<!--\s*rhwp:\s*셀 안 (ColumnDef|Bookmark|PageHide)\b[\s\S]*?-->/g, (_, type: string) => {
    warnings.add(type === 'ColumnDef' ? '표 안의 단 설정은 단일 흐름으로 가져옵니다. 원본의 셀 안 배치를 확인하세요.'
      : type === 'Bookmark' ? '책갈피 위치 정보는 편집본에 전달되지 않습니다. 본문 내용은 유지합니다.'
      : '원본의 쪽 숨김 설정은 편집본에 전달되지 않습니다. 본문 내용은 유지합니다.');
    return '';
  }).replace(/<div data-column-count="[^"]*"[^>]*>\s*<\/div>/g, '').replace(/<!--\s*rhwp:([\s\S]*?)-->/g, () => {
    warnings.add('표 안에 변환하지 못한 개체가 있습니다. 해당 위치의 표시와 원본을 확인하세요.');
    return '<span>[변환하지 못한 개체 — 원본 보기에서 확인]</span>';
  });
  // Verified Hancom glyph meanings, also used by the core reader's text paint.
  // Limit this to the two forms symbols; never guess arbitrary PUA meanings.
  // This Hancom line glyph advances half an em. A plain Unicode substitution
  // advances a full em in the fallback face and doubles the cut-line's width.
  html = html.replace(/<span style="([^"]*)">([^<]*\u{F081C}[^<]*)<\/span>/gu, (tag, css: string, text: string) => {
    const font = [...css.matchAll(/(?:^|;)font-size:([\d.]+)(pt|px)/g)].at(-1);
    if (!font) return tag;
    const size = Number(font[1]) * (font[2] === 'pt' ? 4 / 3 : 1);
    const content = text.split(/(\u{F081C}+)/gu).map(part => part.includes('\u{F081C}')
      ? `<span style="${css};letter-spacing:${-size / 2}px;">${part.replace(/\u{F081C}/gu, '┈')}</span>` : part).join('');
    return `<span style="${css}">${content}</span>`;
  }).replace(/\u{F081C}/gu, '┈').replace(/\u{F012B}/gu, '(인)');
  // Some HWPX manifests label embedded BMPs as application/octet-stream.
  // Sniff only known raster signatures; never permit arbitrary data resources.
  html = html.replace(/src="data:application\/octet-stream;base64,([A-Za-z0-9+/=\s]+)"/g, (source, base64: string) => {
    const b = base64.replace(/\s/g, '');
    const mime = b.startsWith('Qk') ? 'bmp' : b.startsWith('iVBORw0KGgo') ? 'png' : b.startsWith('/9j/') ? 'jpeg' : b.startsWith('R0lGOD') ? 'gif' : '';
    return mime ? `src="data:image/${mime};base64,${b}"` : source;
  });
  // A paragraph's CSS strut otherwise inherits Word's 11pt font even when every
  // HWP run is 9pt, inflating line-height percentages throughout long tables.
  html = html.replace(/<p\b([^>]*?)style="([^"]*)"([^>]*)>(\s*<span style="([^"]*)">)/g, (_, before: string, css: string, after: string, span: string, runCss: string) => {
    const font = ['font-size', 'font-family'].map(key => runCss.match(new RegExp(`(?:^|;)${key}:([^;]+)`))?.[0]).filter(Boolean).join(';');
    return `<p${before}style="${css}${font};"${after}>${span}`;
  });
  // HWP percentage leading is based on the character em. Word's AUTO leading
  // multiplies a font's full line box instead, so preserve the physical pitch
  // for these imported paragraphs. Ordinary authored HTML keeps its own rule.
  html = html.replace(/<p\b[^>]*>/g, tag => {
    // Tabs and other source attributes must not bypass line normalization.
    // CSS percentages use the em; DOCX AUTO uses the larger font line box.
    const css = tag.match(/\sstyle="([^"]*)"/)?.[1] ?? '';
    const line = [...css.matchAll(/(?:^|;)line-height:([\d.]+)%;/g)].at(-1);
    const font = [...css.matchAll(/(?:^|;)font-size:([\d.]+)(pt|px)(?=;|$)/g)].at(-1);
    // Do not consume the semicolon: adjacent declarations can override one
    // another, and skipping the second would replace measured exact leading
    // with an earlier percentage again.
    if (!line || !font || [...css.matchAll(/(?:^|;)line-height:([^;]+)(?=;|$)/g)].at(-1)?.[1] !== line[1] + '%') return tag;
    const pitch = Number(font[1]) * (font[2] === 'pt' ? 4 / 3 : 1) * Number(line[1]) / 100;
    return cssStyle(tag, `line-height:${Math.max(.1, Math.round(pitch * 100) / 100)}px;`);
  });
  if (captionDefinitions.length) {
    const rendered = renderTableCaptions(doc, captionDefinitions, isolated => editableHwp(isolated, undefined, undefined, noteStory, captionDepth + 1));
    rendered.warnings.forEach(w => warnings.add(w));
    html = restoreTableCaptions(html, rendered.captions, warnings);
  }
  // Script-relative glyph sizes do not replace the base em in the preceding
  // paragraph and cell leading calculations.
  html = html.replace(/<span(?: style="([^"]*)")? data-hwp-script-size="([\d.]+pt)">/g,
    (_, style: string | undefined, size: string) => `<span style="${style ?? ''}font-size:${size};">`);
  checkHwpHtmlSize(hwpHtmlBytes(html));
  warnings.add('복잡한 표 테두리·셀 크기, 글상자·수식·각주·필드 등 고급 서식은 완전히 호환되지 않습니다. 편집본과 원본을 비교하세요.');
  return { html: html || '<p></p>', page: firstPage, warnings: [...warnings] };
}

export function setHwpPage(doc: HwpDocument, page: PageSettings, section = 0) {
  doc.setSectionDef(section, JSON.stringify({ hideEmptyLine: page.hideEmptyLines === true }));
  const m = page.margins ?? { top: page.margin, right: page.margin, bottom: page.margin, left: page.margin };
  const letter = page.size === 'Letter';
  const header = Math.min(m.top, page.headerDistance ?? Math.max(0, m.top / 2 - 3));
  const footer = Math.min(m.bottom, page.footerDistance ?? Math.max(0, m.bottom / 2 - 3));
  doc.setPageDef(section, JSON.stringify({ width: page.size === 'Custom' ? hu(page.width!) : letter ? 61200 : 59528, height: page.size === 'Custom' ? hu(page.height!) : letter ? 79200 : 84188, landscape: page.orientation === 'landscape',
    // HWP top/bottom margins locate the furniture at the paper edge; the
    // header/footer areas are the additional distance to the body boundary.
    marginTop: hu(header), marginRight: activeHwpHorizontalMargin(page, 'right') ?? hu(m.right),
    marginBottom: hu(footer), marginLeft: activeHwpHorizontalMargin(page, 'left') ?? hu(m.left),
    marginHeader: hu(m.top - header), marginFooter: hu(m.bottom - footer), marginGutter: hu(page.gutter ?? 0) }));
  if (page.columns) doc.setColumnDef(section, page.columns.count, page.columns.balance ? 1 : 0, 1, Math.round(page.columns.space * 75));
  for (const isHeader of [true, false]) {
    const text = page[isHeader ? 'header' : 'footer'];
    if (text) { doc.createHeaderFooter(section, isHeader, 0); doc.insertTextInHeaderFooter(section, isHeader, 0, 0, 0, text); }
  }
}

// The pinned native paste API returns success after converting oversized
// HTML to plain text. Keep each call within its structured UTF-8 budget.
const nativeStructuredPasteBytes = 2_000_000;
function pasteStructuredHwpHtml(doc: HwpDocument, html: string, pageBreak: string) {
  const before = '<p></p>';
  const fail = () => new Error('문서 내용을 구조를 유지한 한글 형식으로 변환하지 못했습니다.');
  if (hwpHtmlBytes(html) + before.length * 2 <= nativeStructuredPasteBytes) {
    if (!JSON.parse(doc.pasteHtml(0, 0, 0, before + html + before)).ok) throw fail();
    return;
  }

  // Prefer the existing printed page boundaries. An oversized page can be
  // split between complete sibling blocks without inserting a document break.
  const tail = 'WORDSTUDIO-PASTE-END-' + crypto.randomUUID();
  const after = '<p>' + tail + '</p>';
  const budget = nativeStructuredPasteBytes - before.length - after.length;
  let index = 0;
  const append = (body: string) => {
    const start = doc.getParagraphCount(0) - 1;
    if (start < 0 || doc.getParagraphLength(0, start) !== 0 ||
        !JSON.parse(doc.pasteHtml(0, start, 0, before + body + after)).ok) throw fail();
    const end = doc.getParagraphCount(0) - 1;
    if (end <= start || doc.getTextRange(0, end, 0, tail.length + 1) !== tail) throw fail();
    doc.deleteText(0, end, 0, tail.length);
    if (index++) {
      if (doc.getParagraphLength(0, start) !== 0) throw fail();
      doc.deleteParagraph(0, start);
    }
  };

  const batch: string[] = [];
  let size = 0;
  const appendPart = (part: string, bytes: number) => {
    if (bytes > budget) throw new Error('한 문단이나 개체의 변환 데이터가 한글 엔진의 문서 구조 보존 한도를 넘습니다.');
    if (size + bytes > budget) {
      append(batch.join(''));
      batch.length = 0;
      size = 0;
    }
    batch.push(part);
    size += bytes;
  };
  for (const [i, page] of html.split(pageBreak).entries()) {
    const part = i ? pageBreak + page : page;
    const bytes = hwpHtmlBytes(part);
    if (bytes <= budget) appendPart(part, bytes);
    else {
      const blocks = topLevelHwpPasteBlocks(part);
      if (!blocks) throw fail();
      for (const block of blocks) appendPart(block, hwpHtmlBytes(block));
    }
  }
  if (size) append(batch.join(''));
}

function pasteHwpBody(doc: HwpDocument, converted: string, breakToken: string, onProgress?: HwpProgressReporter) {
  checkHwpHtmlSize(hwpHtmlBytes(converted));
  // The first pasted paragraph inherits the host's paragraph shape even on
  // the multi-paragraph path. Sacrifice a leading blank, then remove both
  // sentinels. The writer rebuilds section controls from the section model.
  onProgress?.({ label: '본문·표·그림을 한글 문서로 변환하는 중…' });
  pasteStructuredHwpHtml(doc, converted, `<p>${breakToken}</p>`);
  onProgress?.({ label: '한글 문서의 쪽 나누기를 반영하는 중…' });
  // These edits address the parsed paragraph model. Intermediate pages
  // cannot be shown until the complete set of authored boundaries is in place.
  doc.beginBatch();
  try {
    const last = doc.getParagraphCount(0) - 1;
    if (last > 0 && doc.getParagraphLength(0, last) === 0) doc.deleteParagraph(0, last);
    if (doc.getParagraphCount(0) > 1) doc.deleteParagraph(0, 0);
    for (let p = 0; p < doc.getParagraphCount(0); p++) {
      if (doc.getTextRange(0, p, 0, 100) !== breakToken) continue;
      if (p + 1 < doc.getParagraphCount(0) && doc.getTextRange(0, p + 1, 0, 100) !== breakToken) doc.deleteParagraph(0, p);
      else doc.deleteText(0, p, 0, breakToken.length);
      doc.applyParaFormat(0, p, JSON.stringify({ pageBreakBefore: true }));
    }
  } finally { doc.endBatch(); }
}

export function serializeHwp(Core: typeof HwpDocument, html: string, page: PageSettings, format: 'hwp' | 'hwpx',
  onProgress?: HwpProgressReporter, textBoxes: readonly HwpExportTextBox[] = [], nativePrint?: NativePrintPlan
): { bytes: Uint8Array; pages: number; warnings: string[] } {
  if (typeof html !== 'string' || !html) throw new Error('저장할 편집 문서가 없습니다.');
  checkHwpHtmlSize(hwpHtmlBytes(html));
  if (!['hwp', 'hwpx'].includes(format)) throw new Error('지원하지 않는 한글 저장 형식입니다.');
  let doc = Core.createEmpty();
  try {
    // A real Hancom template supplies valid HWP 5.1 DocInfo and default styles.
    doc.createBlankDocument();
    const breakToken = 'WORDSTUDIO-PAGE-BREAK-' + crypto.randomUUID();
    const restoreGrid = planHwpGridRows(html);
    const tables = stageHwpTables(html);
    const sections = stageHwpSections(tables.html, page);
    const furniture = stageHwpFurniture(sections.html, sections.settings);
    const notes = stageHwpNotes(furniture.html);
    const pictures = stageHwpPictures(notes.html, nativeStructuredPasteBytes);
    const fields = stageHwpFields(pictures.html);
    const navigation = stageHwpTextNavigation(fields.html);
    const columns = stageHwpColumns(navigation.html, page, sections.columnDefaults);
    let characters = stageHwpCharacterStyles(columns.html);
    const convertPageBreaks = (body: string) => body.replace(/<div\b[^>]*data-page-break=["'][^"']*["'][^>]*>[\s\S]*?<\/div>/gi, `<p>${breakToken}</p>`);
    const paste = (body: string) => pasteHwpBody(doc, convertPageBreaks(body), breakToken, onProgress);
    const retryPaste = (body: string) => {
      const retry = Core.createEmpty();
      doc.free();
      doc = retry;
      doc.createBlankDocument();
      paste(body);
    };
    paste(characters.html);
    let withCharacters: HwpDocument | null;
    try { withCharacters = characters.restore(Core, doc); }
    catch (error) {
      if (!characters.hasSourceFonts) throw error;
      withCharacters = null;
    }
    const warnings: string[] = [];
    if (withCharacters === null && characters.hasSourceFonts) {
      // Font attribution is optional. Retry with the established width
      // transport so unsupported font runs cannot remove unrelated widths.
      characters = stageHwpCharacterStyles(columns.html, false);
      retryPaste(characters.html);
      withCharacters = characters.restore(Core, doc);
      warnings.push('일부 문자의 원본 글꼴을 한글 형식으로 옮기지 못했습니다.');
    }
    if (withCharacters === null) {
      // No temporary color or another run's ratio can leak into the file.
      retryPaste(columns.html);
      warnings.push('일부 글자의 장평을 한글 형식으로 옮기지 못했습니다.');
    } else if (withCharacters !== doc) { doc.free(); doc = withCharacters; }
    setHwpPage(doc, page);
    const textboxPaint: NativeBoxPaint = new Map();
    const withTextBoxes = restoreHwpExportTextBoxes(Core, doc, textBoxes, textboxPaint);
    if (withTextBoxes !== doc) { doc.free(); doc = withTextBoxes; }
    const withPictures = pictures.restore(Core, doc);
    if (withPictures !== doc) { doc.free(); doc = withPictures; }
    const withFields = fields.restore(Core, doc);
    if (withFields !== doc) { doc.free(); doc = withFields; }
    const withNavigation = navigation.restore(Core, doc);
    if (withNavigation !== doc) { doc.free(); doc = withNavigation; }
    const empty = nativePrint?.emptyParagraphs;
    const withEmptyParagraphs = empty ? restorePastedEmptyParagraphs(Core, doc, empty) : doc;
    if (nativePrint && empty && !withEmptyParagraphs) {
      const result = serializeHwp(Core, withoutPastedEmptyCarriers(html, empty), page, format, onProgress,
        textBoxes, { ...nativePrint, emptyParagraphs: undefined });
      result.warnings.push('일부 빈 문단의 글자 서식을 한글 형식으로 옮기지 못했습니다.');
      return result;
    }
    if (withEmptyParagraphs && withEmptyParagraphs !== doc) { doc.free(); doc = withEmptyParagraphs; }
    const withTables = tables.restore(Core, doc);
    if (withTables !== doc) { doc.free(); doc = withTables; }
    const withColumns = columns.restore(Core, doc);
    if (withColumns !== doc) { doc.free(); doc = withColumns; }
    const withNotes = notes.restore(Core, doc);
    if (withNotes !== doc) { doc.free(); doc = withNotes; }
    const withSections = sections.restore(Core, doc, setHwpPage);
    if (withSections !== doc) { doc.free(); doc = withSections; }
    const withFurniture = furniture.restore(Core, doc);
    if (withFurniture !== doc) { doc.free(); doc = withFurniture; }
    const withGrid = restoreGrid(Core, doc);
    if (withGrid !== doc) { doc.free(); doc = withGrid; }
    const withTextboxPaint = restoreNativeBoxPaint(Core, doc, textboxPaint);
    if (withTextboxPaint !== doc) { doc.free(); doc = withTextboxPaint; }
    const withTableLead = restoreNativeTableLead(Core,doc,nativePrint,format);
    if (withTableLead !== doc) { doc.free(); doc = withTableLead; }
    const withContinuationPlacement = restoreNativeContinuationPlacement(Core, doc, nativePrint?.continuations, nativePrint?.pages ?? 0);
    if (withContinuationPlacement !== doc) { doc.free(); doc = withContinuationPlacement; }
    const withPaperHeaderTable = restoreNativePaperHeaderTable(Core, doc, nativePrint?.paperHeaderTable);
    if (withPaperHeaderTable !== doc) { doc.free(); doc = withPaperHeaderTable; }
    const withPictureLine = restoreStandalonePictureLines(Core, doc, pictures.standaloneIds);
    if (withPictureLine !== doc) { doc.free(); doc = withPictureLine; }
    const withRasterTable = restoreNativeRasterTablePair(Core, doc, nativePrint?.rasterTablePairs, nativePrint?.pages);
    if (withRasterTable !== doc) { doc.free(); doc = withRasterTable; }
    const withEmptyFonts = restoreTerminalEmptyFonts(Core,doc,nativePrint?.terminalEmptyFonts,nativePrint?.pages);
    if (withEmptyFonts !== doc) { doc.free(); doc = withEmptyFonts; }
    const withTerminalBreak = restoreTerminalPageBreak(Core,doc,nativePrint?.terminalPageBreaks,nativePrint?.pages);
    if (withTerminalBreak !== doc) { doc.free(); doc = withTerminalBreak; }
    const withEmptyTableRows = restoreNativeEmptyTableRows(Core,doc,nativePrint?.emptyTableRows,nativePrint?.pages,format);
    if (withEmptyTableRows !== doc) { doc.free(); doc = withEmptyTableRows; }
    const withAuthoredTerminal = restoreAuthoredTerminalWrap(Core, doc, nativePrint?.authoredTerminalWraps, nativePrint?.pages);
    if (withAuthoredTerminal !== doc) { doc.free(); doc = withAuthoredTerminal; }
    const withSourcePageHead = restoreNativePageHead(Core, doc, nativePrint?.pageHeads, nativePrint?.pages);
    if (withSourcePageHead !== doc) { doc.free(); doc = withSourcePageHead; }
    const withAuthoredLine = restoreAuthoredNativeSingleLine(Core, doc, nativePrint?.authoredSingleLines, nativePrint?.pages);
    if (withAuthoredLine !== doc) { doc.free(); doc = withAuthoredLine; }
    onProgress?.({ label: `${format.toUpperCase()} 파일을 만드는 중…` });
    const output = format === 'hwp' ? doc.exportHwpWithReport() : doc.exportHwpxWithReport();
    try {
      const loss = output.contentLoss(), report = JSON.parse(loss);
      let bytes = output.takeBytes();
      if (bytes.byteLength > documentLimits.documentFileBytes) throw new Error(`변환한 한글 파일이 ${documentLimits.documentFileBytes / 1024 / 1024} MiB를 넘습니다.`);
      // A successful writer alone is insufficient: actually parse its result.
      onProgress?.({ label: '변환한 한글 파일을 다시 읽어 확인하는 중…' });
      const check = new Core(bytes);
      let pages = 0;
      try {
        pages = check.pageCount();
        if (!check.getSectionCount() || !pages) throw new Error('저장한 한글 문서를 다시 읽지 못했습니다.');
        const repaired = format === 'hwp' ?
          repairBinaryTerminalEmptyPage(Core,doc,check,nativePrint?.terminalEmptyFonts,nativePrint?.pages,loss) ??
          repairBinaryDuplicatePage(Core,doc,check,pages,loss) : undefined;
        if (repaired) ({bytes,pages} = repaired);
      } finally { check.free(); }
      if (report.count) warnings.push(`한글 저장 엔진에서 내용 손실 ${report.count}건을 보고했습니다. ${JSON.stringify(report.losses).slice(0, 1600)}`);
      return { bytes, pages, warnings };
    } finally { output.free(); }
  } finally { doc.free(); }
}
