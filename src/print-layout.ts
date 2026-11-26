import { layoutCheckpoint } from './editor-layout-task';
import { sourceFootnoteEndBottom } from './print-hwp-note-terminal';
import { hwpNestedCellFaceAttribute } from './hwp-nested-table-layout';
import { layoutHwpNestedCellFaces } from './hwp-nested-cell-face-print';
import { layoutHwpNestedTables } from './hwp-nested-table-print';
import { layoutHwpRelatedTables } from './hwp-related-table-print';
import { printedHwpContainerHeight, reserveHwpInlineStrokes } from './hwp-row-reserve';
import { activeHwpTableOverhang, firstTableOverhang, matchesHwpTableFragment, parseHwpTableContinuation, visibleHwpTableOverhang } from './hwp-table-continuation';
import { hwpKeptTablePage, hwpTableTopGap } from './table-anchor-data';
import { layoutHwpCenteredTableBorders, tableBorderFits } from './hwp-centered-table-borders';
import { layoutWordCellSpacing } from './word-cell-spacing';
import { layoutWordAutoFitTables } from './word-table-autofit';
import { preparePrintFonts, type PrintFontRequest } from './print-fonts';
import { wordKeepChains, wordKeepParagraph, wordKeepPrefixHeight, wordFlowHasWrappingAnchors, wordWidowSplitBounds } from './word-keep-policy';
import { prefersIntactCenteredRow, rowSplitPolicy } from './row-split-policy';
import { matchesHwpIntactTopRow } from './hwp-intact-row';
import { layoutTextBoxes } from './textbox-frame';
import { keepShortWordParagraphTogether, paragraphLineOffsets, paragraphTextWalker, prepareParagraphEnds, splitText } from './paragraph-fragments';
import { canSplitProtectedMergedRows, centeredCellFits, hasCenteredOrdinaryCell, hasHwpCellLines, measureCenteredCell, multiParagraphHwpOwners, ordinaryHwpTerminalCells, protectedTablePlacement, tableFragmenter } from './table-fragments';
import { hasUniformTablePageBody, pageMargins, pageFurniture, paperDimensions, activeDocxPrintPaper, type PageSettings } from './types';
import { layoutTabs } from './document-tabs';
import { layoutWordLines, markWordReferenceDensity, layoutWordListLines, markMeasuredWordSymbolItems } from './word-line-metrics';
import { layoutWordContextualSpacing } from './word-contextual-spacing';
import { layoutWordListLabels } from './word-list-labels';
import { layoutHwpCellLines } from './hwp-cell-line-metrics';
import { layoutHwpAuthoredCellTerminals } from './hwp-authored-cell-terminal';
import { layoutCellCharacterScale } from './cell-character-scale';
import { layoutHwpInlineTextBoxLines } from './hwp-inline-textbox-print';
import { hwpFloatingBoxFontFallback, hwpFloatingBoxProofSelector, layoutHwpFloatingTextBoxLines } from './hwp-floating-textbox-print';
import { layoutNestedHwpHangingScale } from './hwp-nested-hanging-layout';
import { layoutHwpSourceAdvances, layoutShortHwpCellWidths, layoutMixedHwpCellRuns, prepareHwpSourceFlow } from './hwp-text-advance-layout';
import { layoutHwpAuthoredHangingCellScale, layoutHwpCellWidthFallbacks, layoutHwpCompactCellFaces, layoutHwpExpandedCellRows, layoutHwpFittedCellLineEnds, layoutHwpScaledCellLines, layoutHwpSingleCellLines, layoutHwpSourceCellLines, markHwpSourceCellLineGroups, preserveHwpCellLinesBeforeSplit, restoreHwpContinuedTableLines } from './hwp-source-cell-lines';
import { clearSourceBodyEndPaint, layoutHwpAuthoredBodyLines, layoutHwpBodyLines, restoreHwpBodyBeforeKeptTable, restoreHwpBodyBeforePageBreak, sourceBodyEndBottom, sourceForcedBodyEndBottom } from './hwp-body-line-metrics';
import { sanitizeImportedHtml } from './document-html';
import { layoutColumns, columnGroupFits } from './column-layout';
import { followsWithHwpColumnBreak } from './hwp-column-end';
import { layoutHwpFootnotes } from './print-hwp-footnotes';
import { markSourceColumnTails } from './hwp-column-source-layout';
import { parseHwpBodyLine } from './hwp-body-line';
import { hwpBodyFallbackSelector, hwpCellFaceFallbackSelector, hwpCellWidthFallbackSelector, hwpSourceWrapAttribute, hwpWrapFingerprint, parseHwpSourceWrap } from './hwp-text-advances';
import { hwpCellRasterFillAttribute, hwpCellRasterPaintSource, parseHwpCellRasterFill } from './hwp-cell-raster-fill';
import { applyFloatingLayout, positionPaperImagesInPrintColumns } from './floating-layout';
import { layoutWordParagraphBottomBorders } from './word-paragraph-border-layout';
import { layoutParagraphBorderBoxes } from './paragraph-border-box-layout';
import { imageOutsidePaper, imageFloating } from './image-floating';
import { layoutFurnitureImages, tableStoryExtent } from './furniture-layout';
import { layoutTableAnchors, tableAnchor } from './table-anchor';
import { paragraphFitBottom } from './paragraph-fit';
import { splitListItem, wholeListItem } from './list-fragments';
import { canHideEmptyPageLine, hiddenEmptyLineClass } from './page-empty-lines';
import { isOrdinaryWordBodyParagraph, isWordAlignedBulletParagraph, isWordGroupedTextBoxParagraph, type WordLayoutOptions } from './word-trailing-spaces';

export type PrintLayout = { element: HTMLElement; pages: number; pageSettings: WeakMap<Element, PageSettings>; warnings: string[]; dispose(): void };
export type PrintLayoutProgress = { label: string; completed?: number; total?: number; unit?: string };
const PX = 96 / 25.4;
const dimensions = paperDimensions;
type PrintOptions = WordLayoutOptions & { nativeDesktop?: boolean; paintHwpCellLines?: boolean; hwpNativeExport?: boolean; interactive?: boolean };

/** Build one measured page tree used unchanged by preview, PDF and native printing. */
export async function createPrintLayout(html: string, initialPage: PageSettings,
  onProgress?: (value: PrintLayoutProgress) => void, signal?: AbortSignal, wordLayout: PrintOptions = {}): Promise<PrintLayout> {
  signal?.throwIfAborted();
  if (onProgress || wordLayout.interactive) {
    onProgress?.({ label: '쪽 배치를 준비하는 중…' });
    await new Promise(resolve => setTimeout(resolve, 24));
    signal?.throwIfAborted();
  }
  const host = document.createElement('div');
  host.className = 'print-layout measuring';
  if (wordLayout.language) host.lang = wordLayout.language;
  host.setAttribute('aria-label', '인쇄 미리보기 문서');
  document.body.append(host);
  const style = document.createElement('style'); host.append(style);
  const source = document.createElement('div'); source.innerHTML = html;
  if (wordLayout.language) source.lang = wordLayout.language;
  layoutHwpNestedTables(source);
  if (wordLayout.paintHwpCellLines !== false) layoutHwpRelatedTables(source, initialPage);
  layoutCellCharacterScale(source);
  // Columns are cloned and fragmented later, after a following break can
  // leave the fragment. Check the authored adjacency in the source DOM.
  for (const paragraph of source.querySelectorAll<HTMLElement>('p[data-hwp-body-line]')) {
    if (parseHwpBodyLine(paragraph.dataset.hwpBodyLine)?.align === 'column-end' && followsWithHwpColumnBreak(paragraph))
      paragraph.dataset.hwpColumnEndSource = 'true';
  }
  markSourceColumnTails(source);
  const flatSectionBoundaries = !!source.querySelector(':scope > [data-section-page]');
  const framesWithSectionBoundaries = new WeakSet(
    [...source.querySelectorAll('[data-section-content]')].filter(frame => frame.querySelector('[data-section-page]')),
  );
  const wordFloatingSpaces = wordLayout.wordHangingSpaces && !source.querySelector('[data-note-reference]');
  if (wordFloatingSpaces) for (const paragraph of source.children)
    if (paragraph instanceof HTMLParagraphElement && isOrdinaryWordBodyParagraph(paragraph, source))
      paragraph.dataset.wordFloatingSpaceSource = 'true';
  const wrappingAnchors = wordFlowHasWrappingAnchors(source);
  if (wrappingAnchors && wordLayout.wordBodyStrut) markWordReferenceDensity(source);
  const canPlaceProtectedTable = protectedTablePlacement(source);
  if (!wrappingAnchors && (wordLayout.wordHangingSpaces || wordLayout.wordBodyStrut)) {
    for (const paragraph of source.querySelectorAll<HTMLParagraphElement>('p[data-word-paragraph-style]')) {
      const ordinary = isOrdinaryWordBodyParagraph(paragraph, source);
      if (wordLayout.wordHangingSpaces && (ordinary || isWordAlignedBulletParagraph(paragraph)))
        paragraph.dataset.wordHangingSpaces = 'true';
      if (wordLayout.wordBodyStrut && ordinary) paragraph.dataset.wordBodyStrut = 'true';
    }
  }
  if (wordLayout.wordHangingSpaces) {
    for (const paragraph of source.querySelectorAll<HTMLParagraphElement>('[data-textbox-content] > p[data-word-paragraph-style]'))
      if (isWordGroupedTextBoxParagraph(paragraph)) paragraph.dataset.wordHangingSpaces = 'true';
  }
  prepareParagraphEnds(source);
  if (wordLayout.paintHwpCellLines !== false) markHwpSourceCellLineGroups(source);
  layoutWordParagraphBottomBorders(source);
  layoutParagraphBorderBoxes(source);
  layoutWordContextualSpacing(source);
  layoutWordListLabels(source);
  if (wordLayout.wordBodyStrut && !wrappingAnchors) markMeasuredWordSymbolItems(source);
  // The editable table NodeView fixes the width from colgroup; static HTML
  // must do the same or nested tables stretch to the whole printable width.
  for (const table of source.querySelectorAll('table')) {
    const columns = [...table.querySelectorAll<HTMLTableColElement>(':scope > colgroup > col')];
    if (columns.length && columns.every(c => /^\d+(?:\.\d+)?px$/.test(c.style.width)))
      table.style.width = `${columns.reduce((sum, c) => sum + parseFloat(c.style.width), 0)}px`;
  }
  const warnings = new Set<string>();
  const pageSettings = new WeakMap<Element, PageSettings>();
  const noteBodies = new Map<string, HTMLElement>();
  for (const note of source.querySelectorAll<HTMLElement>('[data-document-note][data-note-kind="footnote"]')) {
    noteBodies.set(note.dataset.documentNote!, note); note.remove();
  }
  const keepChains = wordKeepChains(source), protectedKeeps = new WeakSet<HTMLElement>();
  let hwpSourceFlow: Awaited<ReturnType<typeof prepareHwpSourceFlow>>;
  let settings = initialPage, pageNumber = 0, content: HTMLElement, footnotes: HTMLElement;
  let deferredNotes = new Set<string>();
  let noteGap = 12;
  let sectionNumber = 0;
  let sectionPage = 0;
  let hiddenEmptyLines = 0;
  let columnOverflowContent: HTMLElement | undefined;
  let sectionJustStarted = false;
  let sourceTableOverhang: ReturnType<typeof firstTableOverhang>;
  const overhangContinuations = new WeakSet<HTMLTableElement>();
  const pageStyles: string[] = [];
  function placeFootnotes(ids: string[]) {
    footnotes.replaceChildren(...ids.map(id => noteBodies.get(id)!.cloneNode(true)));
    footnotes.dataset.notes = ids.join('|');
    noteGap = layoutHwpFootnotes(footnotes) ?? 12;
  }
  function refreshNotes() {
    if (!content?.isConnected) return;
    const current = content.closest('.print-page')!;
    const fullBottom = content.getBoundingClientRect().top + Number(content.dataset.bodyHeight);
    const references = [...content.querySelectorAll<HTMLElement>('[data-note-reference]')];
    let ids = [...new Set(references
      .filter(el => el.getBoundingClientRect().top < fullBottom - .5).map(el => el.dataset.noteReference!))]
      .filter(id => !deferredNotes.has(id) && noteBodies.has(id) && ![...host.querySelectorAll<HTMLElement>('.print-footnotes [data-document-note]')].some(el => el.dataset.documentNote === id && el.closest('.print-page') !== current));
    if (footnotes.dataset.notes !== ids.join('|')) placeFootnotes(ids);
    const noteBoundary = fullBottom - (ids.length ? footnotes.getBoundingClientRect().height + noteGap : 0);
    // A later reference in a continuing row must move with its note. Keep
    // that decision while filling the other columns up to the page bottom;
    // otherwise a late note reserves blank space on the preceding page.
    for (const ref of references) if (ids.includes(ref.dataset.noteReference!) && rowSplitPolicy(ref.closest('tr')?.dataset) === 'allow' &&
      ref.getBoundingClientRect().bottom > noteBoundary + .5) deferredNotes.add(ref.dataset.noteReference!);
    if (ids.some(id => deferredNotes.has(id))) {
      ids = ids.filter(id => !deferredNotes.has(id));
      placeFootnotes(ids);
    }
    const reserve = ids.length ? footnotes.getBoundingClientRect().height + noteGap : 0;
    content.style.height = `${Math.max(1, Number(content.dataset.bodyHeight) - reserve)}px`;
  }
  function newPage(keepBlank = false) {
    hiddenEmptyLines = 0;
    refreshNotes();
    deferredNotes = new Set();
    noteGap = 12;
    const [width, height] = dimensions(settings), margins = pageMargins(settings);
    const pageName = `studioSection${sectionNumber}`;
    const precisePaper = activeDocxPrintPaper(settings);
    const [printWidth, printHeight] = precisePaper
      ? settings.orientation === 'landscape' ? [precisePaper[1], precisePaper[0]] : precisePaper
      : [width, height];
    pageStyles.push(`@page ${pageName} { size:${printWidth}mm ${printHeight}mm; margin:0; }`);

    const paper = document.createElement('section'); paper.className = 'print-page';
    pageSettings.set(paper, settings);
    paper.dataset.page = String(++pageNumber);
    if (keepBlank) paper.dataset.keepBlank = 'true';
    const first = ++sectionPage === 1;
    paper.dataset.sectionFirst = String(first);
    const furniture = pageFurniture(settings, pageNumber, first);
    Object.assign(paper.style, { width: `${width * PX}px`, height: `${height * PX}px`, page: pageName });
    content = document.createElement('div'); content.className = 'tiptap print-content';
    Object.assign(content.style, { left: `${margins.left * PX}px`, top: `${margins.top * PX}px`,
      width: `${(width - margins.left - margins.right) * PX}px`, height: `${(height - margins.top - margins.bottom) * PX}px` });
    const header = document.createElement('div'); header.className = 'print-header document-furniture';
    if (furniture.headerHtml) header.innerHTML = sanitizeImportedHtml(furniture.headerHtml).html; else header.textContent = furniture.header;
    const footer = document.createElement('div'); footer.className = 'print-footer';
    const footerText = document.createElement('div'); footerText.className = 'document-furniture';
    if (furniture.footerHtml) footerText.innerHTML = sanitizeImportedHtml(furniture.footerHtml).html; else footerText.textContent = furniture.footer;
    footer.append(footerText);
    for (const field of [...header.querySelectorAll('[data-page-number]'), ...footer.querySelectorAll('[data-page-number]')]) field.textContent = String(pageNumber);
    if (furniture.pageNumbers && !header.querySelector('[data-page-number]') && !footer.querySelector('[data-page-number]')) { const number = document.createElement('span'); number.className = 'print-automatic-number'; number.textContent = String(pageNumber); footer.append(number); }
    for (const el of [header, footer]) { el.style.left = `${margins.left * PX}px`; el.style.right = `${margins.right * PX}px`; }
    header.style.top = `${(settings.headerDistance ?? Math.max(0, margins.top / 2 - 3)) * PX}px`;
    footer.style.bottom = `${(settings.footerDistance ?? Math.max(0, margins.bottom / 2 - 3)) * PX}px`;
    layoutWordContextualSpacing(header); layoutWordContextualSpacing(footer);
    layoutWordListLabels(header); layoutWordListLabels(footer);
    paper.append(header, content, footer); host.append(paper);
    layoutWordListLines(header); layoutWordListLines(footer);
    layoutFurnitureImages(paper);
    const meaningful = (element: HTMLElement) => element.textContent?.trim() || element.querySelector('img,table');
    const story = {
      top: meaningful(header) ? header.offsetTop + header.getBoundingClientRect().height + 2 : 0,
      bottom: meaningful(footer) ? parseFloat(footer.style.bottom) + footer.getBoundingClientRect().height + 2 : 0,
    };
    const top = Math.max(margins.top * PX, story.top), bottom = Math.max(margins.bottom * PX, story.bottom);
    paper.dataset.uniformTablePageBody = String(hasUniformTablePageBody(settings, {
      top: Math.max(story.top, tableStoryExtent(header, 'top')),
      bottom: Math.max(story.bottom, tableStoryExtent(footer, 'bottom')),
    }));
    content.style.top = `${top}px`; content.style.height = `${Math.max(1, height * PX - top - bottom)}px`;
    content.dataset.bodyHeight = String(Math.max(1, height * PX - top - bottom));
    paper.dataset.pageHeight = String(height * PX);
    paper.dataset.margin = String(top); paper.dataset.bottomMargin = String(bottom);
    paper.dataset.leftMargin = String(margins.left * PX); paper.dataset.rightMargin = String(margins.right * PX);
    footnotes = document.createElement('div'); footnotes.className = 'tiptap print-footnotes';
    Object.assign(footnotes.style, { left: `${margins.left * PX}px`, right: `${margins.right * PX}px`, bottom: `${bottom}px` });
    paper.append(footnotes);
  }
  function fits() {
    refreshNotes();
    layoutTableAnchors(content);
    layoutWordAutoFitTables(content, true);
    layoutWordLines(content);
    layoutHwpCellLines(content);
    layoutHwpBodyLines(content, wordLayout.paintHwpCellLines !== false);
    layoutTabs(content);
    layoutColumns(content, content.getBoundingClientRect().bottom);
    // Removing overlay flow can move its zero-height frame to another
    // column. Resolve its coordinates against the final column geometry.
    layoutTableAnchors(content);
    layoutTextBoxes(content);
    applyFloatingLayout(content, wordFloatingSpaces);
    layoutTextBoxes(content);
    const last = content.lastElementChild as HTMLElement | null;
    const nativeEmptyLine = settings.hideEmptyLines === true && hiddenEmptyLines === 2;
    const overhang = activeHwpTableOverhang(sourceTableOverhang, last, content);
    const bodyBottom = (visibleHwpTableOverhang(overhang) ?? content.getBoundingClientRect().bottom) + .5;
    if (last) clearSourceBodyEndPaint(last);
    const preserveBottomLeading = settings.suppressBottomSpacing === false;
    const ordinaryBottom = last ? paragraphFitBottom(last, !preserveBottomLeading, nativeEmptyLine) : -Infinity;
    let paintedBottom: number | undefined;
    if (last && ordinaryBottom > bodyBottom && !preserveBottomLeading) {
      // The verified single terminal line fits the source page. Native
      // export needs that placement while leaving its paragraph untouched.
      if (wordLayout.hwpNativeExport) paintedBottom = sourceForcedBodyEndBottom(last, content, false);
      else if (wordLayout.paintHwpCellLines !== false)
        paintedBottom = sourceBodyEndBottom(last, content) ?? sourceForcedBodyEndBottom(last, content) ??
          sourceFootnoteEndBottom(last, content, footnotes);
    }
    return !hasDeferredNote(content) && (!last || (paintedBottom ?? ordinaryBottom) <= bodyBottom ||
      tableBorderFits(last, content, bodyBottom))
      // Behind/in-front objects do not consume text flow. A page decoration
      // extending into the margin must not push its anchor onto another page.
      && [...content.querySelectorAll<HTMLElement>('img[data-image-floating]')].every(img =>
        imageOutsidePaper(img) || imageFloating(img.dataset.imageFloating)?.wrap?.type === 0 || img.getBoundingClientRect().bottom <= content.getBoundingClientRect().bottom + .5)
      // Page-relative floating pictures can extend past a column box without
      // overflowing its text. scrollWidth includes those absolute descendants
      // and would move otherwise fitting columns onto another sheet.
      && [...content.querySelectorAll<HTMLElement>('.table-wrap-frame')].every(table => table.getBoundingClientRect().bottom <= content.getBoundingClientRect().bottom + .5)
      && [...content.querySelectorAll<HTMLElement>('[data-column-count]')].every(group =>
        columnGroupFits(group));
  }
  function hasDeferredNote(element: HTMLElement) {
    return [...element.querySelectorAll<HTMLElement>('[data-note-reference]')].some(ref => deferredNotes.has(ref.dataset.noteReference!));
  }
  function listBodyHeight(item: HTMLElement) {
    const ids = [...new Set([...item.querySelectorAll<HTMLElement>('[data-note-reference]')].map(ref => ref.dataset.noteReference!))]
      .filter(id => noteBodies.has(id) && ![...host.querySelectorAll<HTMLElement>('.print-footnotes [data-document-note]')]
        .some(note => note.dataset.documentNote === id && note.closest('.print-page') !== content.closest('.print-page')));
    let own = 0;
    for (const id of ids) {
      const existing = [...footnotes.children].find(note => (note as HTMLElement).dataset.documentNote === id);
      const note = existing || noteBodies.get(id)!.cloneNode(true) as HTMLElement;
      if (!existing) footnotes.append(note);
      own += 4 + note.getBoundingClientRect().height;
      if (!existing) note.remove();
    }
    return Math.min(content.getBoundingClientRect().height, Number(content.dataset.bodyHeight) - (ids.length ? 19 + own : 0));
  }
  function scaleOversized(element: HTMLElement) {
    const available = content.getBoundingClientRect().bottom - element.getBoundingClientRect().top;
    const height = element.getBoundingClientRect().height;
    if (height > available && available > 0) {
      element.style.zoom = String(Math.min(1, available / height));
      warnings.add('한 페이지보다 큰 그림·병합 셀 묶음은 잘리지 않도록 축소했습니다.');
    }
  }
  const { splitMergedGroup, splitMergedRow, splitElement } = tableFragmenter({
    bodyHeight: () => content.getBoundingClientRect().height, hasDeferredNote, deferredNotes: () => deferredNotes,
    preserveSplitCellWrap: paragraph => hwpSourceFlow?.preserveSplitCellWrap(paragraph,
      overhangContinuations.has(paragraph.closest('table')!)) ??
      (wordLayout.paintHwpCellLines !== false ? preserveHwpCellLinesBeforeSplit(paragraph) : undefined),
    keepCenteredMergeLabels: wordLayout.paintHwpCellLines !== false || (wordLayout.hwpNativeExport && 'first'),
  });
  type MeasuredCell = {
    cell: HTMLTableCellElement;
    before: NonNullable<ReturnType<typeof measureCenteredCell>>;
    ordinaryTerminal: boolean;
  };
  type ProtectedTablePlan = { fragments: HTMLTableRowElement[][]; measured: MeasuredCell[] };

  function measuredCellsFit(rows: HTMLTableRowElement[], measured: MeasuredCell[]) {
    return measured.every(({ cell, before, ordinaryTerminal }) => {
      if (!rows.includes(cell.parentElement as HTMLTableRowElement)) return true;
      const border = getComputedStyle(cell.closest('table')!).borderCollapse === 'collapse' ? .5 : 1;
      return centeredCellFits(cell, before, border, ordinaryTerminal);
    });
  }

  function planProtectedGroup(
    table: HTMLTableElement, headers: HTMLTableRowElement[], rows: HTMLTableRowElement[], firstCapacity: number, capacity: number,
    ordinaryTerminals: HTMLTableCellElement[], multiParagraphOwners: HTMLTableCellElement[],
  ): ProtectedTablePlan | null {
    const width = table.getBoundingClientRect().width;
    if (!(width > 0) || !(firstCapacity > 0) || !(capacity > 0)) return null;
    const holder = document.createElement('div');
    holder.className = 'tiptap';
    holder.style.cssText = `position:fixed;visibility:hidden;pointer-events:none;left:-100000px;top:0;margin:0;padding:0;min-height:0;width:${width}px`;
    const probe = table.cloneNode(false) as HTMLTableElement;
    probe.style.margin = '0';
    probe.style.width = `${width}px`;
    table.querySelectorAll(':scope > colgroup').forEach(columns => probe.append(columns.cloneNode(true)));
    const body = probe.createTBody();
    holder.append(probe);
    const refresh = () => { layoutWordLines(probe); layoutHwpCellLines(probe); layoutTabs(probe); };
    const copyRow = (row: HTMLTableRowElement) => row.cloneNode(true) as HTMLTableRowElement;
    const repeats = headers.map(copyRow);
    const fragments: HTMLTableRowElement[][] = [];
    let pending = rows.map(copyRow);
    const measured: MeasuredCell[] = [];
    for (const cell of [...ordinaryTerminals, ...multiParagraphOwners]) {
      const before = measureCenteredCell(cell);
      const row = pending[rows.indexOf(cell.parentElement as HTMLTableRowElement)];
      const ordinaryTerminal = ordinaryTerminals.includes(cell);
      if (!before || ordinaryTerminal && !before[0].terminal || !row?.cells[cell.cellIndex]) return null;
      measured.push({ cell: row.cells[cell.cellIndex], before, ordinaryTerminal });
    }
    content.append(holder);
    try {
      // Plan every continuation before changing the visible page. A failure
      // on any later whole row leaves the source table untouched.
      for (let pass = 0; pending.length && pass < 512; pass++) {
        body.replaceChildren(...repeats, ...pending);
        prepareParagraphEnds(probe);
        refresh();
        const height = pass ? capacity : firstCapacity;
        const fitsHere = () => {
          refresh();
          return probe.getBoundingClientRect().height <= height + .5 && measuredCellsFit([...body.rows], measured);
        };
        if (fitsHere()) {
          fragments.push(pending);
          return fragments.length > 1 ? { fragments, measured } : null;
        }
        const tail = splitMergedGroup(pending, probe.getBoundingClientRect().top + height, fitsHere);
        if (!tail) return null;
        fragments.push(pending);
        pending = tail;
      }
      return null;
    } finally { holder.remove(); }
  }

  // Keep the existing first body start when it is already on the source
  // sheet, even if the preceding table was placed one sheet earlier.
  function followingBodyStartsHere(next: Element | null, sourcePage: number) {
    const paragraph = next?.nextElementSibling;
    if (!(next instanceof HTMLParagraphElement) || !(paragraph instanceof HTMLParagraphElement) ||
        next.firstElementChild || next.textContent?.trim() ||
        [next, paragraph].some(node => node.matches('[data-page-break],[data-column-break],[data-section-page]')) ||
        paragraph.querySelector('img,svg,table,[data-note-reference],[data-document-textbox]')) return false;
    const proof = parseHwpSourceWrap(paragraph.dataset.hwpSourceWrap);
    if (proof?.authoredBodyLines?.origin.page !== sourcePage ||
        proof.fingerprint !== hwpWrapFingerprint(paragraph.textContent ?? '')) return false;

    const blank = next.cloneNode(true) as HTMLParagraphElement;
    const sample = paragraph.cloneNode(true) as HTMLParagraphElement;
    let fragment: HTMLElement | undefined;
    try {
      content.append(blank);
      if (!fits()) return false;
      content.append(sample);
      if (fits()) return true;
      const offsets = paragraphLineOffsets(sample), height = content.getBoundingClientRect().height;
      const lines = offsets.length + 1;
      if (!offsets.length || keepShortWordParagraphTogether(sample, lines, height)) return false;
      const bounds = wordWidowSplitBounds(sample, lines, height, 1, () => offsets[0]);
      const minimum = bounds?.minimumHead ?? 1;
      const cut = offsets[minimum - 1];
      if (!cut || bounds && minimum > bounds.maximumHead) return false;
      [fragment] = splitText(sample, cut);
      sample.replaceWith(fragment);
      return fits();
    } finally {
      blank.remove();
      sample.remove();
      fragment?.remove();
      fits();
    }
  }

  function* appendTable(table: HTMLTableElement): Generator<void> {
    const next = table.nextElementSibling;
    const beforeBreak = next?.matches('[data-page-break]') || next?.matches('p:empty') && next.nextElementSibling?.matches('[data-page-break]');
    if (beforeBreak && table.rows.length === 1) {
      const row = table.rows[0], observedHeight = printedHwpContainerHeight(row);
      if (observedHeight !== undefined) row.style.height = `${observedHeight}px`;
    }
    const frame = table.closest('[data-section-content]');
    // The editor reserves one envelope for an entire section frame. Keep
    // nonempty stories out of the measured path when that frame has a break.
    const hasSectionBoundary = frame ? framesWithSectionBoundaries.has(frame) : flatSectionBoundaries;
    const uniformTablePageBody = () => content.closest<HTMLElement>('.print-page')?.dataset.uniformTablePageBody === 'true' &&
      (!hasSectionBoundary || hasUniformTablePageBody(settings));
    const [width, height] = dimensions(settings), { top, bottom } = pageMargins(settings);
    const protectedTableFlow = canPlaceProtectedTable(table, { width: width * PX, height: height * PX, top: top * PX, bottom: bottom * PX });
    // Measure the complete source table once. Every printed continuation must
    // retain the same columns, including when its widest cell appears later.
    content.append(table);
    layoutWordAutoFitTables(content, true);
    layoutHwpCenteredTableBorders(table, wordLayout.paintHwpCellLines !== false);
    layoutShortHwpCellWidths(table);
    layoutMixedHwpCellRuns(table);
    if (wordLayout.paintHwpCellLines !== false) {
      layoutNestedHwpHangingScale(table);
      layoutHwpSingleCellLines(table);
      layoutHwpScaledCellLines(table);
      layoutHwpExpandedCellRows(table);
      if (hwpSourceFlow?.missingCompactCellFace) layoutHwpCellWidthFallbacks(table);
      layoutHwpAuthoredHangingCellScale(table);
      layoutHwpAuthoredCellTerminals(table);
      if (hwpSourceFlow?.missingCompactCellFace || hwpSourceFlow?.missingCompactHeaderFace)
        layoutHwpCompactCellFaces(table, hwpSourceFlow.missingCompactCellFace, hwpSourceFlow.missingCompactHeaderFace);
    } else if (wordLayout.hwpNativeExport && hwpSourceFlow?.missingCompactCellFace) {
      layoutHwpCellWidthFallbacks(table, true);
    }
    if ((wordLayout.paintHwpCellLines !== false || wordLayout.hwpNativeExport) && hwpSourceFlow?.missingSingleBodyFace &&
        table.querySelector(`[${hwpNestedCellFaceAttribute}]`))
      layoutHwpNestedCellFaces(table, wordLayout.hwpNativeExport);
    if (hwpSourceFlow && table.rows.length === 1 && table.rows[0].cells.length === 1 && table.querySelector('p[data-hwp-source-wrap]')) {
      layoutHwpCellLines(table);
      for (const paragraph of table.rows[0].cells[0].children)
        if (paragraph instanceof HTMLParagraphElement) hwpSourceFlow.restoreMissingWrap(paragraph);
    }
    // An earlier layout change can already have moved the preceding
    // table to another page. In that case leave the current flow intact.
    if (hwpKeptTablePage(table.dataset.hwpKeptTablePage) === pageNumber + 1 && !table.dataset.tableAnchor &&
        !table.closest('[data-column-count]') && !table.querySelector('img,svg,table,[data-note-reference],[data-document-textbox]')) {
      const box = table.getBoundingClientRect(), body = content.getBoundingClientRect();
      if (box.width <= body.width + .5 && box.height > 0 && box.height <= body.height + .5) {
        if (fits() && (wordLayout.paintHwpCellLines === false || !followingBodyStartsHere(next, pageNumber + 1) ||
            !restoreHwpBodyBeforeKeptTable(content, table))) return;
        if (content.children.length > 1) {
          table.remove();
          newPage();
          content.append(table);
          if (fits()) return;
        }
      }
    }
    // DOCX spacing is measured from the regular flow. The printed copy can
    // recover HWP line ends without converting them into Word line spacing.
    const continuationProof = wordLayout.paintHwpCellLines !== false &&
      !table.parentElement?.closest('td,th,[data-column-count],[data-document-textbox]')
      ? parseHwpTableContinuation(table.dataset.hwpTableContinuation) : undefined;
    const firstFragment = continuationProof?.overhang ? undefined : continuationProof;
    const overhangRow = continuationProof?.overhang && continuationProof.rows === table.rows.length
      ? table.rows[continuationProof.row] : undefined;
    // The independent cell witness applies to a reopened DOCX grid. When
    // the table has a complete first-fragment witness, respect that policy.
    const reopenedCellFlow = !firstFragment && table.hasAttribute('data-word-table-layout') &&
      !!table.querySelector(':scope > colgroup > col');
    table.remove();
    const rows = Array.from(table.rows);
    const rowIndices = firstFragment?.rows === rows.length ? new WeakMap(rows.map((row, index) => [row, index])) : undefined;
    let nextFragment = false, finishedFirstFragment = false;
    const headers: HTMLTableRowElement[] = [];
    while (rows.length && rows[0].cells.length > 0 && [...rows[0].cells].every(cell => cell.tagName === 'TH')) headers.push(rows.shift()!);
    if (!headers.length && table.dataset.hwpTableTopGap) {
      delete table.dataset.hwpTableTopGap;
      table.style.removeProperty('margin-top');
    }
    let retainedOverhang = false;
    const shell = (continuation = false) => {
      const next = table.cloneNode(false) as HTMLTableElement;
      if (retainedOverhang) overhangContinuations.add(next);
      if (continuation && next.dataset.tableAnchor) {
        const anchor = tableAnchor(next.dataset.tableAnchor);
        if (anchor?.reference === 'paragraph') {
          next.dataset.tableAnchor = JSON.stringify({ ...anchor, y: 0 });
          next.style.marginTop = '0px';
          if (wordLayout.hwpNativeExport && next.dataset.hwpCenteredBorders === 'continuation')
            next.setAttribute('data-print-native-continued-table', '');
        }
      }
      table.querySelectorAll(':scope > colgroup').forEach(el => next.append(el.cloneNode(true)));
      const body = document.createElement('tbody'); next.append(body);
      headers.forEach(row => body.append(row.cloneNode(true)));
      content.append(next); return { table: next, body };
    };
    let part = shell();
    if (!rows.length) { if (!fits() && content.children.length > 1) { part.table.remove(); newPage(); part = shell(); } scaleOversized(part.table); return; }
    let bodyRows = 0;
    const protectedStarts = new WeakSet<HTMLTableRowElement>();
    const openContinuation = (index: number) => {
      if (rowIndices && firstFragment && !finishedFirstFragment && bodyRows > 0) {
        finishedFirstFragment = true;
        nextFragment = matchesHwpTableFragment(part.table, content, firstFragment, row => rowIndices.get(row));
      }
      newPage(); part = shell(index > 0); bodyRows = 0;
    };
    const applyProtectedPlan = (group: HTMLTableRowElement[], { fragments, measured }: ProtectedTablePlan) => {
      const source = { part, content, footnotes, pageNumber, sectionPage, hiddenEmptyLines, deferredNotes };
      const styleCount = pageStyles.length;
      const pages: Element[] = [];
      group.forEach(row => row.remove());
      part.body.append(...fragments[0]);
      let complete = fits() && measuredCellsFit(fragments[0], measured);
      for (const pageRows of fragments.slice(1)) {
        if (!complete) break;
        newPage();
        pages.push(content.closest('.print-page')!);
        part = shell(true);
        part.body.append(...pageRows);
        complete = uniformTablePageBody() && fits() && measuredCellsFit(pageRows, measured);
      }
      if (complete) return true;

      // Recheck against the real page shells even after the detached proof.
      // Roll back paper numbering as well as the source rows on failure.
      pages.forEach(page => page.remove());
      part = source.part; content = source.content; footnotes = source.footnotes;
      pageNumber = source.pageNumber; sectionPage = source.sectionPage;
      hiddenEmptyLines = source.hiddenEmptyLines; deferredNotes = source.deferredNotes;
      pageStyles.length = styleCount;
      fragments[0].forEach(row => row.remove());
      part.body.append(...group);
      return false;
    };
    for (let i = 0; i < rows.length;) {
      // The preceding complete row group has settled. Never suspend the
      // probing, rollback or continuation of a protected merge halfway.
      if (wordLayout.interactive) yield;
      if (protectedStarts.has(rows[i])) openContinuation(i);
      let end = i + 1;
      for (let j = i; j < end; j++) for (const cell of rows[j].cells) end = Math.min(rows.length, Math.max(end, j + cell.rowSpan));
      const originals = rows.slice(i, end);
      const group = originals.map(row => row.cloneNode(true) as HTMLTableRowElement);
      group.forEach((row, index) => {
        const sourceIndex = rowIndices?.get(originals[index]);
        if (sourceIndex !== undefined) rowIndices!.set(row, sourceIndex);
        part.body.append(row);
      });
      const verifiedPrior = nextFragment && firstFragment && rowIndices &&
        (rowIndices.get(group[0]) ?? -1) > firstFragment.row;
      if (wordLayout.paintHwpCellLines !== false && group.length === 1 &&
          (verifiedPrior || reopenedCellFlow && group[0].querySelector('p[data-hwp-source-wrap*="splitCell"]')))
        restoreHwpContinuedTableLines(group[0], !!verifiedPrior);
      if (overhangRow && originals[0] === overhangRow && group.length === 1)
        sourceTableOverhang = firstTableOverhang(part.table, group[0], content, continuationProof!, pageNumber);
      const split = () => {
        const paper = content.getBoundingClientRect();
        const overhang = activeHwpTableOverhang(sourceTableOverhang, part.table, content);
        if (overhang) {
          const original = group[0].cloneNode(true) as HTMLTableRowElement;
          const tail = splitElement(group[0], overhang.bottom);
          if (tail && fits()) {
            retainedOverhang = true;
            rows.splice(end, 0, tail as HTMLTableRowElement);
            return;
          }
          group[0].replaceWith(original);
          group[0] = original;
          sourceTableOverhang = undefined;
        }
        let protectedFit: typeof fits | undefined;
        if (group.length > 1 && protectedTableFlow && uniformTablePageBody() && !ordinaryHwpTerminalCells(group).length && !multiParagraphHwpOwners(group).length) {
          const repeated = [...part.body.rows].slice(0, headers.length);
          if (canSplitProtectedMergedRows([...repeated, ...group])) {
            const tableBox = part.table.getBoundingClientRect();
            const headerHeight = (repeated.at(-1)?.getBoundingClientRect().bottom ??
              part.table.rows[0].getBoundingClientRect().top) - tableBox.top;
            const groupHeight = tableBox.bottom - group[0].getBoundingClientRect().top;
            if (headerHeight + groupHeight <= paper.height + .5) protectedFit = fits;
          }
        }
        const rowTail = !protectedFit && splitMergedRow(group, paper.bottom, fits);
        if (rowTail) {
          if (rowIndices && group.length === 1) {
            const sourceIndex = rowIndices.get(group[0]);
            if (sourceIndex !== undefined) rowIndices.set(rowTail[0], sourceIndex);
          }
          protectedStarts.add(rowTail[0]);
          rows.splice(end, 0, ...rowTail); return;
        }
        const mergedTail = splitMergedGroup(group, paper.bottom, protectedFit);
        if (mergedTail) {
          if (rowIndices && group.length === 1) {
            const sourceIndex = rowIndices.get(group[0]);
            if (sourceIndex !== undefined) rowIndices.set(mergedTail[0], sourceIndex);
          }
          if (protectedFit) protectedStarts.add(mergedTail[0]);
          rows.splice(end, 0, ...mergedTail); return;
        }
        if (group.length !== 1 || [...group[0].cells].some(c => c.rowSpan !== 1)) return;
        // DOCX enables row splitting to preserve continuing vertical merges.
        // A final one-row fragment is still part of that merge: keep its row
        // boundary instead of moving only its empty height to the next sheet.
        if (group[0].querySelector(':scope > [data-print-merge-continuation]') && group[0].getBoundingClientRect().height <= content.getBoundingClientRect().height + .5) return;
        if (rowSplitPolicy(group[0].dataset) !== 'allow' && !group[0].querySelector('table') && group[0].getBoundingClientRect().height <= content.getBoundingClientRect().height + .5) return;
        const tail = splitElement(group[0], paper.bottom);
        if (tail) {
          const sourceIndex = rowIndices?.get(group[0]);
          if (sourceIndex !== undefined) rowIndices!.set(tail as HTMLTableRowElement, sourceIndex);
          rows.splice(end, 0, tail as HTMLTableRowElement);
        }
      };
      // A successful fit already laid out this unchanged row group. Repeating
      // every normalizer three times made large-table DOCX saves stall. A
      // split or a move changes geometry and always needs a new measurement.
      let groupFits = fits();
      if (groupFits && rows.length === 1 && group[0].dataset.hwpRowReserve) reserveHwpInlineStrokes(part.table);
      if (!groupFits && i === 0 && group.length > 1 && protectedTableFlow && uniformTablePageBody() &&
        group.some(row => [...row.cells].some(cell => cell.rowSpan > 1 && getComputedStyle(cell).verticalAlign === 'middle'))) {
        const paper = content.getBoundingClientRect(), tableBox = part.table.getBoundingClientRect();
        const repeats = [...part.body.rows].slice(0, headers.length);
        const ordinaryTerminals = ordinaryHwpTerminalCells(group);
        const multiParagraphOwners = multiParagraphHwpOwners(group);
        const fitsFresh = tableBox.height <= paper.height + .5;
        const canPlan = ordinaryTerminals.length || multiParagraphOwners.length ? fitsFresh :
          !fitsFresh && headers.length > 0 && !hasCenteredOrdinaryCell(group) && !hasHwpCellLines([...repeats, ...group]);
        if (canPlan && canSplitProtectedMergedRows([...repeats, ...group])) {
          const plan = planProtectedGroup(part.table, repeats, group, paper.bottom - tableBox.top, paper.height, ordinaryTerminals, multiParagraphOwners);
          if (plan && applyProtectedPlan(group, plan)) {
            bodyRows = plan.fragments.at(-1)!.length; i = end; continue;
          }
          if (plan) groupFits = fits();
        }
      }
      // cantSplit also governs oversized rows: move their start to a fresh
      // body before continuing them, rather than shrinking or filling a tail.
      let fitsIntactOnFreshPage = false;
      if (!groupFits && i > 0 && group.length === 1 && (prefersIntactCenteredRow(group[0]) ||
          wordLayout.paintHwpCellLines !== false && !firstFragment && headers.length > 0 &&
          matchesHwpIntactTopRow(group[0], group[0].getBoundingClientRect().height, 1))) {
        const tableBox = part.table.getBoundingClientRect();
        const headerBottom = headers.length ? part.body.rows[headers.length - 1].getBoundingClientRect().bottom
          : part.body.rows[0].getBoundingClientRect().top;
        const headerGap = hwpTableTopGap(part.table.dataset.hwpTableTopGap) ?? 0;
        const rowTop = group[0].getBoundingClientRect().top;
        const freshHeight = headerBottom - tableBox.top + headerGap + tableBox.bottom - rowTop;
        fitsIntactOnFreshPage = freshHeight <= content.getBoundingClientRect().height + .5;
      }
      if (!groupFits && group.length === 1 && (rowSplitPolicy(group[0].dataset) === 'avoid' || fitsIntactOnFreshPage) &&
        (bodyRows > 0 || content.children.length > 1)) {
        group.forEach(row => row.remove());
        if (!bodyRows) part.table.remove();
        openContinuation(i);
        group.forEach(row => part.body.append(row)); groupFits = fits();
      }
      if (!groupFits) { split(); groupFits = fits(); }
      if (!groupFits && (bodyRows > 0 || content.children.length > 1)) {
        group.forEach(row => row.remove());
        if (!bodyRows) part.table.remove();
        openContinuation(i);
        group.forEach(row => part.body.append(row));
        groupFits = fits();
        if (!groupFits) { split(); groupFits = fits(); }
      }
      if (!groupFits) scaleOversized(part.table);
      bodyRows += group.length; i = end;
    }
  }
  const pairedParagraphs = new WeakSet<HTMLElement>();
  function* append(element: HTMLElement): Generator<void> {
    if (pairedParagraphs.has(element)) return;
    if (element.hasAttribute('data-section-page')) {
      const next = JSON.parse(element.getAttribute('data-section-page')!) as PageSettings;
      const sameSize = dimensions(settings).join(',') === dimensions(next).join(',');
      settings = next; sectionNumber++; sectionPage = 0;
      if (next.breakType === 'continuous' && sameSize) {
        sectionPage = 1;
        const paper = content.closest<HTMLElement>('.print-page')!, furniture = pageFurniture(settings, pageNumber, true);
        const header = paper.querySelector<HTMLElement>('.print-header')!, footer = paper.querySelector<HTMLElement>('.print-footer .document-furniture')!;
        if (furniture.headerHtml) header.innerHTML = sanitizeImportedHtml(furniture.headerHtml).html; else header.textContent = furniture.header;
        if (furniture.footerHtml) footer.innerHTML = sanitizeImportedHtml(furniture.footerHtml).html; else footer.textContent = furniture.footer;
        paper.querySelector('.print-automatic-number')?.remove();
        if (furniture.pageNumbers && !header.querySelector('[data-page-number]') && !footer.querySelector('[data-page-number]')) {
          const number = document.createElement('span'); number.className = 'print-automatic-number'; number.textContent = String(pageNumber); footer.parentElement!.append(number);
        }
        layoutFurnitureImages(paper);
        paper.dataset.uniformTablePageBody = String(hasUniformTablePageBody(settings));
        return;
      }
      const meaningful = content.textContent?.trim() || content.querySelector('img,table,[data-note-reference]');
      if (!meaningful) { content.closest('.print-page')!.remove(); pageNumber--; }
      if (next.breakType === 'oddPage' || next.breakType === 'evenPage') {
        const target = next.breakType === 'oddPage' ? 1 : 0;
        if ((pageNumber + 1) % 2 !== target) { newPage(true); sectionPage = 0; }
      }
      newPage(true); sectionJustStarted = true; return;
    }
    // A break outside a multi-column group advances to the next sheet,
    // including breaks unwrapped from a one-column section.
    if (element.hasAttribute('data-page-break') || element.hasAttribute('data-column-break')) {
      if (sectionJustStarted && !content.textContent?.trim() && !content.querySelector('img,table')) { sectionJustStarted = false; return; }
      const group = content.firstElementChild;
      const blank = group?.firstElementChild;
      // The empty-line option also applies when fitting columns leaves an
      // unpainted line at the beginning of an automatically created sheet.
      // Keep the line and use that same sheet for the pending page break.
      if (element.hasAttribute('data-page-break') && settings.hideEmptyLines && content === columnOverflowContent &&
          content.children.length === 1 && group?.matches('[data-column-count]') && group.children.length === 1 &&
          blank instanceof HTMLElement && canHideEmptyPageLine(blank)) {
        blank.classList.add(hiddenEmptyLineClass);
        group.replaceWith(blank);
        hiddenEmptyLines = 1;
        content.parentElement!.dataset.keepBlank = 'true';
        columnOverflowContent = undefined;
        return;
      }
      newPage(true); return;
    }
    if (element.hasAttribute('data-section-content')) { for (const child of [...element.children]) yield* append(child as HTMLElement); return; }
    if (element.tagName === 'TABLE' && tableAnchor(element.dataset.tableAnchor)?.wrap && element.nextElementSibling?.matches('p,h1,h2,h3,h4,h5,h6')) {
      const following = element.nextElementSibling as HTMLElement, prior = content.children.length;
      content.append(element, following); fits();
      if (element.classList.contains('table-wrap-frame')) {
        // A wrapping table belongs to the next regular paragraph in OOXML.
        // Its first actual text line must travel with it even when an edit
        // makes that line taller than the original label's row.
        const cut = paragraphLineOffsets(following)[0];
        const first = cut ? splitText(following, cut)[0] : following.cloneNode(true) as HTMLElement;
        following.remove(); content.append(first);
        const firstFits = fits(); first.remove();
        if (!firstFits && prior) { element.remove(); newPage(); content.append(element); }
        yield* append(following); pairedParagraphs.add(following); return;
      }
      // Resized/oversized labels use the ordinary splittable table path.
      following.remove(); element.remove();
    }
    if (element.tagName === 'TABLE' && tableAnchor(element.dataset.tableAnchor)?.overlay) {
      content.append(element); fits();
      if (element.classList.contains('table-overlay-active')) return;
      element.remove();
      element.removeAttribute('data-table-anchor');
    }
    if (element.hasAttribute('data-column-count')) {
      if ([...element.children].some(child => child.hasAttribute('data-page-break') || child.hasAttribute('data-section-page'))) {
        let part = element.cloneNode(false) as HTMLElement;
        for (const child of [...element.children]) {
          if (child.hasAttribute('data-page-break') || child.hasAttribute('data-section-page')) {
            if (part.children.length) yield* append(part);
            yield* append(child as HTMLElement); part = element.cloneNode(false) as HTMLElement;
          } else part.append(child);
        }
        if (part.children.length) yield* append(part);
        return;
      }
      if (Number(element.dataset.columnCount) === 1) { for (const child of [...element.children]) yield* append(child as HTMLElement); return; }
      content.append(element); if (fits()) return;
      // Split at block boundaries first. Character-only ranges lose the
      // position of empty paragraphs, pictures and forced column breaks.
      const children = [...element.children] as HTMLElement[];
      const fragment = (count: number) => {
        const part = element.cloneNode(false) as HTMLElement;
        part.replaceChildren(...children.slice(0, count).map(child => child.cloneNode(true)));
        return part;
      };
      element.remove(); let low = 1, high = children.length, best = 0;
      while (low <= high) {
        const mid = (low + high) >> 1, head = fragment(mid); content.append(head);
        if (fits()) { best = mid; low = mid + 1; } else high = mid - 1;
        head.remove();
      }
      const head = fragment(best), tail = element.cloneNode(false) as HTMLElement;
      tail.replaceChildren(...children.slice(best).map(child => child.cloneNode(true)));
      content.append(head);
      const next = tail.firstElementChild as HTMLElement | null;
      if (next?.matches('p,h1,h2,h3,h4,h5,h6') && ((next.textContent?.length ?? 0) > 1 || next.querySelector('br'))) {
        const walker = paragraphTextWalker(next); let text = '';
        while (walker.nextNode()) text += walker.currentNode instanceof HTMLBRElement ? '\n' : walker.currentNode.textContent;
        const offsets = [...new Intl.Segmenter(undefined, { granularity: 'grapheme' }).segment(text)].map(s => s.index).slice(1);
        let left = 0, right = offsets.length - 1, cut = 0;
        while (left <= right) {
          const mid = (left + right) >> 1, [part] = splitText(next, offsets[mid]); head.append(part);
          if (fits()) { cut = offsets[mid]; left = mid + 1; } else right = mid - 1;
          part.remove();
        }
        if (cut) { const [part, rest] = splitText(next, cut); head.append(part); next.replaceWith(rest); }
      }
      if (head.children.length) {
        fits(); newPage();
        columnOverflowContent = content;
        // The break that overflowed the last column is fulfilled by this new
        // sheet. Applying it again at its start would skip the first column.
        if (tail.firstElementChild?.hasAttribute('data-column-break')) tail.firstElementChild.remove();
        if (tail.children.length) yield* append(tail);
        return;
      }
      head.remove();
      if (content.children.length) { newPage(); yield* append(element); return; }
      content.append(element); fits(); scaleOversized(element); return;
    }
    if (element.textContent?.trim() || element.querySelector('img,table')) sectionJustStarted = false;
    const table = element.matches('table') ? element : element.querySelector(':scope > table');
    if (table) { yield* appendTable(table as HTMLTableElement); return; }
    if (element.matches('ul,ol')) {
      const children = [...element.children] as HTMLElement[];
      let shell = element.cloneNode(false) as HTMLElement; content.append(shell);
      let index = Number(element.getAttribute('start')) || 1;
      const continueList = () => {
        if (!shell.children.length) shell.remove();
        newPage(); shell = element.cloneNode(false) as HTMLElement;
        if (shell.tagName === 'OL') shell.setAttribute('start', String(index));
        content.append(shell);
      };
      for (const original of children) {
        let child = original; shell.append(child);
        // A short item is indivisible. A larger one continues at the same
        // measured lines / nested-item boundaries used by editor pagination.
        while (!fits()) {
          if (!wholeListItem(child.getBoundingClientRect().height, listBodyHeight(child))) {
            const tail = splitListItem(child, { fits, bodyHeight: listBodyHeight });
            if (tail) { continueList(); child = tail; shell.append(child); continue; }
          }
          if (shell.children.length > 1 || content.children.length > 1) {
            child.remove(); continueList(); shell.append(child); continue;
          }
          scaleOversized(shell); break;
        }
        index++;
      }
      return;
    }
    if (element.matches('blockquote')) {
      const children = [...element.children] as HTMLElement[];
      let shell = element.cloneNode(false) as HTMLElement; content.append(shell);
      for (const child of children) {
        shell.append(child);
        if (!fits() && (shell.children.length > 1 || content.children.length > 1)) {
          child.remove(); if (!shell.children.length) shell.remove();
          newPage(); shell = element.cloneNode(false) as HTMLElement;
          content.append(shell); shell.append(child);
        }
        if (!fits()) scaleOversized(shell);
      }
      return;
    }
    const followedByPageBreak = element.nextElementSibling?.matches('[data-page-break]');
    if (followedByPageBreak && element.dataset.hwpSourceWrap?.includes('"forcedBodyEnd"'))
      element.dataset.hwpPrintBeforeExplicitBreak = 'true';
    content.append(element);
    if (wordLayout.paintHwpCellLines !== false && element.hasAttribute('data-document-textbox'))
      layoutHwpInlineTextBoxLines(element);
    if (wordKeepParagraph(element) && (element.dataset.wordKeepNext === 'true' || element.dataset.wordKeepLines === 'true')) {
      fits();
      const bodyHeight = content.getBoundingClientRect().height;
      let budget = 0;
      const chain = !protectedKeeps.has(element) && keepChains.get(element);
      if (chain) {
        for (const heading of chain.slice(0, -1)) protectedKeeps.add(heading);
        // Measure the future siblings in the same page width and line rules.
        // Clones are view-only and never enter the editor or saved document.
        const probes = chain.slice(1).map(p => p.cloneNode(true) as HTMLElement);
        content.append(...probes); fits();
        const last = probes[probes.length - 1];
        budget = last.getBoundingClientRect().top - element.getBoundingClientRect().top + wordKeepPrefixHeight(last, bodyHeight);
        probes.forEach(p => p.remove()); fits();
      }
      if (!budget && element.dataset.wordKeepLines === 'true' && !element.dataset.printParagraphFragment && element.getBoundingClientRect().height > bodyHeight + .5) budget = element.getBoundingClientRect().height;
      if (budget && content.children.length > 1 && element.getBoundingClientRect().top + budget > content.getBoundingClientRect().bottom + .5) {
        element.remove(); newPage();
        // Word suppresses space-before at an automatic keep break. The editor
        // gap already places this paragraph at the body top. Retain the saved
        // spacing; this adjustment belongs only to the new printed fragment.
        element.style.marginTop = '0'; yield* append(element); return;
      }
    }
    let fitsPage = fits();
    if (!fitsPage && followedByPageBreak && hwpSourceFlow?.missingSingleBodyFace &&
        (wordLayout.paintHwpCellLines !== false || wordLayout.hwpNativeExport) && element instanceof HTMLParagraphElement)
      fitsPage = restoreHwpBodyBeforePageBreak(element, fits, !wordLayout.hwpNativeExport);
    if (fitsPage && hwpSourceFlow && element instanceof HTMLParagraphElement) {
      const restore = hwpSourceFlow.restoreMissingWrap(element);
      if (restore && !fits()) { restore(); fits(); }
    }
    if (!fitsPage && element.dataset.wordFloatingSpaces === 'true') {
      // Once a whole-paragraph fit fails, measure the original line layout
      // again before finding offsets for fragments with different geometry.
      delete element.dataset.wordFloatingSpaceSource;
      fitsPage = fits();
    }
    if (!fitsPage && settings.hideEmptyLines && content.children.length > 1 && hiddenEmptyLines < 2 &&
      element.getBoundingClientRect().height <= content.getBoundingClientRect().height && canHideEmptyPageLine(element)) {
      hiddenEmptyLines++; element.classList.add(hiddenEmptyLineClass); fits(); return;
    }
    if (fitsPage) return;
    // Explicit empty lines have flow even when the paragraph has no letters.
    if (element.matches('p,h1,h2,h3,h4,h5,h6,pre') && ((element.textContent?.length ?? 0) > 1 || element.querySelector('br'))) {
      const offsets = paragraphLineOffsets(element);
      const bodyHeight = content.getBoundingClientRect().height;
      const lineCount = offsets.length + 1;
      const keepTogether = keepShortWordParagraphTogether(element, lineCount, bodyHeight);
      const bounds = wrappingAnchors ? undefined : wordWidowSplitBounds(element, lineCount, bodyHeight, 1, () => offsets[0]);
      // Offset zero is the cut following line one.
      let low = bounds ? bounds.minimumHead - 1 : 0;
      let high = keepTogether ? -1 : (bounds ? bounds.maximumHead - 1 : offsets.length - 1);
      let best = 0;
      element.remove();
      while (low <= high) {
        const mid = (low + high) >> 1;
        const [a] = splitText(element, offsets[mid]); content.append(a);
        if (fits()) { best = offsets[mid]; low = mid + 1; } else high = mid - 1;
        a.remove();
      }
      if (best) {
        const [a, b] = splitText(element, best);
        content.append(a); newPage(); yield* append(b); return;
      }
      if (content.children.length) {
        // A rejected probe may have wrapped older text around this anchor.
        // Its exclusion leaves this sheet together with the paragraph.
        fits(); newPage();
        if (wordKeepParagraph(element) && element.dataset.wordKeepLines === 'true') element.style.marginTop = '0';
        yield* append(element); return;
      }
      content.append(element); scaleOversized(element); return;
    }
    if (content.children.length > 1) {
      // Removing a floating anchor also removes its exclusion from earlier
      // paragraphs. Reflow the old sheet, then resolve the picture against
      // its new paper instead of carrying its old relative CSS offset.
      element.remove(); fits(); newPage(); content.append(element); fits();
    }
    scaleOversized(element);
  }
  try {
    onProgress?.({ label: '문서의 글꼴을 준비하는 중…' });
    const sample = source.textContent || '한글';
    const serifFallback = /휴먼명조|한양신명조|함초롬바탕|바탕|Batang/i.test(html);
    const fontRequests: PrintFontRequest[] = [];
    for (const family of ['Noto Serif KR', 'Noto Sans KR', 'Studio Text', 'Studio Math', 'Studio Symbols', 'Studio Symbols 2', 'Studio Devanagari', 'Studio Emoji']) if (html.includes(family) || (family === 'Noto Serif KR' && serifFallback))
      for (const weight of [400, 700]) fontRequests.push({ font: `${weight} 12px "${family}"`, text: sample });
    if (wordLayout.paintHwpCellLines !== false) {
      const hwpFallbackText = [...source.querySelectorAll(
        `p[data-hwp-source-wrap*="scaledCellLine"],${hwpBodyFallbackSelector},${hwpCellFaceFallbackSelector},${hwpCellWidthFallbackSelector},p[${hwpNestedCellFaceAttribute}]`,
      )].map(paragraph => paragraph.textContent).join('');
      const boxFallbackText = [...source.querySelectorAll(hwpFloatingBoxProofSelector)].map(p => p.textContent).join('');
      if (boxFallbackText) for (const weight of [400, 700])
        fontRequests.push({ font: `${weight} 12px "${hwpFloatingBoxFontFallback}"`, text: boxFallbackText });
      if (hwpFallbackText) for (const weight of [400, 700])
        fontRequests.push({ font: `${weight} 12px "Noto Sans KR"`, text: hwpFallbackText });
    }
    if (source.querySelector('[data-symbol-font],[data-list-glyph]')) for (const family of ['Studio Symbols 2','Studio Symbols','Studio Math','Studio Text','Studio Devanagari','Studio Emoji'])
      fontRequests.push({ font: `12px "${family}"`, text: ' ' });
    if (wordLayout.hwpNativeExport && wordLayout.paintHwpCellLines === false) {
      // Native export can use the measured body fallback without adding
      // control characters to the document written from this DOM.
      const text = [...source.querySelectorAll(`${hwpBodyFallbackSelector},[${hwpNestedCellFaceAttribute}],${hwpCellWidthFallbackSelector}`)]
        .map(paragraph => paragraph.textContent).join('');
      if (text) fontRequests.push({ font: '400 12px "Noto Sans KR"', text });
    }
    if (await preparePrintFonts(host, source, noteBodies.values(), initialPage, fontRequests, signal))
      warnings.add('일부 글꼴을 불러오지 못해 대체 글꼴로 쪽 배치를 계산했습니다.');
    signal?.throwIfAborted();
    const images = [...source.querySelectorAll('img')]; let decoded = 0;
    if (images.length) onProgress?.({ label: '쪽 배치에 사용할 그림을 준비하는 중…', completed: 0, total: images.length, unit: '개' });
    await Promise.all(images.map(async img => {
      await img.decode().catch(() => {});
      if (!signal?.aborted) onProgress?.({ label: '쪽 배치에 사용할 그림을 준비하는 중…', completed: ++decoded, total: images.length, unit: '개' });
    }));
    const fills = new Set([...source.querySelectorAll(`td[${hwpCellRasterFillAttribute}]`)].flatMap(cell => {
      const fill = parseHwpCellRasterFill(cell.getAttribute(hwpCellRasterFillAttribute));
      return fill ? [hwpCellRasterPaintSource(fill)] : [];
    }));
    await Promise.all([...fills].map(async source => {
      const raster = new Image();
      raster.src = source;
      await raster.decode().catch(() => {});
    }));
    signal?.throwIfAborted();
    hwpSourceFlow = await prepareHwpSourceFlow(source);
    newPage();
    const children = [...source.children]; let processed = 0, yieldedAt = performance.now();
    const checkRows = wordLayout.interactive ? layoutCheckpoint(() => signal?.throwIfAborted()) : undefined;
    for (const child of children) {
      if ((onProgress && child.tagName === 'TABLE') ||
          ((onProgress || wordLayout.interactive) && performance.now() - yieldedAt >= 32)) {
        onProgress?.(child.tagName === 'TABLE' ? { label: '표의 쪽 배치를 계산하는 중…' } : { label: 'Word 쪽 배치를 계산하는 중…', completed: processed, total: children.length, unit: '항목' });
        await new Promise(resolve => setTimeout(resolve, 0)); signal?.throwIfAborted(); yieldedAt = performance.now();
      }
      for (const _ of append(child as HTMLElement)) {
        const pause = checkRows?.();
        if (pause) await pause;
      }
      processed++;
      onProgress?.({ label: 'Word 쪽 배치를 계산하는 중…', completed: processed, total: children.length, unit: '항목' });
    }
    refreshNotes();
    // Empty trailing paragraphs can spill onto an automatically created sheet.
    // Keep intentional page/section breaks and parity sheets, but do not print
    // an extra blank sheet produced only by those trailing line boxes.
    let last = host.querySelector<HTMLElement>('.print-page:last-child');
    while (pageNumber > 1 && last && !last.dataset.keepBlank && !last.querySelector('.print-content')?.textContent?.trim()
      && !last.querySelector('.print-content img,.print-content table,.print-content hr,.print-content [data-print-keep-trailing],.print-footnotes [data-document-note]')) {
      last.remove(); pageNumber--; last = host.querySelector<HTMLElement>('.print-page:last-child');
    }
    for (const field of host.querySelectorAll<HTMLElement>('[data-page-number="NUMPAGES"]')) field.textContent = String(pageNumber);
    for (const field of host.querySelectorAll<HTMLElement>('[data-page-number="PAGE"]')) field.textContent = field.closest<HTMLElement>('.print-page')!.dataset.page!;
    // Named @page rules affect output paper, not screen measurement. Publish
    // them once so every new sheet does not invalidate the live editor's styles.
    style.textContent = [...new Set(pageStyles)].join('\n');
    await layoutWordCellSpacing(host);
    await layoutHwpSourceAdvances(host, wordLayout.interactive ? layoutCheckpoint(() => signal?.throwIfAborted()) : undefined);
    if (wordLayout.paintHwpCellLines !== false) {
      layoutHwpSourceCellLines(host);
      if (hwpSourceFlow?.missingCompactCellFace) {
        const tables = new Set([...host.querySelectorAll(hwpCellFaceFallbackSelector)].map(p => p.closest('table')));
        for (const table of tables) if (table)
          layoutHwpCompactCellFaces(table, true, hwpSourceFlow.missingCompactHeaderFace, 'paint');
      }
      layoutHwpAuthoredBodyLines(host);
      for (const box of host.querySelectorAll<HTMLElement>('.print-content > [data-document-textbox][data-textbox-position="floating"]'))
        layoutHwpFloatingTextBoxLines(box);
      layoutHwpFittedCellLineEnds(host);
    }
    signal?.throwIfAborted();
    positionPaperImagesInPrintColumns(host);
    if (wordLayout.nativeDesktop) {
      for (const group of host.querySelectorAll<HTMLElement>('.print-content > [data-column-count]:only-child')) {
        const height = group.style.getPropertyValue('column-height');
        if (!(Number(group.dataset.columnCount) > 1) || !height) continue;
        // Chromium can fragment a trimmed column group again during physical
        // printing. Use the measured column height when it owns the body.
        group.dataset.nativePrintColumnHeight = 'true';
        group.style.setProperty('--native-print-column-height', height);
      }
    }
    return { element: host, pages: pageNumber, pageSettings, warnings: [...warnings], dispose: () => host.remove() };
  } catch (error) { host.remove(); throw error; }
}
