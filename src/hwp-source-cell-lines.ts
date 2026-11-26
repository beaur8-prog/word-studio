import type { LayoutCheckpoint } from './editor-layout-task';
import { documentScale } from './document-scale';
import { characterScaleProperty, parseCharacterScale } from './character-scale';
import { authoredCellParagraphScale, cellCharacterScaleDimensions } from './cell-character-scale';
import { applyHwpCellTerminal, hwpCellTerminalLayout, unwrapHwpCellPrefixes } from './hwp-cell-terminal';
import { measuredHwpInkAtLineOffsets, measuredHwpLineInk, retainedTextBands } from './hwp-line-ink';
import { hwpSourceAdvanceParagraph, matchesStoredWrap, plainHwpSourceParagraph } from './hwp-text-advance-layout';
import { hwpCellFaceFallbackSelector, hwpCellWidthFallbackSelector, hwpCompactCellFace, hwpCompactHeaderFace, hwpScaledCellFace, hwpSourceWrapAttribute, hwpWrapFingerprint, normalizeHwpWrapText, parseHwpSourceWrap, type HwpSourceWrap } from './hwp-text-advances';
import { layoutHwpCellLines } from './hwp-cell-line-metrics';
import { paragraphLineOffsets, paragraphTextWalker } from './paragraph-fragments';
import { hwpCellAdvance, parseHwpCellLine } from './hwp-cell-line';
import { rowSplitPolicy } from './row-split-policy';

const graphemes = new Intl.Segmenter(undefined, { granularity: 'grapheme' });
type Boundary = { node: Node; at: number; offset: number };

export function sourceLineBoundaries(paragraph: HTMLElement, ends: readonly number[]) {
  const boundaries: Boundary[] = [{ node: paragraph, at: 0, offset: 0 }];
  const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT);
  let offset = 0, normalized = 0;
  while (boundaries.length <= ends.length && walker.nextNode()) {
    const leaf = walker.currentNode as Text;
    for (const { segment, index } of graphemes.segment(leaf.data)) {
      const length = normalizeHwpWrapText(segment).length;
      if (!length) continue;
      if (normalized === ends[boundaries.length - 1]) boundaries.push({ node: leaf, at: index, offset: offset + index });
      normalized += length;
    }
    offset += leaf.length;
  }
  if (boundaries.length !== ends.length + 1) return;
  boundaries.push({ node: paragraph, at: paragraph.childNodes.length, offset: paragraph.textContent?.length ?? 0 });
  return boundaries;
}

export function lineOpportunities(paragraph: HTMLElement, boundaries: readonly Boundary[], range: Range) {
  const content = document.createDocumentFragment();
  for (let index = 0; index + 1 < boundaries.length; index++) {
    const start = boundaries[index], end = boundaries[index + 1];
    range.setStart(start.node, start.at);
    range.setEnd(end.node, end.at);
    const line = document.createElement('span');
    const contents = range.cloneContents();
    let ancestor = range.commonAncestorContainer;
    if (ancestor instanceof Text) ancestor = ancestor.parentNode!;
    // A whole line inside one styled span would otherwise lose that span
    // when Range.cloneContents omits the common ancestor.
    while (ancestor !== paragraph) {
      const wrapper = ancestor.cloneNode(false);
      wrapper.appendChild(contents);
      contents.append(wrapper);
      ancestor = ancestor.parentNode!;
    }
    line.append(contents);
    // White space in a source line remains literal. WBR provides the
    // opportunity between lines while the available width still places them.
    for (const inline of [line, ...line.querySelectorAll<HTMLElement>('*')]) inline.style.whiteSpace = 'pre';
    if (index) content.append(document.createElement('wbr'));
    content.append(line);
  }
  return content;
}

function geometry(element: Element) {
  const { left, top, width, height } = element.getBoundingClientRect();
  return [left, top, width, height].join(':');
}

const groupAttribute = 'data-hwp-print-wrap-group';
const excludedInline = 'a,[id],[data-docx-bookmark],[data-equation],mark[data-comment-ids]';

export function markHwpSourceCellLineGroups(source: HTMLElement) {
  let next = 0;
  for (const paragraph of source.querySelectorAll<HTMLElement>(`p[${hwpSourceWrapAttribute}]`)) {
    if (paragraph.parentElement?.tagName === 'TD' && hwpSourceAdvanceParagraph(paragraph) &&
        !paragraph.querySelector(excludedInline))
      paragraph.setAttribute(groupAttribute, String(++next));
  }
}

let inkCanvas: CanvasRenderingContext2D | null | undefined;

function inkContext() {
  if (inkCanvas === undefined) {
    inkCanvas = document.createElement('canvas').getContext('2d');
    if (inkCanvas) {
      inkCanvas.direction = 'ltr';
      inkCanvas.textAlign = 'left';
    }
  }
  return inkCanvas;
}

export function horizontalInkFits(paragraph: HTMLElement, range: Range, bounds: Pick<DOMRect, 'left' | 'right'>,
  scale = 1, band?: Pick<DOMRect, 'top' | 'bottom'>): boolean {
  const canvas = inkContext();
  if (!canvas || scale === 1 && getComputedStyle(paragraph).transform !== 'none') return false;
  const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT);
  // The browser's Range gives the placed glyph origin. The finer canvas
  // outline distinguishes an actual side bearing from a full glyph box.
  const detail = 16;
  while (walker.nextNode()) {
    const leaf = walker.currentNode as Text;
    if (!leaf.data.trim()) continue;
    const css = getComputedStyle(leaf.parentElement!);
    const size = parseFloat(css.fontSize);
    if (!(size > 0) || css.direction !== 'ltr' || css.writingMode !== 'horizontal-tb' ||
        css.transform !== 'none' || css.fontVariationSettings !== 'normal' || css.textTransform !== 'none') return false;
    canvas.font = `${css.fontStyle} ${css.fontWeight} ${size * detail}px ${css.fontFamily}`;
    for (const { segment, index } of graphemes.segment(leaf.data)) {
      if (!segment.trim()) continue;
      range.setStart(leaf, index);
      range.setEnd(leaf, index + segment.length);
      const boxes = [...range.getClientRects()].filter(box => box.width > 0 && box.height > 0);
      if (boxes.length !== 1) return false;
      const [box] = boxes;
      const center = (box.top + box.bottom) / 2;
      if (band && (center < band.top || center >= band.bottom)) continue;
      const ink = canvas.measureText(segment);
      const left = box.left - scale * (ink.actualBoundingBoxLeft + 1) / detail;
      const right = box.left + scale * (ink.actualBoundingBoxRight + 1) / detail;
      if (!(left >= bounds.left && right <= bounds.right)) return false;
    }
  }
  return true;
}

function fixedSingleLineTable(table: HTMLTableElement, allowRelatedMember = false) {
  // Native HWP grids can have real fixed columns before DOCX gives the
  // model an explicit Word layout attribute.
  const related = allowRelatedMember && table.dataset.hwpRelatedInlineMember === 'true' &&
    table.parentElement?.matches('td') && table.parentElement?.closest('table')?.hasAttribute('data-hwp-related-print-group');
  return table.dataset.wordTableLayout !== 'autofit' && table.dataset.wordTableWidthAuto !== 'true' &&
    (related || !table.parentElement?.closest('td,th,[data-column-count],[data-document-textbox]')) &&
    !!table.querySelector(':scope > colgroup > col') && getComputedStyle(table).tableLayout === 'fixed';
}

function matchesRoundedCellMeasure(cell: HTMLTableCellElement, cellWidth: NonNullable<HwpSourceWrap['cellWidth']>,
  cellBox: DOMRect, paragraphBox: DOMRect, allowRelatedMember = false) {
  const table = cell.closest('table');
  if (!table || !fixedSingleLineTable(table, allowRelatedMember) || getComputedStyle(table).borderCollapse !== 'collapse') return false;
  const row = cell.parentElement as HTMLTableRowElement;
  const columnIndex = [...row.cells].slice(0, cell.cellIndex).reduce((sum, sibling) => sum + sibling.colSpan, 0);
  const column = table.querySelectorAll<HTMLTableColElement>(':scope > colgroup > col')[columnIndex];
  const roundedWidth = Math.round(cellWidth.outer);
  if (!column || column.span !== 1 || column.style.width !== `${roundedWidth}px`) return false;
  // HWP stores the width before borders. The imported fixed grid rounds the
  // column, and collapsed borders occupy half their width inside the cell.
  const css = getComputedStyle(cell);
  const horizontalBorders = parseFloat(css.borderLeftWidth) + parseFloat(css.borderRightWidth);
  const roundedText = cellWidth.text + roundedWidth - cellWidth.outer - horizontalBorders / 2;
  const precision = .05 + 1 / 15 + 2 / 64;
  return Math.abs(cellBox.width - roundedWidth) <= precision && Math.abs(paragraphBox.width - roundedText) <= precision;
}

function inkFitsRoundedCellLines(paragraph: HTMLElement, cell: HTMLTableCellElement, bounds: DOMRect, range: Range, lineCount: number) {
  const lines = [...paragraph.children].filter((child): child is HTMLSpanElement => child instanceof HTMLSpanElement);
  const box = paragraph.getBoundingClientRect();
  const css = getComputedStyle(cell);
  const left = bounds.left + parseFloat(css.borderLeftWidth) / 2;
  const right = bounds.right - parseFloat(css.borderRightWidth) / 2;
  const pitch = parseFloat(getComputedStyle(paragraph).lineHeight);
  if (lines.length !== lineCount || !(pitch > 0) || !horizontalInkFits(paragraph, range, { left, right })) return false;
  const cellBottom = bounds.bottom - parseFloat(css.borderBottomWidth) / 2;
  let bottom = Math.max(bounds.top + parseFloat(css.borderTopWidth) / 2, box.top);
  for (const line of lines) {
    const ink = measuredHwpLineInk(line, 1, pitch, true);
    if (!ink || ink.top < bottom || ink.bottom > Math.min(cellBottom, box.bottom)) return false;
    bottom = ink.bottom;
  }
  return true;
}

/** Keep the original family first, and retain the trial only when the
 * authored horizontal scale restores the complete row at the original width. */
export function layoutHwpScaledCellLines(table: HTMLTableElement) {
  if (table.rows.length !== 1 || table.dataset.wordTableWidthAuto === 'true' ||
      !table.querySelector(`p[${hwpSourceWrapAttribute}*="scaledCellLine"]`) ||
      table.parentElement?.closest('td,th,[data-column-count],[data-document-textbox]') ||
      getComputedStyle(table).tableLayout !== 'fixed') return;
  const row = table.rows[0];
  const grid = [...table.querySelectorAll<HTMLTableColElement>(':scope > colgroup > col')];
  // The DOCX layout preference can say autofit while its explicit grid
  // provides the actual fixed columns used by this print copy.
  if (!row.cells.length || grid.length !== row.cells.length ||
      grid.some(column => !/^\d+(?:\.\d+)?px$/.test(column.style.width) || !(parseFloat(column.style.width) > 0))) return;
  const range = document.createRange();
  const fallbackStack = (name: string) => {
    const face = name.trim().replace(/^["']|["']$/g, '');
    if (face === hwpScaledCellFace) return `${name}, "Noto Sans KR"`;
    // The native bridge can quote a combined family as one CSS name.
    // Keep that source value and try its authored face before the fallback.
    if (face === `${hwpScaledCellFace},한컴바탕`) return `${name}, "${hwpScaledCellFace}", "Noto Sans KR"`;
  };
  for (const cell of row.cells) {
    if (cell.tagName !== 'TD' || cell.rowSpan !== 1 || cell.colSpan !== 1 || cell.children.length !== 1) continue;
    const paragraph = cell.firstElementChild;
    if (!(paragraph instanceof HTMLParagraphElement)) continue;
    const reference = parseHwpSourceWrap(paragraph.dataset.hwpSourceWrap);
    const witness = reference?.scaledCellLine;
    if (!witness || !reference.cellWidth) continue;
    const run = paragraph.firstElementChild, text = paragraph.textContent || '';
    if (!(run instanceof HTMLSpanElement) || paragraph.childNodes.length !== 1 || run.childNodes.length !== 1 ||
        run.firstChild?.nodeType !== Node.TEXT_NODE || hwpWrapFingerprint(text) !== reference.fingerprint ||
        normalizeHwpWrapText(text).length !== reference.ends[0] ||
        parseCharacterScale(run.style.getPropertyValue(characterScaleProperty)) !== witness.scale)
      continue;
    const beforeParagraph = paragraph.style.fontFamily, beforeRun = run.style.fontFamily;
    const fitParagraph = fallbackStack(beforeParagraph), fitRun = fallbackStack(beforeRun);
    if (!fitParagraph || !fitRun) continue;
    const css = getComputedStyle(paragraph), runCss = getComputedStyle(run);
    const transform = css.transform === 'none' ? null : new DOMMatrix(css.transform);
    const scale = witness.scale / 100;
    if (!transform || !transform.is2D || Math.abs(transform.a - scale) > .0001 ||
        transform.b || transform.c || transform.d !== 1 || transform.e || transform.f ||
        css.textIndent !== '0px' || Math.abs(parseFloat(runCss.fontSize) - witness.size) > .05 ||
        !['normal', '0px'].includes(runCss.letterSpacing) || runCss.textDecorationLine !== 'none' || runCss.textShadow !== 'none' ||
        parseFloat(runCss.getPropertyValue('-webkit-text-stroke-width')) > 0) continue;
    const bounds = cell.getBoundingClientRect(), box = paragraph.getBoundingClientRect();
    // The reader reports a tenth of a pixel, and the editor gives the grid
    // integer pixel columns on reopen. Accept only their rounding interval.
    const cellCss = getComputedStyle(cell), tolerance = parseFloat(cellCss.borderLeftWidth) + parseFloat(cellCss.borderRightWidth) + .55;
    if (Math.abs(bounds.width - reference.cellWidth.outer) > tolerance ||
        Math.abs(box.width - reference.cellWidth.text) > tolerance ||
        bounds.height <= witness.rowHeight + .5 || !paragraphLineOffsets(paragraph).length) continue;
    paragraph.style.fontFamily = fitParagraph;
    run.style.fontFamily = fitRun;
    const fitted = cell.getBoundingClientRect();
    range.selectNodeContents(run);
    const glyphs = [...range.getClientRects()].filter(rect => rect.width > 0 && rect.height > 0);
    if (Math.abs(fitted.height - witness.rowHeight) > .5 || fitted.width !== bounds.width ||
        paragraphLineOffsets(paragraph).length || !glyphs.length ||
        glyphs.some(rect => rect.top < fitted.top - .5 || rect.bottom > fitted.bottom + .5) ||
        !horizontalInkFits(paragraph, range, fitted, scale)) {
      paragraph.style.fontFamily = beforeParagraph;
      run.style.fontFamily = beforeRun;
    }
  }
}

function singleLineCell(p: HTMLElement, scale: number, printedHeader = false) {
  const proof = parseHwpSourceWrap(p.dataset.hwpSourceWrap);
  const cell = p.parentElement;
  if (!proof?.cellWidth || !(cell instanceof HTMLTableCellElement) ||
      !(proof.singleCellLine && cell.tagName === 'TD' && cell.children.length === 2 ||
        printedHeader && proof.singleHeaderLine && cell.tagName === 'TH' && cell.children.length === 1) ||
      cell.lastElementChild !== p || cell.colSpan !== 1 || cell.rowSpan !== 1 ||
      p.querySelector(':not(span,strong,b,em,i),[id],[data-docx-tab],[data-document-field],[data-symbol-font],[data-word-overlap],[data-flow-widget]'))
    return;

  const literal = p.textContent || '';
  const box = p.getBoundingClientRect(), bounds = cell.getBoundingClientRect(), css = getComputedStyle(cell);
  const tolerance = parseFloat(css.borderLeftWidth) + parseFloat(css.borderRightWidth) + .16;
  if (proof.fingerprint !== hwpWrapFingerprint(literal) || proof.ends[0] !== normalizeHwpWrapText(literal).length ||
      Math.abs(box.width / scale - proof.cellWidth.text) > tolerance ||
      Math.abs(bounds.width / scale - proof.cellWidth.outer) > tolerance) return;

  const inline = [p, ...p.querySelectorAll<HTMLElement>('span,strong,b,em,i')];
  if (inline.some(node => {
    const style = getComputedStyle(node);
    return style.textShadow !== 'none' || parseFloat(style.getPropertyValue('-webkit-text-stroke-width')) > 0 ||
      style.textDecorationLine !== 'none';
  })) return;
  return { cell, box, bounds, inline };
}

const singleLineSizing = ['boxSizing', 'paddingLeft', 'paddingRight', 'paddingTop', 'paddingBottom',
  'borderLeftWidth', 'borderRightWidth', 'borderTopWidth', 'borderBottomWidth', 'width', 'minHeight',
  'lineHeight', 'textAlign', 'textIndent', 'wordBreak', 'overflowWrap', 'lineBreak'] as const;
const singleLineTypography = ['fontFamily', 'fontSize', 'fontWeight', 'fontStyle', 'fontStretch', 'fontSynthesis',
  'fontVariantCaps', 'fontVariantLigatures', 'fontFeatureSettings', 'fontVariationSettings', 'fontKerning',
  'letterSpacing', 'wordSpacing', 'verticalAlign', 'lineHeight', 'direction', 'writingMode', 'textTransform'] as const;
export const hwpEditorSingleLine = 'data-hwp-editor-single-cell-line';

/** Measure in unscaled CSS pixels without editing a ProseMirror-owned node. */
export function editorHwpSingleCellLine(p: HTMLElement, scale: number): number | undefined {
  const candidate = singleLineCell(p, scale);
  if (!candidate) return;
  const table = candidate.cell.closest('table');
  const active = p.hasAttribute(hwpEditorSingleLine);
  const css = getComputedStyle(p);
  if (!table || !fixedSingleLineTable(table) || !active && (!paragraphLineOffsets(p).length || css.textIndent !== '0px'))
    return;
  const { box, bounds, inline } = candidate;
  const probe = p.cloneNode(true) as HTMLElement;
  const runs = [probe, ...probe.querySelectorAll<HTMLElement>('span,strong,b,em,i')];
  for (const [index, node] of inline.entries()) {
    if (node.hasAttribute('data-hwp-cell-prefix')) continue;
    const style = getComputedStyle(node);
    if (style.display !== (index ? 'inline' : 'block') || style.transform !== 'none' ||
        style.fontStretch !== '100%' || style.fontVariantCaps !== 'normal' || style.position !== 'static') return;
    for (const property of singleLineTypography) runs[index].style[property] = style[property];
    runs[index].style.whiteSpace = 'pre';
  }
  // When the edit is undone the old two-line prefix can still be in the DOM.
  // Its layout decoration applies only to preceding lines, not this probe.
  unwrapHwpCellPrefixes(probe);
  for (const property of singleLineSizing) probe.style[property] = css[property];
  probe.style.textIndent = '0px';
  probe.style.position = 'fixed';
  probe.style.left = '-100000px';
  probe.style.top = '0';
  probe.style.margin = '0';
  probe.style.visibility = 'hidden';
  probe.setAttribute('aria-hidden', 'true');
  document.body.append(probe);
  try {
    const measurement = probe.getBoundingClientRect();
    if (Math.abs(measurement.width - box.width / scale) > 1 / 64 ||
        !active && measurement.height >= box.height / scale || paragraphLineOffsets(probe).length) return;

    const left = measurement.left + (bounds.left - box.left) / scale;
    const right = measurement.left + (bounds.right - box.left) / scale;
    const range = document.createRange();
    range.selectNodeContents(probe);
    const advance = range.getBoundingClientRect();
    let indent = 0;
    if (advance.left < left || advance.right > right) {
      // The browser start-aligns an overflowing centered line. Retain its
      // authored center when the advance and both caret edges fit the cell.
      if (css.textAlign !== 'center') return;
      indent = (measurement.left + measurement.right - advance.left - advance.right) / 2;
      probe.style.textIndent = `${indent}px`;
    }
    const glyphs = [...range.getClientRects()].filter(rect => rect.width > 0 && rect.height > 0);
    if (!glyphs.length || glyphs.some(rect => rect.left < left || rect.right > right ||
        rect.top - measurement.top < (bounds.top - box.top) / scale - .5 ||
        rect.bottom - measurement.top > (bounds.bottom - box.top) / scale + .5) ||
        !horizontalInkFits(probe, range, { left, right })) return;
    return indent;
  } finally { probe.remove(); }
}

/** Restore an authored short cell line only where the real glyph ink fits
 * at the unchanged measured column width. Changing the print copy leaves
 * all literal whitespace and the editable model in their original order. */
export function layoutHwpSingleCellLines(table: HTMLTableElement) {
  const candidates = table.querySelectorAll<HTMLParagraphElement>(`td > p[${hwpSourceWrapAttribute}*="singleCellLine"],th > p[${hwpSourceWrapAttribute}*="singleHeaderLine"]`);
  if (!candidates.length || !fixedSingleLineTable(table)) return;
  const range = document.createRange();
  for (const p of candidates) {
    const candidate = singleLineCell(p, 1, true);
    if (!candidate || candidate.cell.closest('table') !== table || !paragraphLineOffsets(p).length) continue;
    const { cell, box, bounds, inline } = candidate;
    const previous = inline.map(node => node.style.whiteSpace);
    inline.forEach(node => { node.style.whiteSpace = 'pre'; });
    const fresh = p.getBoundingClientRect(), cellBox = cell.getBoundingClientRect();
    range.selectNodeContents(p);
    const glyphs = [...range.getClientRects()].filter(rect => rect.width > 0 && rect.height > 0);
    if (fresh.width !== box.width || fresh.height >= box.height || cellBox.width !== bounds.width ||
        paragraphLineOffsets(p).length || !glyphs.length ||
        glyphs.some(rect => rect.top < cellBox.top - .5 || rect.bottom > cellBox.bottom + .5) ||
        !horizontalInkFits(p, range, cellBox)) {
      inline.forEach((node, index) => { node.style.whiteSpace = previous[index]; });
    }
  }
}

// A compact last line can keep the same strut while a newly moved letter
// descends below it. Measure the existing line's baseline and visible tail.
function terminalInkBottom(paragraph: HTMLElement, range: Range): number | null {
  const canvas = inkContext();
  if (!canvas) return null;
  const scale = documentScale(paragraph);
  const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT);
  const runs: { text: Text; box: string; descent: number; probe: HTMLSpanElement }[] = [];
  const boxOf = (text: Text) => {
    range.selectNodeContents(text);
    const boxes = [...range.getClientRects()].filter(rect => rect.width > 0 && rect.height > 0);
    if (boxes.length === 1) {
      const { x, y, width, height } = boxes[0];
      return [x, y, width, height].join(':');
    }
  };
  while (walker.nextNode()) {
    const text = walker.currentNode as Text;
    if (!text.data.trim() || text.parentElement?.closest('[data-hwp-cell-prefix]')) continue;
    const style = getComputedStyle(text.parentElement!);
    const size = parseFloat(style.fontSize), box = boxOf(text);
    if (!box || !(size > 0) || style.direction !== 'ltr' || style.writingMode !== 'horizontal-tb' ||
        style.transform !== 'none' || style.fontVariationSettings !== 'normal' || style.fontStretch !== '100%' ||
        style.fontVariantCaps !== 'normal' || style.textTransform !== 'none' || style.textShadow !== 'none' ||
        style.textDecorationLine !== 'none' || parseFloat(style.getPropertyValue('-webkit-text-stroke-width')) > 0)
      return null;
    canvas.font = `${style.fontStyle} ${style.fontWeight} ${size * 16}px ${style.fontFamily}`;
    const descent = (canvas.measureText(text.data).actualBoundingBoxDescent + 1) * scale / 16;
    if (!Number.isFinite(descent)) return null;
    const probe = document.createElement('span');
    probe.style.cssText = 'display:inline-block;vertical-align:baseline;width:0;height:0;line-height:0;font-size:0;margin:0;padding:0;border:0';
    runs.push({ text, box, descent, probe });
  }
  if (!runs.length) return null;
  const before = geometry(paragraph);
  try {
    runs.forEach(({ text, probe }) => text.after(probe));
    if (geometry(paragraph) !== before || runs.some(({ text, box }) => boxOf(text) !== box)) return null;
    return Math.max(...runs.map(({ probe, descent }) => probe.getBoundingClientRect().bottom + descent));
  } finally {
    runs.forEach(({ probe }) => probe.remove());
  }
}

function retainedFirstLineTop(paragraph: HTMLElement, range: Range, body: DOMRect, paragraphBox: DOMRect): number {
  const normalTop = paragraphBox.top - .5;
  const previous = paragraph.previousElementSibling;
  if (!(previous instanceof HTMLParagraphElement) || !previous.textContent?.trim() ||
      previous.matches('[data-paragraph-border-box],[data-word-paragraph-bottom-border]') ||
      previous.querySelector('img,svg,table,mark,[data-note-reference],[data-document-field],[data-word-overlap],' +
        '[style*="background"],[style*="box-shadow"],[style*="outline"],[style*="border-"]')) return normalTop;
  const style = getComputedStyle(previous);
  if (style.backgroundColor !== 'rgba(0, 0, 0, 0)' || style.boxShadow !== 'none' ||
      style.outlineStyle !== 'none' || [style.borderTopWidth, style.borderBottomWidth].some(width => parseFloat(width) > 0))
    return normalTop;
  range.selectNodeContents(paragraph);
  const firstTop = Math.min(...[...range.getClientRects()].filter(rect => rect.width > 0 && rect.height > 0).map(rect => rect.top));
  if (firstTop >= normalTop) return normalTop;
  range.selectNodeContents(previous);
  const previousBottom = Math.max(...[...range.getClientRects()].filter(rect => rect.width > 0 && rect.height > 0).map(rect => rect.bottom));
  return Math.min(normalTop, Math.max(firstTop, body.top, previousBottom));
}

/** Preserve the column's ordinary justification when just one interior
 * line end is different. Check ink on both affected lines and retain the
 * full paragraph's geometry. */
function tryLocalColumnBoundary(paragraph: HTMLElement, text: string, sourceBreaks: number[], previous: number[],
  bounds: Pick<DOMRect, 'left' | 'right'>, range: Range) {
  const changed = previous.flatMap((offset, index) =>
    normalizeHwpWrapText(text.slice(0, offset)).length === sourceBreaks[index] ? [] : [index]);
  const index = changed[0], box = paragraph.getBoundingClientRect(), pitch = parseFloat(getComputedStyle(paragraph).lineHeight);
  if (changed.length !== 1 || index < 1 || index + 1 >= sourceBreaks.length ||
      !(pitch > 0 && Math.abs(box.height - (sourceBreaks.length + 1) * pitch) <= 2 / 64) ||
      paragraph.querySelector('[style*="white-space"]')) return;
  const cuts = sourceLineBoundaries(paragraph, [sourceBreaks[index], sourceBreaks[index + 1]]);
  if (!cuts) return;
  const content = lineOpportunities(paragraph, cuts, range);
  const lines = [...content.children].filter((node): node is HTMLSpanElement => node instanceof HTMLSpanElement);
  if (lines.length !== 3) return;
  for (const [at, line] of lines.entries()) {
    line.style.display = 'contents';
    if (at !== 1) for (const node of [line, ...line.querySelectorAll<HTMLElement>('*')]) node.style.removeProperty('white-space');
  }
  const original = [...paragraph.childNodes], originalGeometry = geometry(paragraph);
  const restore = () => paragraph.replaceChildren(...original);
  const affected = { top: box.top + index * pitch, bottom: box.top + (index + 2) * pitch };
  // A tall run may already paint above its nominal line box. Preserve the
  // original vertical envelope there, but keep the neighbouring lines clear.
  const bands = () => {
    range.selectNodeContents(paragraph);
    const lines = Array.from({ length: 4 }, () => ({ top: Infinity, bottom: -Infinity }));
    for (const rect of range.getClientRects()) {
      if (rect.width <= 0 || rect.height <= 0) continue;
      const slot = Math.floor(((rect.top + rect.bottom) / 2 - affected.top) / pitch) + 1;
      const line = lines[slot];
      if (!line) continue;
      line.top = Math.min(line.top, rect.top);
      line.bottom = Math.max(line.bottom, rect.bottom);
    }
    return lines;
  };
  const before = bands();
  if (before.some(line => !Number.isFinite(line.top))) return;
  paragraph.replaceChildren(content);
  const placed = paragraphLineOffsets(paragraph);
  const after = bands();
  const neighboursStable = [0, 3].every(at => before[at].top === after[at].top && before[at].bottom === after[at].bottom);
  const localInkStaysClear = [1, 2].every(at => {
    const start = affected.top + (at - 1) * pitch;
    const top = Math.max(after[at - 1].bottom, Math.min(start - .5, before[at].top));
    const bottom = Math.min(after[at + 1].top, Math.max(start + pitch + .5, before[at].bottom));
    return Number.isFinite(after[at].top) && after[at].top >= top && after[at].bottom <= bottom;
  });
  if (paragraph.textContent !== text || geometry(paragraph) !== originalGeometry || paragraph.getClientRects().length !== 1 ||
      placed.length !== sourceBreaks.length ||
      placed.some((offset, at) => normalizeHwpWrapText(text.slice(0, offset)).length !== sourceBreaks[at]) ||
      !neighboursStable || !localInkStaysClear ||
      !horizontalInkFits(paragraph, range, bounds, 1, affected)) {
    restore(); return;
  }
  return restore;
}

/** Keep the existing two vertical slots for a mixed-size column paragraph.
 * The exact source paint witnesses the taller first line and the original cut. */
function tryMixedColumnPair(paragraph: HTMLElement, text: string, cut: number, pitches: [number, number],
  bounds: Pick<DOMRect, 'left' | 'right'>, range: Range) {
  const box = paragraph.getBoundingClientRect(), css = getComputedStyle(paragraph);
  const [first, second] = pitches;
  const font = parseFloat(css.fontSize), exact = parseFloat(paragraph.style.getPropertyValue('--exact-line-height'));
  if (text.length > 250 || css.lineHeight !== '0px' || !(parseFloat(css.textIndent) < 0) ||
      !Number.isFinite(exact) || Math.abs(exact - first) > .16 || Math.abs(box.height - first - second) > .16)
    return;
  const boundaries = sourceLineBoundaries(paragraph, [cut]);
  if (!boundaries) return;
  const projected = lineOpportunities(paragraph, boundaries, range);
  const lines = [...projected.children].filter((element): element is HTMLSpanElement => element instanceof HTMLSpanElement);
  if (lines.length !== 2) return;
  Object.assign(lines[0].style, { display: 'block', height: first + 'px', textAlignLast: 'justify', whiteSpace: 'break-spaces' });
  for (const inline of lines[0].querySelectorAll<HTMLElement>('*')) inline.style.whiteSpace = 'break-spaces';
  Object.assign(lines[1].style, { display: 'block', textIndent: '0px' });
  const original = [...paragraph.childNodes], before = geometry(paragraph);
  const restore = () => paragraph.replaceChildren(...original);
  paragraph.replaceChildren(...lines);
  lines[1].style.height = box.height - lines[0].getBoundingClientRect().height + 'px';
  const offsets = paragraphLineOffsets(paragraph);
  // Inspect both slots together. A justification probe inside one slot can
  // alter word spacing, so use the measured baseline with that spacing intact.
  const ink = measuredHwpInkAtLineOffsets(paragraph, offsets, font, true);
  if (paragraph.textContent !== text || geometry(paragraph) !== before || offsets.length !== 1 ||
      offsets[0] !== boundaries[1].offset || !horizontalInkFits(paragraph, range, bounds) || ink?.length !== 2 ||
      !lines.every((line, index) => {
        const slot = line.getBoundingClientRect(), visible = ink[index];
        return Math.abs(slot.height - pitches[index]) <= .16 && slot.top >= box.top && slot.bottom <= box.bottom &&
          visible.top >= slot.top && visible.bottom <= slot.bottom;
      })) {
    restore();
    return;
  }
  return restore;
}

function tryWrap(paragraph: HTMLElement, reference: HwpSourceWrap, sourceBreaks: number[], range: Range,
  fragment = false, explicit = false, preserveMatchingBreaks = false) {
  const terminal = !fragment && !!reference.cellWidth && paragraph.dataset.hwpCellTerminal === 'true';
  const column = reference.columnWidth !== undefined;
  const columnCell = reference.columnCellHeight !== undefined;
  const columnPitch = reference.columnCellPitch;
  if (paragraph.querySelector(excludedInline) ||
      (column || columnCell ? paragraph.querySelector(':not(span,strong,b,em,i,u),[data-word-overlap],[data-flow-widget]') :
        !hwpSourceAdvanceParagraph(paragraph, fragment, terminal))) return false;
  if (terminal && (paragraph.hasAttribute('data-print-soft-end') ||
      paragraph.querySelectorAll('[data-hwp-cell-prefix]').length !== 1)) return false;
  const text = paragraph.textContent || '';
  const parent = paragraph.parentElement;
  const sheet = paragraph.closest('.print-page');
  if (!parent || !sheet) return false;
  const parentBox = parent.getBoundingClientRect(), paragraphBox = paragraph.getBoundingClientRect();
  const elements = [paragraph, parent, sheet];
  let columnBounds: { left: number; right: number; top: number; bottom: number } | undefined;
  let roundedCellMeasure = false;
  if (parent instanceof HTMLTableCellElement && parent.tagName === 'TD' && parent.rowSpan === 1 && parent.colSpan === 1 && reference.cellWidth) {
    const css = getComputedStyle(parent);
    const borders = (parseFloat(css.borderLeftWidth) || 0) + (parseFloat(css.borderRightWidth) || 0);
    const tolerance = borders + .05 + 1 / 15;
    const row = parent.parentElement!;
    if (columnCell) {
      const table = row.closest('table'), style = getComputedStyle(paragraph);
      const indent = Math.min(0, parseFloat(style.textIndent) || 0);
      const cellBorderTolerance = tolerance + 2 / 64;
      // The imported grid rounds each column to whole CSS pixels. Use the
      // same delta for its measured text rather than widening the tolerance.
      const roundedWidth = Math.round(reference.cellWidth.outer);
      const projectedText = reference.cellWidth.text + roundedWidth - reference.cellWidth.outer;
      if (fragment || terminal && columnPitch === undefined || !table || table.rows.length !== 1 || row.children.length !== 1 ||
          !table.closest('[data-column-count="2"]') ||
          table.parentElement?.closest('td,th,[data-document-note],[data-document-textbox]') ||
          table.dataset.wordTableLayout === 'autofit' || table.dataset.wordTableWidthAuto === 'true' ||
          getComputedStyle(table).tableLayout !== 'fixed' || style.textAlign !== 'justify' ||
          style.writingMode !== 'horizontal-tb' || style.direction !== 'ltr' || style.transform !== 'none' ||
          Math.abs(roundedWidth - parentBox.width) > cellBorderTolerance ||
          Math.abs(projectedText - (paragraphBox.width - indent)) > cellBorderTolerance) return false;
      const heightTolerance = .1 + 1 / 15 + 2 / 64;
      if (columnPitch === undefined) {
        if (Math.abs(reference.columnCellHeight! - parentBox.height) > heightTolerance) return false;
      } else {
        const body = paragraph.closest('.print-content')?.getBoundingClientRect();
        const advance = terminal ? Number(paragraph.dataset.hwpCellAdvance) : parseFloat(style.lineHeight);
        const borders = parseFloat(css.borderTopWidth) + parseFloat(css.borderBottomWidth);
        if (!body || parent.children.length !== (reference.columnCellStories ?? 1) ||
            (reference.columnCellStories && (terminal || parent.lastElementChild === paragraph || indent ||
              [...parent.children].some(child => child.tagName !== 'P'))) || !Number.isFinite(advance) ||
            Math.abs(columnPitch - advance) > .05 + 1 / 15 + 2 / 64 ||
            Math.abs(reference.columnCellHeight! - parentBox.height) > heightTolerance + borders ||
            parentBox.top < body.top || parentBox.bottom > body.bottom) return false;
      }
      columnBounds = { left: parentBox.left, right: parentBox.right, top: paragraphBox.top - .5, bottom: paragraphBox.bottom + .5 };
      elements.push(table);
    } else if (Math.abs(reference.cellWidth.outer - parentBox.width) > tolerance ||
        Math.abs(reference.cellWidth.text - paragraphBox.width) > tolerance) {
      if (!explicit || !matchesRoundedCellMeasure(parent, reference.cellWidth, parentBox, paragraphBox, true)) return false;
      roundedCellMeasure = true;
    }
    elements.push(row);
  } else if (reference.bodyWidth && !fragment && parent.matches('.print-content') &&
      !paragraph.closest('td,th,[data-column-count],[data-document-textbox]') && !paragraph.hasAttribute('data-print-soft-end')) {
    const css = getComputedStyle(paragraph), indent = parseFloat(css.textIndent);
    // Allow the source precision, a DOCX twip and each CSS layout edge.
    // A changed body measure keeps its natural wrap.
    const tolerance = .05 + 1 / 15 + 2 / 64;
    if (css.textAlign !== 'justify' || !(indent <= 0) || Math.abs(paragraphBox.width - indent - reference.bodyWidth) > tolerance) return false;
  } else if (column && !fragment && parent.matches('[data-column-count="2"]') &&
      !paragraph.matches('[data-print-paragraph-fragment],[data-print-soft-end],[data-paragraph-border-box],[data-word-paragraph-bottom-border]') &&
      !paragraph.closest('td,th,[data-document-textbox],[data-document-note]') && paragraph.getClientRects().length === 1) {
    const css = getComputedStyle(paragraph);
    const indent = Math.min(0, parseFloat(css.textIndent) || 0);
    const tolerance = .05 + 1 / 15 + 2 / 64;
    if (css.textAlign !== 'justify' || css.writingMode !== 'horizontal-tb' || css.direction !== 'ltr' ||
        css.transform !== 'none' || Math.abs(paragraphBox.width - indent - reference.columnWidth!) > tolerance) return false;
    columnBounds = { left: Math.max(parentBox.left, paragraphBox.left + indent - 2 / 64),
      right: Math.min(parentBox.right, paragraphBox.right + 2 / 64), top: paragraphBox.top - .5, bottom: paragraphBox.bottom + .5 };
  } else return false;
  if (explicit) {
    const table = parent.closest('table');
    if (fragment || terminal || column || columnCell || !reference.cellWidth ||
        !(parent instanceof HTMLTableCellElement) || !table || !fixedSingleLineTable(table, true)) return false;
  }
  const previous = paragraphLineOffsets(paragraph);
  if (previous.length !== sourceBreaks.length) return false;
  if (!preserveMatchingBreaks && previous.every((offset, index) =>
    normalizeHwpWrapText(text.slice(0, offset)).length === sourceBreaks[index])) return;
  if (column && columnBounds) {
    const restore = tryLocalColumnBoundary(paragraph, text, sourceBreaks, previous, columnBounds, range);
    if (restore) return restore;
  }
  let source = paragraph;
  let terminalBounds: { left: number; right: number; top: number; bottom: number } | undefined;
  let terminalMetrics: { font: number; advance: number; end: number; inkLeft: number; oldBottom: number } | undefined;
  if (terminal) {
    const css = getComputedStyle(paragraph);
    const font = parseFloat(css.fontSize), advance = Number(paragraph.dataset.hwpCellAdvance), end = Number(paragraph.dataset.hwpCellEnd);
    if (!(font > 0 && end > 0 && advance > end)) return false;
    const oldBottom = terminalInkBottom(paragraph, range);
    if (oldBottom === null) return false;
    const indent = Math.min(0, parseFloat(css.textIndent) || 0);
    terminalMetrics = { font, advance, end, oldBottom, inkLeft: Math.max(parentBox.left, paragraphBox.left + indent) };
    range.selectNodeContents(paragraph);
    const oldRects = [...range.getClientRects()].filter(rect => rect.width > 0 && rect.height > 0);
    if (!oldRects.length) return false;
    terminalBounds = {
      left: Math.max(parentBox.left - .5, Math.min(...oldRects.map(rect => rect.left))),
      right: Math.min(parentBox.right + .5, Math.max(...oldRects.map(rect => rect.right))),
      top: Math.max(parentBox.top - .5, Math.min(...oldRects.map(rect => rect.top))),
      bottom: Math.min(parentBox.bottom + .5, Math.max(...oldRects.map(rect => rect.bottom))),
    };
    source = paragraph.cloneNode(true) as HTMLElement;
    // The prefix currently ends at the browser's old last-line start.
    // Move it only after the recorded breaks produce a valid new end.
    unwrapHwpCellPrefixes(source);
    const table = parent.closest('table');
    if (!table) return false;
    elements.push(table);
  }
  const boundaries = sourceLineBoundaries(source, sourceBreaks);
  if (!boundaries) return false;
  const sizes = elements.map(geometry);
  // Source top alignment can place the first line's font box above its
  // paragraph. Shortening that same first line retains its leading envelope
  // while the preceding text and physical body still bound the new boxes.
  const keepsFirstLine = reference.bodyWidth && parent.matches('.print-content') &&
    (paragraph.dataset.hwpBodyLayout === 'top' || paragraph.dataset.hwpBodyLayout === 'mixed-top') &&
    sourceBreaks[0] <= normalizeHwpWrapText(text.slice(0, previous[0])).length;
  const top = keepsFirstLine ? retainedFirstLineTop(paragraph, range, parentBox, paragraphBox) : paragraphBox.top - .5;
  const original = [...paragraph.childNodes];
  const terminalStyle = terminal ? paragraph.style.cssText : '';
  const restore = () => {
    paragraph.replaceChildren(...original);
    if (terminal) paragraph.style.cssText = terminalStyle;
  };
  const projectedLines = lineOpportunities(source, boundaries, range);
  if (explicit) {
    for (const marker of projectedLines.querySelectorAll('wbr')) marker.replaceWith(document.createElement('br'));
  }
  if (column || columnCell) {
    // The synthetic line wrapper should not contribute another font strut.
    // Its children still inherit the preserved white-space policy.
    for (const line of projectedLines.children)
      if (line instanceof HTMLSpanElement) line.style.display = 'contents';
  }
  paragraph.replaceChildren(projectedLines);
  if (terminalMetrics) {
    const { font, advance, end } = terminalMetrics;
    const terminalPlacement = hwpCellTerminalLayout(paragraph, font, advance, end);
    if (!terminalPlacement) { restore(); return false; }
    applyHwpCellTerminal(paragraph, terminalPlacement, advance, end);
  }
  range.selectNodeContents(paragraph);
  const boxes = [...range.getClientRects()].filter(rect => rect.width > 0 && rect.height > 0);
  const envelope = terminalBounds ?? columnBounds ?? { left: parentBox.left - .5, right: parentBox.right + .5,
    top, bottom: paragraphBox.bottom + .5 };
  const crossesEnvelope = boxes.some(rect => rect.left < envelope.left || rect.right > envelope.right ||
    rect.top < envelope.top || rect.bottom > envelope.bottom) ||
    (terminalMetrics ? !horizontalInkFits(paragraph, range, { left: terminalMetrics.inkLeft, right: Math.min(parentBox.right, paragraphBox.right) })
      : boxes.some(rect => rect.top < paragraphBox.top - .5) && !horizontalInkFits(paragraph, range, parentBox));
  // An advance box can include trailing spacing beyond the right edge.
  // The source line can still fit if its actual ink is wholly inside the
  // cell and its vertical envelope and the geometry checks below hold.
  const outside = crossesEnvelope && (!explicit || !boxes.length ||
    boxes.some(rect => rect.left < envelope.left || rect.top < envelope.top || rect.bottom > envelope.bottom) ||
    !horizontalInkFits(paragraph, range, parentBox));
  const lines = paragraphLineOffsets(paragraph);
  if (column && paragraph.getClientRects().length !== 1 ||
    paragraph.textContent !== text || elements.some((element, i) => geometry(element) !== sizes[i]) ||
    lines.length !== sourceBreaks.length || lines.some((offset, i) => offset !== boundaries[i + 1].offset + (explicit ? i + 1 : 0)) ||
    (outside || roundedCellMeasure) && (!explicit || !(parent instanceof HTMLTableCellElement) || !reference.cellWidth ||
      !roundedCellMeasure && !matchesRoundedCellMeasure(parent, reference.cellWidth, parentBox, paragraphBox, true) ||
      !inkFitsRoundedCellLines(paragraph, parent, parentBox, range, lines.length + 1))) {
    restore();
    // Keep an existing exact projection. Mixed slots are a fallback when
    // that projection cannot preserve the authored end at the same height.
    return (column && columnBounds && reference.columnPitches &&
      tryMixedColumnPair(paragraph, text, sourceBreaks[0], reference.columnPitches, columnBounds, range)) || false;
  }
  if (terminalMetrics) {
    const bottom = terminalInkBottom(paragraph, range);
    // Leave a rasterization margin for new ink. Near the edge only retain the
    // prior ink envelope, measured in the same paragraph and physical strut.
    if (bottom === null || bottom > Math.max(terminalMetrics.oldBottom, paragraphBox.bottom - .5)) {
      restore();
      return false;
    }
  }
  return restore;
}

/** Restore text expansion witnessed in a complete source row before page
 * splitting. Never set a row height or change the editable paragraph. */
export function layoutHwpExpandedCellRows(table: HTMLTableElement) {
  if (!table.querySelector(`td > p[${hwpSourceWrapAttribute}*="expandedRowHeight"]`) ||
      !fixedSingleLineTable(table)) return;
  const plainGrid = !table.querySelector('table,img,svg,[data-note-reference]');
  const range = document.createRange();
  for (const row of table.rows) {
    const cells = [...row.cells];
    if (cells.length < 2 || row.dataset.wordRowRule === 'exact' ||
        row.querySelector('table,img,svg,[data-note-reference]') ||
        cells.some(cell => cell.tagName !== 'TD' || cell.rowSpan !== 1 || cell.colSpan !== 1)) continue;
    const oldHeight = row.getBoundingClientRect().height;
    const edits: { paragraph: HTMLParagraphElement; cell: HTMLTableCellElement; original: Node[] }[] = [];
    let sourceHeight = 0, minimumPitch = Infinity;
    for (const cell of cells) {
      if (getComputedStyle(cell).verticalAlign !== 'middle') continue;
      for (const paragraph of cell.querySelectorAll<HTMLParagraphElement>(`:scope > p[${hwpSourceWrapAttribute}]`)) {
        const proof = parseHwpSourceWrap(paragraph.dataset.hwpSourceWrap), text = paragraph.textContent ?? '';
        const height = proof?.expandedRowHeight, width = proof?.cellWidth, witnessed = proof?.expandedRowLine;
        if (!height || !width || !(height > oldHeight) || !plainGrid && !witnessed ||
            hwpWrapFingerprint(text) !== proof.fingerprint || normalizeHwpWrapText(text).length !== proof.ends.at(-1) ||
            !plainHwpSourceParagraph(paragraph) || paragraph.querySelector(excludedInline) ||
            paragraphLineOffsets(paragraph).length + 2 !== proof.ends.length) continue;
        const box = paragraph.getBoundingClientRect(), cellBox = cell.getBoundingClientRect();
        const cellCss = getComputedStyle(cell), style = getComputedStyle(paragraph), pitch = parseFloat(style.lineHeight);
        const precision = .05 + 1 / 15 + 2 / 64;
        const borderWidth = parseFloat(cellCss.borderLeftWidth) + parseFloat(cellCss.borderRightWidth);
        const hanging = witnessed ? -Math.min(0, parseFloat(style.textIndent) || 0) : 0;
        if (witnessed && (table.dataset.hwpCenteredBorders !== 'continuation' || cell.children.length < 2 ||
            style.textAlign !== 'justify' || Math.abs(hanging - witnessed.hanging) > precision ||
            Math.abs(parseFloat(style.marginLeft) - hanging) > precision || parseFloat(style.marginRight) ||
            Math.abs(pitch - witnessed.pitch) > precision || authoredCellParagraphScale(paragraph) !== witnessed.scale)) continue;
        const matchesMeasure = Math.abs(width.outer - cellBox.width) <= borderWidth + precision &&
          Math.abs(width.text - box.width - hanging) <= borderWidth + precision;
        if (!(pitch > 0) || !matchesMeasure && (witnessed || !matchesRoundedCellMeasure(cell, width, cellBox, box)) ||
            style.direction !== 'ltr' || style.writingMode !== 'horizontal-tb' || style.transform !== 'none' ||
            sourceHeight && Math.abs(height - sourceHeight) > precision) continue;

        const original = [...paragraph.childNodes];
        paragraph.replaceChildren(...original.map(node => node.cloneNode(true)));
        const cuts = sourceLineBoundaries(paragraph, proof.ends.slice(0, -1));
        if (!cuts) { paragraph.replaceChildren(...original); continue; }
        if (witnessed) {
          const measured = lineOpportunities(paragraph, cuts, range);
          for (const marker of measured.querySelectorAll('wbr')) marker.replaceWith(document.createElement('br'));
          paragraph.replaceChildren(measured);
        } else {
          for (const boundary of cuts.slice(1, -1).reverse()) {
            range.setStart(boundary.node, boundary.at);
            range.collapse(true);
            range.insertNode(document.createElement('br'));
          }
        }
        const ends = paragraphLineOffsets(paragraph), after = paragraph.getBoundingClientRect();
        if (paragraph.textContent !== text || after.width !== box.width || !(after.height > box.height) ||
            ends.length + 1 !== proof.ends.length || ends.some((end, i) => end !== cuts[i + 1].offset + i + 1)) {
          paragraph.replaceChildren(...original);
          continue;
        }
        sourceHeight = height;
        minimumPitch = Math.min(minimumPitch, pitch);
        edits.push({ paragraph, cell, original });
      }
    }
    const newHeight = row.getBoundingClientRect().height;
    // A short row must recover to within half a source line, rather than
    // gaining partial or extra lines. A covered row must keep all painted text.
    if (!(newHeight > oldHeight && Math.abs(newHeight - sourceHeight) < minimumPitch / 2) ||
        edits.some(({ paragraph, cell }) => {
          const bounds = cell.getBoundingClientRect();
          const css = getComputedStyle(cell);
          const top = bounds.top + parseFloat(css.borderTopWidth) / 2;
          const bottom = bounds.bottom - parseFloat(css.borderBottomWidth) / 2;
          range.selectNodeContents(paragraph);
          return [...range.getClientRects()].some(rect => rect.width > 0 && rect.height > 0 &&
            (rect.top < top || rect.bottom > bottom)) || !horizontalInkFits(paragraph, range, bounds);
        })) {
      for (const { paragraph, original } of edits) paragraph.replaceChildren(...original);
    }
  }
}

/** Paint the stored width on a continued cell before the split, only if
 * it restores the observed line count with the existing face and indent. */
export function layoutHwpAuthoredHangingCellScale(table: HTMLTableElement) {
  const candidates = table.querySelectorAll<HTMLParagraphElement>(`p[${hwpSourceWrapAttribute}*='"continued":true,"scale":']`);
  if (!candidates.length || table.dataset.hwpCenteredBorders !== 'continuation' || !fixedSingleLineTable(table)) return;
  const range = document.createRange();
  const precision = .05 + 1 / 15 + 2 / 64;
  for (const paragraph of candidates) {
    const proof = parseHwpSourceWrap(paragraph.dataset.hwpSourceWrap), source = proof?.authoredCellFace;
    const scale = source?.scale, hanging = source?.hanging;
    const cell = paragraph.parentElement, row = cell?.parentElement;
    if (!proof?.cellWidth || !source || !scale || !hanging ||
        !(cell instanceof HTMLTableCellElement) || cell.tagName !== 'TD' ||
        !(row instanceof HTMLTableRowElement) || row.dataset.wordRowRule === 'exact' || cell.closest('table') !== table ||
        cell.colSpan !== 1 || cell.rowSpan !== 1 || cell.children.length < 2 ||
        !plainHwpSourceParagraph(paragraph) || paragraph.querySelector(excludedInline) ||
        authoredCellParagraphScale(paragraph) !== scale) continue;
    const text = paragraph.textContent ?? '';
    if (hwpWrapFingerprint(text) !== proof.fingerprint || normalizeHwpWrapText(text).length !== proof.ends.at(-1)) continue;

    const css = getComputedStyle(paragraph), cellCss = getComputedStyle(cell);
    const cellLine = parseHwpCellLine(paragraph.dataset.hwpCellLine);
    const height = cellLine ? hwpCellAdvance(cellLine, source.size) : parseFloat(css.lineHeight);
    const before = paragraph.getBoundingClientRect();
    const borders = parseFloat(cellCss.borderLeftWidth) + parseFloat(cellCss.borderRightWidth);
    const width = cell.getBoundingClientRect().width;
    if (css.transform !== 'none' || css.textAlign !== 'justify' ||
        Math.abs(parseFloat(css.textIndent) + hanging) > precision ||
        Math.abs(parseFloat(css.marginLeft) - hanging) > precision || parseFloat(css.marginRight) ||
        Math.abs(height - source.pitch) > precision || Math.abs(width - proof.cellWidth.outer) > borders + precision ||
        Math.abs(before.width + hanging - proof.cellWidth.text) > borders + precision ||
        paragraphLineOffsets(paragraph).length + 1 <= proof.ends.length) continue;

    const original = paragraph.style.cssText;
    const originalRow = row.getBoundingClientRect();
    paragraph.style.cssText += ';' + cellCharacterScaleDimensions(scale, hanging);
    const after = paragraph.getBoundingClientRect(), bounds = cell.getBoundingClientRect();
    range.selectNodeContents(paragraph);
    const ink = [...range.getClientRects()].filter(box => box.width > 0 && box.height > 0);
    if (Math.abs(after.width - before.width) > precision || bounds.width !== width ||
        row.getBoundingClientRect().height > originalRow.height || after.height >= before.height ||
        paragraphLineOffsets(paragraph).length + 1 !== proof.ends.length ||
        !ink.length || ink.some(box => box.top < Math.max(bounds.top, after.top) || box.bottom > Math.min(bounds.bottom, after.bottom)) ||
        !horizontalInkFits(paragraph, range, bounds, scale / 100))
      paragraph.style.cssText = original;
  }
}

/** Apply an authored width with the same fallback used by the other
 * continuation cells. Retain it only when the original lines and the painted
 * bounds agree, so the regular table splitter can use the recovered height. */
export function layoutHwpCellWidthFallbacks(table: HTMLTableElement, measureForNative = false) {
  const candidates = table.querySelectorAll<HTMLParagraphElement>(hwpCellWidthFallbackSelector);
  if (!candidates.length || table.dataset.hwpCenteredBorders !== 'continuation' || !fixedSingleLineTable(table)) return;
  const range = document.createRange();
  const precision = .05 + 1 / 15 + 2 / 64;
  const expectedFace = (value: string) => value.split(',')[0].trim().replace(/^['"]|['"]$/g,'') === hwpCompactCellFace;
  const inspect = (p: HTMLParagraphElement) => {
    const proof = parseHwpSourceWrap(p.dataset.hwpSourceWrap), source = proof?.cellWidthFallback;
    const cell = p.parentElement, row = cell?.parentElement, intact = source?.intactRow;
    if (!proof?.cellWidth || !source || measureForNative && intact || !(cell instanceof HTMLTableCellElement) || cell.tagName !== 'TD' ||
        !(row instanceof HTMLTableRowElement) || cell.closest('table') !== table || cell.rowSpan !== 1 ||
        cell.colSpan !== (intact?.columns ?? 1) || intact && rowSplitPolicy(row.dataset) === 'allow' ||
        row.dataset.wordRowRule === 'exact' ||
        p.querySelector(':not(span,br),[id],[data-word-overlap],[data-docx-tab],[data-document-field],[data-flow-widget]') ||
        p.querySelectorAll('br').length !== Number(!!source.hardBreak) ||
        authoredCellParagraphScale(p,source.hardBreak) !== source.scale) return;
    const css = getComputedStyle(p), cellStyle = getComputedStyle(cell);
    const cellBefore = cell.getBoundingClientRect(), prior = p.getBoundingClientRect(), rowBefore = row.getBoundingClientRect();
    const factor = source.scale / 100, alreadyScaled = css.transform === `matrix(${factor}, 0, 0, 1, 0, 0)`;
    const linePolicy = parseHwpCellLine(p.dataset.hwpCellLine);
    const pitch = linePolicy ? hwpCellAdvance(linePolicy,source.size) : parseFloat(css.lineHeight);
    const borders = parseFloat(cellStyle.borderLeftWidth) + parseFloat(cellStyle.borderRightWidth);
    if (!['middle','top'].includes(cellStyle.verticalAlign) || css.textAlign !== 'justify' || css.direction !== 'ltr' ||
        css.writingMode !== 'horizontal-tb' || !alreadyScaled && css.transform !== 'none' ||
        Math.abs(parseFloat(css.marginLeft) - source.hanging) > precision || parseFloat(css.marginRight) ||
        Math.abs(parseFloat(css.textIndent) * (alreadyScaled ? factor : 1) + source.hanging) > precision ||
        Math.abs(pitch - source.pitch) > precision ||
        Math.abs(cellBefore.width - Math.round(proof.cellWidth.outer)) > borders + precision ||
        Math.abs(prior.width + source.hanging - (intact?.textWidth ?? proof.cellWidth.text)) > borders + .5 + precision) return;

    const walker = paragraphTextWalker(p), pieces: string[] = [];
    let supported = true;
    while (walker.nextNode()) {
      const node = walker.currentNode;
      pieces.push(node instanceof Text ? node.data : '\n');
      if (!(node instanceof Text) || !node.data.trim()) continue;
      const style = getComputedStyle(node.parentElement!);
      if (!expectedFace(style.fontFamily) || Math.abs(parseFloat(style.fontSize) - source.size) > .07 ||
          style.fontStyle !== 'normal' || style.fontWeight !== '400' || style.transform !== 'none' && node.parentElement !== p ||
          style.textDecorationLine !== 'none' || style.textShadow !== 'none') supported = false;
    }
    const text = pieces.join('');
    if (!supported || hwpWrapFingerprint(text) !== proof.fingerprint ||
        normalizeHwpWrapText(text).length !== proof.ends.at(-1)) return;
    return { p, proof, source, cell, row, intact, cellStyle, cellBefore, prior, rowBefore, factor, alreadyScaled, linePolicy, text };
  };


  const faceRuns = (paragraph: HTMLParagraphElement) =>
    [paragraph, ...paragraph.querySelectorAll<HTMLElement>('span')].filter(el => expectedFace(el.style.fontFamily));
  const matchesEnds = (paragraph: HTMLParagraphElement, proof: HwpSourceWrap, text: string) => {
    const ends = paragraphLineOffsets(paragraph);
    return ends.length + 1 === proof.ends.length &&
      ends.every((offset, i) => normalizeHwpWrapText(text.slice(0, offset)).length === proof.ends[i]);
  };
  type Measured = NonNullable<ReturnType<typeof inspect>>;
  const processed = new Set<HTMLTableCellElement>();
  const recoverPair = (info: Measured) => {
    const { cell, row, intact, cellBefore, rowBefore, cellStyle } = info;
    if (processed.has(cell) || !intact?.paragraphs || cell.children.length !== intact.paragraphs) return;
    processed.add(cell);
    const pair = [...cell.children].map(paragraph => paragraph instanceof HTMLParagraphElement ? inspect(paragraph) : undefined);
    if (!pair.every((item): item is Measured => !!item?.intact && item.intact.paragraphs === intact.paragraphs &&
        item.intact.height === intact.height && item.intact.textWidth === intact.textWidth &&
        item.proof.cellWidth?.outer === info.proof.cellWidth?.outer)) return;
    if (pair.every(({ p, proof, text }) => matchesEnds(p, proof, text))) return;

    const originals = pair.map(({ p }) => p);
    const trial = pair.map(({ p, alreadyScaled, source }) => {
      const copy = p.cloneNode(true) as HTMLParagraphElement;
      if (!alreadyScaled) copy.style.cssText += ';' + cellCharacterScaleDimensions(source.scale, source.hanging);
      for (const run of faceRuns(copy)) {
        if (!run.style.fontFamily.includes('Noto Sans KR')) run.style.fontFamily += ', "Noto Sans KR"';
      }
      return copy;
    });
    const projected = pair.map(({ p, proof }) => {
      const cuts = sourceLineBoundaries(p, proof.ends.slice(0, -1));
      return cuts && lineOpportunities(p, cuts, range);
    });
    const widths = [...row.cells].map(peer => peer.getBoundingClientRect().width);
    originals.forEach((p, i) => p.replaceWith(trial[i]));
    trial.forEach((copy, i) => {
      const { proof, text } = pair[i];
      layoutHwpCellLines(copy);
      // The normal authored width may already wrap at the right place.
      // Add print-only opportunities only to the sibling that needs them.
      const fitted = projected[i];
      if (!matchesEnds(copy, proof, text) && fitted && fitted.textContent === originals[i].textContent) {
        copy.replaceChildren(fitted);
        for (const run of faceRuns(copy)) {
          if (!run.style.fontFamily.includes('Noto Sans KR')) run.style.fontFamily += ', "Noto Sans KR"';
        }
        layoutHwpCellLines(copy);
      }
    });
    const bounds = cell.getBoundingClientRect();
    const after = row.getBoundingClientRect();
    const left = bounds.left + parseFloat(cellStyle.borderLeftWidth) / 2;
    const right = bounds.right - parseFloat(cellStyle.borderRightWidth) / 2;
    let bottom = bounds.top + parseFloat(cellStyle.borderTopWidth) / 2;
    const limit = bounds.bottom - parseFloat(cellStyle.borderBottomWidth) / 2;
    const valid = bounds.width === cellBefore.width && after.top === rowBefore.top && after.height < rowBefore.height &&
      Math.abs(after.height - intact.height) <= .5 &&
      [...row.cells].every((peer, i) => peer.getBoundingClientRect().width === widths[i]) && trial.every((copy, i) => {
        const { p, proof, source, text, factor, prior } = pair[i];
        const rect = copy.getBoundingClientRect();
        if (copy.textContent !== p.textContent || copy.querySelector('br') ||
            Math.abs(rect.width - prior.width) > precision || !matchesEnds(copy, proof, text) ||
            !horizontalInkFits(copy, range, { left, right }, factor)) return false;
        const paint = measuredHwpInkAtLineOffsets(copy, paragraphLineOffsets(copy), source.size, true);
        if (!paint || paint.length !== proof.ends.length ||
            paint.some((line, j) => line.top < (j ? paint[j - 1].bottom : bottom) || line.bottom > limit)) return false;
        bottom = paint.at(-1)!.bottom;
        return true;
      });
    if (!valid) trial.forEach((copy, i) => copy.replaceWith(originals[i]));
  };

  for (const p of candidates) {
    const measured = inspect(p);
    if (!measured) continue;
    const { proof, source, cell, row, intact, cellStyle, cellBefore, prior, rowBefore, factor, alreadyScaled, linePolicy, text } = measured;
    if (intact?.paragraphs) {
      recoverPair(measured);
      continue;
    }
    if (paragraphLineOffsets(p).length <= proof.ends.length - 1) continue;
    const original = p.style.cssText;
    const styled = faceRuns(p), families = styled.map(el => el.style.fontFamily);
    const nativeGeometry = measureForNative && !alreadyScaled ? nativeCellMeasureProperties.map(property =>
      [p.style.getPropertyValue(property), p.style.getPropertyPriority(property)]) : undefined;
    // Measure compact terminal leading on the same private print copy as
    // the authored width. The original row height and every line end still
    // have to agree before either change can affect the split.
    const measureTerminal = Boolean(intact && linePolicy &&
      (source.hardBreak || linePolicy.advance > linePolicy.end + .05));
    const sourceCuts = measureTerminal && !source.hardBreak && proof.ends.length > 1
      ? sourceLineBoundaries(p, proof.ends.slice(0, -1)) : undefined;
    // Range boundaries point into the original paragraph. Project them there
    // before replacing its children in any trial copy.
    const fittedContent = sourceCuts && lineOpportunities(p, sourceCuts, range);
    const trials = [false, true].flatMap(fittedLines => fittedLines && !fittedContent ? []
      : [false, true].map(missingFace => ({ fittedLines, missingFace })));
    for (const { fittedLines, missingFace } of trials) {
      if (missingFace && !styled.length) continue;
      const candidate = measureTerminal ? p.cloneNode(true) as HTMLParagraphElement : p;
      if (fittedLines && fittedContent) {
        candidate.replaceChildren(fittedContent.cloneNode(true));
        if (candidate.textContent !== p.textContent) continue;
      }
      if (measureTerminal) p.replaceWith(candidate);
      if (missingFace) {
        const runs = measureTerminal ? faceRuns(candidate) : styled;
        for (const run of runs) {
          if (!run.style.fontFamily.includes('Noto Sans KR')) run.style.fontFamily += ', "Noto Sans KR"';
        }
      }
      if (!alreadyScaled) candidate.style.cssText += ';' + cellCharacterScaleDimensions(source.scale,source.hanging);
      if (measureTerminal) layoutHwpCellLines(candidate);
      const after = candidate.getBoundingClientRect(), bounds = cell.getBoundingClientRect();
      range.selectNodeContents(candidate);
      const paint = [...range.getClientRects()].filter(rect => rect.width > 0 && rect.height > 0);
      const rowHeight = row.getBoundingClientRect().height;
      const ink = fittedLines
        ? { left: bounds.left + parseFloat(cellStyle.borderLeftWidth) / 2,
          right: bounds.right - parseFloat(cellStyle.borderRightWidth) / 2 }
        : bounds;
      if (Math.abs(after.width - prior.width) <= precision && bounds.width === cellBefore.width &&
          rowHeight <= rowBefore.height && (!intact || Math.abs(rowHeight - intact.height) <= .5) && after.height < prior.height &&
          matchesEnds(candidate, proof, text) && paint.length && paint.every(rect => rect.top >= bounds.top && rect.bottom <= bounds.bottom) &&
          horizontalInkFits(candidate,range,ink,factor)) {
        if (measureForNative) {
          if (nativeGeometry) p.dataset.hwpNativeCellMeasureGeometry = JSON.stringify(nativeGeometry);
          styled.forEach((el, i) => {
            if (el.style.fontFamily !== families[i]) el.dataset.hwpNativeCellMeasureFace = families[i];
          });
        }
        break;
      }
      if (measureTerminal) candidate.replaceWith(p);
      else {
        p.style.cssText = original;
        styled.forEach((el,i) => { el.style.fontFamily = families[i]; });
      }
    }
  }
}

const nativeCellMeasureProperties = ['width', 'text-indent', 'transform', 'transform-origin'] as const;

/** Native pagination can consume the measured breaks. The written runs and
 * paragraph keep their authored dimensions and font names. */
export function clearHwpNativeCellWidthProjection(root: HTMLElement) {
  for (const node of root.querySelectorAll<HTMLElement>('[data-hwp-native-cell-measure-geometry]')) {
    const values: [string, string][] = JSON.parse(node.dataset.hwpNativeCellMeasureGeometry!);
    nativeCellMeasureProperties.forEach((property, index) => node.style.setProperty(property, ...values[index]));
    delete node.dataset.hwpNativeCellMeasureGeometry;
  }
  for (const node of root.querySelectorAll<HTMLElement>('[data-hwp-native-cell-measure-face]')) {
    node.style.fontFamily = node.dataset.hwpNativeCellMeasureFace!;
    delete node.dataset.hwpNativeCellMeasureFace;
  }
}

/** Supply a substitute only for a proven, still unchanged paragraph on
 * a source table's fragment edge. Let the ordinary row splitter use the
 * recovered height, with all authored fonts, spaces and hard breaks intact. */
export function layoutHwpCompactCellFaces(table: HTMLTableElement, missingBody: boolean, missingHeader: boolean, phase: 'paginate' | 'paint' = 'paginate') {
  const candidates = table.querySelectorAll<HTMLParagraphElement>(hwpCellFaceFallbackSelector);
  if (!candidates.length || !fixedSingleLineTable(table)) return;
  const range = document.createRange();
  const printBody = phase === 'paint' ? table.closest('.print-content')?.getBoundingClientRect() : undefined;
  const precision = .05 + 1 / 15 + 2 / 64;
  for (const paragraph of candidates) {
    const proof = parseHwpSourceWrap(paragraph.dataset.hwpSourceWrap), source = proof?.authoredCellFace;
    const cell = paragraph.parentElement, row = cell?.parentElement;
    const continued = source?.continued === true;
    if (!proof?.cellWidth || !source || !(source.header ? missingHeader : missingBody) ||
        !(cell instanceof HTMLTableCellElement) || cell.tagName !== (source.header ? 'TH' : 'TD') ||
        !(row instanceof HTMLTableRowElement) || cell.closest('table') !== table ||
        source.header && (table.rows[0] !== row || [...row.cells].some(item => item.tagName !== 'TH') || cell.children.length !== 1) ||
        !continued && source.hanging !== undefined && (cell.children.length !== 2 || cell.children[1] !== paragraph) ||
        continued && table.dataset.hwpCenteredBorders !== 'continuation' ||
        cell.colSpan !== 1 || cell.rowSpan !== 1 || row.dataset.wordRowRule === 'exact' ||
        !continued && Math.abs(parseFloat(row.style.height) - source.rowHeight) > precision ||
        paragraph.querySelector(':not(span,br),[id],[data-word-overlap],[data-docx-tab],[data-document-field],[data-flow-widget]') ||
        paragraph.querySelectorAll('br').length !== (continued ? 0 : proof.ends.length - 1)) continue;

    const style = getComputedStyle(paragraph), bounds = cell.getBoundingClientRect(), box = paragraph.getBoundingClientRect();
    const padding = getComputedStyle(cell), indent = Math.min(0, parseFloat(style.textIndent) || 0);
    const borders = parseFloat(padding.borderLeftWidth) + parseFloat(padding.borderRightWidth);
    const cellLine = parseHwpCellLine(paragraph.dataset.hwpCellLine);
    const pitch = cellLine ? hwpCellAdvance(cellLine, source.size) : parseFloat(style.lineHeight);
    if ((!continued && padding.verticalAlign !== 'middle' || continued && !['middle', 'top'].includes(padding.verticalAlign)) ||
        style.direction !== 'ltr' || style.writingMode !== 'horizontal-tb' ||
        style.transform !== 'none' || !(Math.abs(pitch - source.pitch) <= precision) ||
        source.header && (!cellLine || style.textAlign !== 'center' || indent || parseFloat(style.marginLeft) || parseFloat(style.marginRight)) ||
        source.hanging !== undefined && (style.textAlign !== 'justify' || Math.abs(indent + source.hanging) > precision ||
          Math.abs(parseFloat(style.marginLeft) - source.hanging) > precision || parseFloat(style.marginRight)) ||
        Math.abs(bounds.width - Math.round(proof.cellWidth.outer)) > borders + precision ||
        Math.abs(box.width - indent - proof.cellWidth.text) > borders + .5 + precision) continue;

    const family = source.header ? hwpCompactHeaderFace : hwpCompactCellFace;
    const originalFace = (value: string) => value.replace(/^['"]|['"]$/g, '') === family;
    const walker = paragraphTextWalker(paragraph), logical: string[] = [];
    let safeRuns = true;
    while (walker.nextNode()) {
      const node = walker.currentNode;
      logical.push(node instanceof Text ? node.data : '\n');
      if (!(node instanceof Text) || !node.data.trim()) continue;
      const css = getComputedStyle(node.parentElement!);
      const originalWeight = css.fontWeight === '400' || source.header && css.fontWeight === '600';
      if (!originalFace(css.fontFamily) || Math.abs(parseFloat(css.fontSize) - source.size) > .07 ||
          css.fontStyle !== 'normal' || !originalWeight || css.transform !== 'none' ||
          css.textDecorationLine !== 'none' || css.textShadow !== 'none') safeRuns = false;
    }
    const text = logical.join('');
    if (!safeRuns || hwpWrapFingerprint(text) !== proof.fingerprint ||
        normalizeHwpWrapText(text).length !== proof.ends.at(-1)) continue;
    const endpointsMatch = (offsets: number[]) => offsets.length + 1 === proof.ends.length &&
      offsets.every((offset, index) => normalizeHwpWrapText(text.slice(0, offset)).length === proof.ends[index]);
    const current = paragraphLineOffsets(paragraph);
    // The row splitter can use the horizontal after-last-glyph position.
    // A same-count face change belongs only on an already paginated fragment.
    if (phase === 'paint') {
      if (!printBody || !continued || current.length + 1 !== proof.ends.length || endpointsMatch(current)) continue;
    } else if (current.length + 1 <= proof.ends.length) continue;
    const styled = [paragraph, ...paragraph.querySelectorAll<HTMLElement>('span')].filter(node => originalFace(node.style.fontFamily));
    if (!styled.length) continue;
    const previous = styled.map(node => [node.style.fontFamily, node.style.fontWeight]);
    // Header rows can have a stored minimum that hides a second compact
    // strut. Compare the same line metrics used by the later page fit.
    if (source.header) layoutHwpCellLines(paragraph);
    const paragraphBefore = paragraph.getBoundingClientRect();
    const rowBefore = row.getBoundingClientRect();
    for (const node of styled) {
      node.style.fontFamily += ', "Noto Sans KR"';
      if (source.header) node.style.fontWeight = '700';
    }
    if (source.header) layoutHwpCellLines(paragraph);
    const next = paragraph.getBoundingClientRect(), after = row.getBoundingClientRect(), cellAfter = cell.getBoundingClientRect();
    const ends = paragraphLineOffsets(paragraph);
    range.selectNodeContents(paragraph);
    const glyphs = [...range.getClientRects()].filter(rect => rect.width > 0 && rect.height > 0);
    const rowFits = continued ? after.height <= rowBefore.height : after.height < rowBefore.height &&
      Math.abs(after.height - source.rowHeight) < source.pitch / 2;
    const heightFits = phase === 'paint'
      ? next.x === paragraphBefore.x && next.y === paragraphBefore.y && next.height === paragraphBefore.height &&
        after.width === rowBefore.width && after.height === rowBefore.height
      : next.height < paragraphBefore.height && rowFits;
    const valid = next.width === paragraphBefore.width && heightFits && endpointsMatch(ends) &&
      glyphs.length && glyphs.every(rect => rect.top >= Math.max(cellAfter.top, printBody?.top ?? -Infinity) &&
        rect.bottom <= Math.min(cellAfter.bottom, printBody?.bottom ?? Infinity)) &&
      horizontalInkFits(paragraph, range, cellAfter);
    if (!valid) {
      styled.forEach((node, i) => { [node.style.fontFamily, node.style.fontWeight] = previous[i]; });
      if (source.header) layoutHwpCellLines(paragraph);
    }
  }
}

/** Try the original line ends of an otherwise unchanged, one-line-short
 * paragraph after its table established a natural first-page break. */
export function restoreHwpContinuedTableLines(row: HTMLTableRowElement, priorFragmentMatches: boolean) {
  if (row.dataset.wordRowRule === 'exact') return;
  const range = document.createRange();
  for (const cell of row.cells) {
    if (cell.tagName !== 'TD' || cell.rowSpan !== 1 || cell.colSpan !== 1 ||
        cell.querySelector('table,[data-note-reference]')) continue;
    for (const element of cell.children) {
      if (!(element instanceof HTMLParagraphElement) ||
          !hwpSourceAdvanceParagraph(element) || element.querySelector(excludedInline)) continue;
      const reference = parseHwpSourceWrap(element.dataset.hwpSourceWrap);
      const text = element.textContent ?? '';
      if (!reference?.cellWidth || reference.ends.length < 2 ||
          (!priorFragmentMatches && (!reference.splitCell || row.cells.length !== 2)) ||
          hwpWrapFingerprint(text) !== reference.fingerprint ||
          reference.ends.at(-1) !== normalizeHwpWrapText(text).length ||
          paragraphLineOffsets(element).length + 2 !== reference.ends.length) continue;
      const css = getComputedStyle(cell), prior = element.getBoundingClientRect(), frame = cell.getBoundingClientRect();
      const tolerance = (parseFloat(css.borderLeftWidth) || 0) + (parseFloat(css.borderRightWidth) || 0) + .05 + 1 / 15;
      if (!(parseFloat(getComputedStyle(element).textIndent) <= 0) ||
          Math.abs(reference.cellWidth.outer - frame.width) > tolerance ||
          Math.abs(reference.cellWidth.text - prior.width) > tolerance) continue;
      const offsets = sourceLineBoundaries(element, reference.ends.slice(0, -1));
      if (!offsets) continue;
      const contents = lineOpportunities(element, offsets, range);
      const opportunities = [...contents.querySelectorAll('wbr')];
      if (priorFragmentMatches)
        for (const marker of opportunities) marker.replaceWith(document.createElement('br'));
      const original = [...element.childNodes];
      element.replaceChildren(contents);
      let actual = paragraphLineOffsets(element);
      let addedTerminalBreak = false;
      // The recovered prefix already breaks at every authored position.
      // A narrower fallback can still join the last two lines across WBR.
      // Give that one witnessed boundary a hard break on this print clone.
      if (!priorFragmentMatches && reference.splitCell && actual.length >= 2 &&
          actual.length === opportunities.length - 1 &&
          actual.every((offset, index) => offset === offsets[index + 1].offset)) {
        opportunities.at(-1)!.replaceWith(document.createElement('br'));
        actual = paragraphLineOffsets(element);
        addedTerminalBreak = true;
      }
      const placed = element.getBoundingClientRect(), bounds = cell.getBoundingClientRect();
      range.selectNodeContents(element);
      const textBoxes = [...range.getClientRects()].filter(rect => rect.width > 0 && rect.height > 0);
      const contains = textBoxes.length && textBoxes.every(rect =>
        rect.left >= bounds.left - .12 && rect.right <= bounds.right + .12 && rect.top >= bounds.top - .12);
      const exactLines = actual.length === offsets.length - 2 && actual.every((offset, index) => {
        const syntheticBreaks = priorFragmentMatches ? index + 1 : Number(addedTerminalBreak && index === actual.length - 1);
        return offset === offsets[index + 1].offset + syntheticBreaks;
      }) && (priorFragmentMatches || horizontalInkFits(element, range, bounds));
      if (element.textContent !== text || placed.width !== prior.width || bounds.width !== frame.width ||
          placed.height <= prior.height || !contains ||
          !exactLines ||
          textBoxes.some(rect => rect.bottom > bounds.bottom + .12) && (terminalInkBottom(element, range) ?? Infinity) > bounds.bottom - .5) {
        element.replaceChildren(...original);
      }
    }
  }
}

/** Verify complete source lines before the print splitter chooses its cut.
 * The regular final paint pass cannot move text across separate pages. */
export function preserveHwpCellLinesBeforeSplit(paragraph: HTMLElement) {
  const cell = paragraph.parentElement;
  const table = cell?.closest('table');
  if (!(paragraph instanceof HTMLParagraphElement) || !(cell instanceof HTMLTableCellElement) ||
      cell.tagName !== 'TD' || cell.rowSpan !== 1 || cell.colSpan !== 1 ||
      !table || table.parentElement?.closest('td,th') || cell.querySelector('[data-note-reference]')) return;
  const related = table.dataset.hwpRelatedRowFlow === 'true';
  const sibling = cell.querySelector('table');
  if (sibling && (!related || sibling.parentElement !== cell || cell.querySelectorAll('table').length !== 1)) return;
  const nestedTableGeometry = sibling ? geometry(sibling) : undefined;
  const reference = parseHwpSourceWrap(paragraph.dataset.hwpSourceWrap);
  const text = paragraph.textContent ?? '';
  if (!reference?.cellWidth || reference.ends.length < 2 ||
      hwpWrapFingerprint(text) !== reference.fingerprint || normalizeHwpWrapText(text).length !== reference.ends.at(-1)) return;
  const ends = reference.ends.slice(0, -1), range = document.createRange();
  // A validated preceding row may now open a different physical fragment.
  // Keep its authored line ends if all original ink and geometry still fit.
  // When a nested table is a direct sibling, verify its own box too.
  const natural = tryWrap(paragraph, reference, ends, range);
  const restore = natural || (related && tryWrap(paragraph, reference, ends, range, false, true, true));
  if (!restore) return;
  if (sibling && nestedTableGeometry !== geometry(sibling)) { restore(); return; }
  return restore;
}

/** Reflow an already fitted ordinary cell without changing its measured
 * row or the sizing/spacing supplied by the authored text runs. */
function restoreFittedCellLineEnds(paragraph: HTMLParagraphElement, reference: HwpSourceWrap, range: Range) {
  const cell = paragraph.parentElement, table = cell?.closest('table'), body = cell?.closest('.print-content');
  const source = reference.cellWidth;
  if (!source || !(cell instanceof HTMLTableCellElement) || cell.tagName !== 'TD' ||
      cell.rowSpan !== 1 || cell.colSpan !== 1 || !table || !body ||
      table.dataset.hwpCenteredBorders !== 'continuation' || !fixedSingleLineTable(table) ||
      paragraph.matches('[data-print-paragraph-fragment],[data-print-soft-end],[data-paragraph-border-box],[data-word-paragraph-bottom-border]') ||
      paragraph.querySelector(excludedInline + ',:not(span,strong,b,em,i,u),[data-note-reference],[data-document-field],[data-word-overlap],[data-flow-widget]')) return;

  const text = paragraph.textContent ?? '', offsets = paragraphLineOffsets(paragraph);
  if (offsets.length + 1 !== reference.ends.length ||
      offsets.every((at, i) => normalizeHwpWrapText(text.slice(0, at)).length === reference.ends[i])) return;

  const style = getComputedStyle(paragraph), cellStyle = getComputedStyle(cell);
  const scale = style.transform === 'none' ? 1 : (authoredCellParagraphScale(paragraph) ?? 0) / 100;
  if (style.direction !== 'ltr' || style.writingMode !== 'horizontal-tb' || !scale ||
      scale !== 1 && style.transform !== `matrix(${scale}, 0, 0, 1, 0, 0)`) return;
  const cellBox = cell.getBoundingClientRect(), bodyBox = body.getBoundingClientRect();
  const paragraphBox = paragraph.getBoundingClientRect();
  const borders = parseFloat(cellStyle.borderLeftWidth) + parseFloat(cellStyle.borderRightWidth);
  const rounded = cellBox.width - source.outer;
  const precision = .05 + 1 / 15 + 2 / 64;
  const hanging = Math.min(0, (parseFloat(style.textIndent) || 0) * scale);
  const wholeWidth = reference.cellWidthFallback?.intactRow?.textWidth ?? (reference.authoredCellFace ? source.text : undefined);
  const expectedWidth = wholeWidth === undefined ? source.text : wholeWidth + hanging;
  if (Math.abs(rounded) > borders + .5 + precision ||
      Math.abs(paragraphBox.width - expectedWidth - rounded) > borders + precision) return;

  const bounds = { left: Math.max(bodyBox.left, cellBox.left + parseFloat(cellStyle.borderLeftWidth) / 2),
    right: Math.min(bodyBox.right, cellBox.right - parseFloat(cellStyle.borderRightWidth) / 2),
    top: Math.max(bodyBox.top, cellBox.top), bottom: Math.min(bodyBox.bottom, cellBox.bottom) };
  const priorBands = retainedTextBands(paragraph, range, bounds);
  if (!priorBands) return;
  const terminal = paragraph.dataset.hwpCellTerminal === 'true';
  const advance = Number(paragraph.dataset.hwpCellAdvance), end = Number(paragraph.dataset.hwpCellEnd);
  const font = parseFloat(style.fontSize);
  if (terminal && !(font > 0 && advance > end && end > 0)) return;
  const input = terminal ? paragraph.cloneNode(true) as HTMLParagraphElement : paragraph;
  if (terminal) unwrapHwpCellPrefixes(input);
  const cuts = sourceLineBoundaries(input, reference.ends.slice(0, -1));
  if (!cuts) return;

  const elements = [paragraph, paragraph.nextElementSibling, paragraph.previousElementSibling,
    cell, cell.parentElement, table, body].filter((node): node is Element => !!node);
  const sizes = elements.map(geometry);
  const original = [...paragraph.childNodes], originalStyle = paragraph.style.cssText;
  const restore = () => {
    paragraph.replaceChildren(...original);
    if (terminal) paragraph.style.cssText = originalStyle;
  };
  paragraph.replaceChildren(lineOpportunities(input, cuts, range));
  if (terminal) {
    const placement = hwpCellTerminalLayout(paragraph, font, advance, end);
    if (!placement) { restore(); return; }
    applyHwpCellTerminal(paragraph, placement, advance, end);
  }
  const bands = retainedTextBands(paragraph, range, bounds);
  if (!bands || bands.size !== priorBands.size || [...bands].some(band => !priorBands.has(band)) ||
      elements.some((element, i) => geometry(element) !== sizes[i]) || paragraph.textContent !== text ||
      !matchesStoredWrap(paragraph, text, reference.ends) || !horizontalInkFits(paragraph, range, bounds, scale)) {
    restore();
    return;
  }
  return restore;
}

function fragmentBreaks(group: HTMLElement[], reference: HwpSourceWrap) {
  const text = group.map(p => p.textContent || '').join('');
  if (hwpWrapFingerprint(text) !== reference.fingerprint) return;
  let offset = 0, page: number | undefined;
  const breaks: number[][] = [];
  for (const paragraph of group) {
    const nextPage = Number(paragraph.closest<HTMLElement>('.print-page')?.dataset.page);
    const end = offset + normalizeHwpWrapText(paragraph.textContent || '').length;
    if (!Number.isInteger(nextPage) || page !== undefined && nextPage !== page + 1 ||
        end <= offset || !reference.ends.includes(end) || !paragraph.hasAttribute('data-print-paragraph-fragment') &&
        !paragraph.hasAttribute('data-print-soft-end')) return;
    breaks.push(reference.ends.filter(value => value > offset && value < end).map(value => value - offset));
    offset = end;
    page = nextPage;
  }
  if (offset === reference.ends.at(-1)) return breaks;
}

/** Use recorded HWP line starts only on finished print sheets. A split
 * paragraph must still have the original text and a cut at a source line end. */
function* sourceCellLineSteps(root: HTMLElement) {
  const paragraphs = [...root.querySelectorAll<HTMLElement>(`p[${hwpSourceWrapAttribute}]`)]
    .filter(paragraph => (parseHwpSourceWrap(paragraph.dataset.hwpSourceWrap)?.ends.length ?? 0) > 1);
  if (!paragraphs.length) return;
  const range = document.createRange(), groups = new Map<string, HTMLElement[]>();
  const boxes = [...root.querySelectorAll('.print-page,.print-content,table,tr,td,th,p,h1,h2,h3,h4,h5,h6')];
  const initialBoxes = boxes.map(geometry);
  const rollback: (() => void)[] = [];
  for (const paragraph of paragraphs) {
    yield;
    const group = paragraph.getAttribute(groupAttribute);
    paragraph.removeAttribute(groupAttribute);
    if (group !== null) {
      const members = groups.get(group) ?? [];
      members.push(paragraph); groups.set(group, members);
    }
    const reference = parseHwpSourceWrap(paragraph.getAttribute(hwpSourceWrapAttribute));
    const text = paragraph.textContent || '';
    if (!reference || hwpWrapFingerprint(text) !== reference.fingerprint ||
      normalizeHwpWrapText(text).length !== reference.ends.at(-1)) continue;
    const ends = reference.ends.slice(0, -1);
    const restore = tryWrap(paragraph, reference, ends, range);
    if (restore) rollback.push(restore);
    else if (restore === false && reference.cellWidth && !reference.columnCellHeight) {
      const alternate = tryWrap(paragraph, reference, ends, range, false, true);
      if (alternate) rollback.push(alternate);
    }
  }
  for (const group of groups.values()) {
    if (group.length < 2) continue;
    const value = group[0].getAttribute(hwpSourceWrapAttribute);
    const reference = parseHwpSourceWrap(value);
    if (!reference || !group.every(p => p.getAttribute(hwpSourceWrapAttribute) === value)) continue;
    const breaks = fragmentBreaks(group, reference);
    if (!breaks) continue;
    const applied: (() => void)[] = [];
    let valid = true;
    for (const [index, paragraph] of group.entries()) {
      yield;
      const restore = tryWrap(paragraph, reference, breaks[index], range, true);
      if (restore === false) { valid = false; break; }
      if (restore) applied.push(restore);
    }
    if (valid) rollback.push(...applied);
    else applied.forEach(restore => restore());
  }
  if (rollback.length && boxes.some((element, index) => geometry(element) !== initialBoxes[index])) rollback.forEach(restore => restore());
}

/** The print copy stays hidden until the full pass, including rollback. */
export function layoutHwpSourceCellLines(root: HTMLElement, checkpoint?: LayoutCheckpoint) {
  const steps = sourceCellLineSteps(root);
  if (!checkpoint) {
    for (const _ of steps) { /* Finish before returning to synchronous callers. */ }
    return;
  }
  return (async () => {
    for (const _ of steps) {
      const pause = checkpoint();
      if (pause) await pause;
    }
  })();
}

/** Keep late, same-height cell wraps separate from the initial source
 * layout so a changed peer cannot undo an earlier, unrelated restoration. */
export function layoutHwpFittedCellLineEnds(root: HTMLElement) {
  const range = document.createRange();
  for (const table of root.querySelectorAll<HTMLTableElement>('.print-content table[data-hwp-centered-borders="continuation"]')) {
    if (!fixedSingleLineTable(table)) continue;
    const paragraphs = table.querySelectorAll<HTMLParagraphElement>(`td > p[${hwpSourceWrapAttribute}]`);
    if (!paragraphs.length) continue;
    const boxes = [table, ...table.querySelectorAll('tr,td,th,p,h1,h2,h3,h4,h5,h6')];
    const initialBoxes = boxes.map(geometry);
    const rollback: (() => void)[] = [];
    for (const paragraph of paragraphs) {
      const reference = parseHwpSourceWrap(paragraph.getAttribute(hwpSourceWrapAttribute));
      const text = paragraph.textContent ?? '';
      if (!reference?.cellWidth || reference.columnCellHeight || reference.ends.length < 2 ||
          reference.fingerprint !== hwpWrapFingerprint(text) || reference.ends.at(-1) !== normalizeHwpWrapText(text).length) continue;
      const restore = restoreFittedCellLineEnds(paragraph, reference, range);
      if (restore) rollback.push(restore);
    }
    if (rollback.length && boxes.some((element, index) => geometry(element) !== initialBoxes[index]))
      rollback.reverse().forEach(restore => restore());
  }
}
