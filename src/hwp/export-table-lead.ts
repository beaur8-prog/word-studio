import type { HwpDocument } from '@rhwp/core';
import { strFromU8, strToU8, zipSync } from 'fflate';
import { measuredHwpLineInk } from '../hwp-line-ink';
import { paragraphLineOffsets } from '../paragraph-fragments';
import { hwpWrapFingerprint, parseHwpSourceWrap } from '../hwp-text-advances';
import { elementRanges } from './markup-ranges';
import type { PaperHeaderTable, RasterTablePair, TablePlacement } from './export-table-placement';
import type { PastedEmptyParagraphs, TerminalEmptyFonts } from './export-empty-paragraph-format';
import type { AuthoredTerminalWrap, TerminalPageBreak } from './export-terminal-page-break';
import type { AdjacentNativeTables } from './export-adjacent-tables';
import type { EmptyTableRows } from './export-empty-table-rows';
import { boundedNativeParts, nativeBodyRuns, nativeGlyphBaselines, nativePaintedControls, sameNativeLineContent, type NativeBodyRun } from './export-binary-pages';

type TableLead = { sheet: number; literal: string; width: number; height: number; rows: number };
export type NativePrintPlan = { pages: number; tableLeads: readonly TableLead[];
  continuations?: readonly TablePlacement[]; adjacentTables?: readonly AdjacentNativeTables[]; paperHeaderTable?: PaperHeaderTable;
  rasterTablePairs?: readonly RasterTablePair[]; terminalEmptyFonts?: readonly TerminalEmptyFonts[];
  terminalPageBreaks?: readonly TerminalPageBreak[]; authoredTerminalWraps?: readonly AuthoredTerminalWrap[]; emptyTableRows?: readonly EmptyTableRows[];
  emptyParagraphs?: PastedEmptyParagraphs };
const round = (pixels: number) => Math.round(pixels * 75);
const nonspace = (text: string) => text.replace(/\s/gu, '');
const same = (left: unknown, right: unknown) => JSON.stringify(left) === JSON.stringify(right);
const epsilon = .15;

/** Collect one-line heading and last-table pairs from the actual print layout
 * before the HTML bridge detaches each sheet. */
export function measuredTableLead(body: Element, sheet: number): TableLead | undefined {
  const table = body.lastElementChild, lead = table?.previousElementSibling;
  if (!(table instanceof HTMLTableElement) || !(lead instanceof HTMLParagraphElement) ||
      table.matches('.table-overlay-active') || lead.hasAttribute('data-print-paragraph-fragment') ||
      lead.querySelector(':not(span),[data-flow-widget],[data-docx-tab]')) return;
  const literal = lead.textContent ?? '';
  if (!nonspace(literal) || literal.length > 512 || /[\t\r\n\u0000-\u001f\ufffc]/u.test(literal)) return;
  const frame = body.getBoundingClientRect(), bounds = table.getBoundingClientRect();
  if (bounds.height < frame.height / 2 || bounds.width <= 0 || bounds.bottom > frame.bottom + epsilon ||
      !table.rows.length) return;
  // A confirmed substitute can have a font box taller than its visible
  // ink. This path already verified the source line during pagination.
  const source = lead.dataset.hwpPrintSourceFlow === 'true' ? parseHwpSourceWrap(lead.dataset.hwpSourceWrap) : undefined;
  const ink = source?.singleBodyFallback && source.origin?.page === sheet + 1 &&
    source.bodyWidth !== undefined && Math.abs(frame.width - source.bodyWidth) <= epsilon && source.fingerprint === hwpWrapFingerprint(literal) &&
    !paragraphLineOffsets(lead).length
    ? measuredHwpLineInk(lead, 1, parseFloat(getComputedStyle(lead).fontSize), true) : null;
  if (ink && (ink.top < frame.top - epsilon || ink.bottom > bounds.top + epsilon)) return;
  const walker = document.createTreeWalker(lead, NodeFilter.SHOW_TEXT), range = document.createRange();
  let top: number | undefined;
  while (walker.nextNode()) {
    const node = walker.currentNode;
    if (!nonspace(node.textContent ?? '')) continue;
    range.selectNodeContents(node);
    const rects = [...range.getClientRects()].filter(rect => rect.width > 0 && rect.height > 0);
    if (!rects.length) return;
    top ??= rects[0].top;
    const lineTop = top;
    if (rects.some(rect => Math.abs(lineTop - rect.top) > epsilon || rect.left < frame.left - epsilon ||
        rect.right > frame.right + epsilon || !ink && (rect.top < frame.top - epsilon || rect.bottom > bounds.top + epsilon))) return;
  }
  if (top === undefined) return;
  return { sheet, literal, width: bounds.width, height: bounds.height, rows: table.rows.length };
}

function measuredOneLine(source: HwpDocument, page: number, section: number, index: number, runs: NativeBodyRun[]) {
  if (runs.length !== 2 || runs[0].charStart || runs[1].charStart <= 0 || runs[0].y >= runs[1].y ||
      runs[0].charShapeId !== runs[1].charShapeId) return;
  const layout = JSON.parse(source.getParagraphLineMetrics(section,index,'[]'));
  const style = JSON.parse(source.getParaPropertiesAt(section,index));
  if (layout.stored || layout.controlCount || style.lineSpacingType !== 'Percent' || style.alignment !== 'justify' ||
      style.marginLeft || style.marginRight || style.indent >= 0 ||
      JSON.parse(source.getColumnDef(section)).columnCount !== 1) return;
  const { fontSize: font } = JSON.parse(source.getCharPropertiesAt(section,index,0));
  if (!Number.isSafeInteger(font) || font <= 0 || runs.some(run => Math.abs(run.h - font / 75) > .065)) return;
  const paper = JSON.parse(source.getPageInfo(page));
  const baselineAt = nativeGlyphBaselines(source.renderPageSvg(page));
  const first = baselineAt(runs[0]), second = baselineAt(runs[1]);
  if (first === undefined || second === undefined) return;
  const ascent = round(first - runs[0].y), origin = round(first - paper.marginTop - paper.marginHeader) - ascent;
  const width = round(paper.width - paper.marginLeft - paper.marginRight);
  const gap = Math.round(font * (style.lineSpacing / 100 - 1));
  if (![ascent,origin,width,gap].every(Number.isSafeInteger) || ascent <= 0 || origin < 0 ||
      width <= 0 || gap < 0 || Math.abs(round(second - first) - font - gap) > 1) return;
  return `<hp:linesegarray><hp:lineseg textpos="0" vertpos="${origin}" vertsize="${font}" textheight="${font}" baseline="${ascent}" spacing="${gap}" horzpos="0" horzsize="${width}" flags="393216"/></hp:linesegarray>`;
}

/** The table and its nested cells may translate vertically together. All
 * other attributes, including the run advances and object widths, are fixed. */
function sameTranslatedY(before: unknown, after: unknown, delta: number): boolean {
  if (before === after) return true;
  if (!before || !after || typeof before !== 'object' || typeof after !== 'object') return false;
  const left = before as Record<string, unknown>, right = after as Record<string, unknown>;
  const keys = Object.keys(left);
  return same(keys,Object.keys(right)) && keys.every(key => {
    const a = left[key], b = right[key];
    return key === 'y' && typeof a === 'number' && typeof b === 'number'
      ? Math.abs(a + delta - b) <= epsilon : sameTranslatedY(a,b,delta);
  });
}

function nativeLossReport(doc: HwpDocument, format: 'hwp' | 'hwpx') {
  const output = format === 'hwp' ? doc.exportHwpWithReport() : doc.exportHwpxWithReport();
  try { return output.contentLoss(); }
  finally { output.free(); }
}

function tryTableLead(Core: typeof HwpDocument, source: HwpDocument, plan: NativePrintPlan, proof: TableLead) {
  const { sheet: page, literal, width, height, rows } = proof;
  if (!Number.isInteger(page) || page < 0 || page + 1 >= plan.pages ||
      typeof literal !== 'string' || !literal || literal.length > 512 ||
      ![width,height,rows].every(value => Number.isFinite(value) && value > 0)) return;
  const current = nativeBodyRuns(source,page), spill = nativePaintedControls(source,page + 1);
  if (spill.length !== 1 || spill[0].type !== 'table') return;
  const [table] = spill, { secIdx: section, paraIdx: index } = table, lead = index - 1;
  const dimensions = JSON.parse(source.getTableDimensions(section,index,table.controlIdx));
  if (!table.cells?.length || dimensions.rowCount !== rows || (table.rowStart ?? 0) !== 0 || (table.rowEnd ?? rows) !== rows ||
      Math.abs(table.w - width) > 2 || Math.abs(table.h - height) > 2 ||
      lead < 0 || index + 1 >= source.getParagraphCount(section) ||
      !JSON.parse(source.getParaPropertiesAt(section,index + 1)).pageBreakBefore ||
      source.getTextRange(section,lead,0,0xffffffff) !== literal) return;
  const isLead = (run: NativeBodyRun) => run.secIdx === section && run.paraIdx === lead && !run.cellPath?.length;
  const inTable = (run: NativeBodyRun) => run.secIdx === section && run.parentParaIdx === index;
  const beforeRuns = current.filter(isLead), spilledRuns = nativeBodyRuns(source,page + 1);
  if (!spilledRuns.length || !spilledRuns.every(inTable) ||
      nonspace(beforeRuns.map(run => run.text).join('')) !== nonspace(literal)) return;
  const cache = measuredOneLine(source,page,section,lead,beforeRuns);
  if (!cache) return;
  const files = boundedNativeParts(source.exportHwpx()), file = `Contents/section${section}.xml`;
  const xml = strFromU8(files[file]), owners = elementRanges(xml,'hp:p'), owner = owners[lead];
  if (!owner || owners.length !== source.getParagraphCount(section) || /<hp:(?:ctrl|tbl|linesegarray)\b/.test(owner.inner)) return;
  const close = owner.end - '</hp:p>'.length;
  files[file] = strToU8(xml.slice(0,close) + cache + xml.slice(close));
  const candidate = new Core(zipSync(files,{level:1}));
  let accepted = false;
  try {
    if (candidate.pageCount() !== plan.pages || candidate.getTextRange(section,lead,0,0xffffffff) !== literal ||
        !sameNativeLineContent(source,candidate,file,lead)) return;
    const painted = nativeBodyRuns(candidate,page), moved = painted.filter(inTable), leads = painted.filter(isLead);
    const afterControls = nativePaintedControls(candidate,page), originalControls = nativePaintedControls(source,page);
    if (leads.length !== 1 || afterControls.length !== originalControls.length + 1) return;
    const [line] = leads, [first] = beforeRuns, placed = afterControls.at(-1)!;
    const paper = JSON.parse(candidate.getPageInfo(page));
    if (nonspace(line.text) !== nonspace(literal) ||
        !same([line.x,line.y,line.h,line.charShapeId],[first.x,first.y,first.h,first.charShapeId]) ||
        line.x < paper.marginLeft - epsilon || line.x + line.w > paper.width - paper.marginRight + epsilon ||
        line.y + line.h > placed.y || placed.y + placed.h > paper.height - paper.marginBottom - paper.marginFooter + epsilon ||
        placed.x < paper.marginLeft - epsilon || placed.x + placed.w > paper.width - paper.marginRight + epsilon ||
        !same(originalControls,afterControls.slice(0,-1)) ||
        !same(current.filter(run => !isLead(run)),painted.filter(run => !isLead(run) && !inTable(run)))) return;
    const delta = placed.y - table.y;
    if (!sameTranslatedY(table,placed,delta) || !sameTranslatedY(spilledRuns,moved,delta)) return;

    // The patch can change one page only. Ignore repeated furniture, whose
    // odd/even choice is supposed to follow the corrected physical index.
    for (let i = 0; i < plan.pages; i++) {
      if (i === page) continue;
      const previous = i > page ? i + 1 : i;
      if (!same(nativeBodyRuns(source,previous),nativeBodyRuns(candidate,i)) ||
          !same(nativePaintedControls(source,previous),nativePaintedControls(candidate,i))) return;
    }
    accepted = true;
    return candidate;
  } finally {
    if (!accepted) candidate.free();
  }
}

export function restoreNativeTableLead(Core: typeof HwpDocument, source: HwpDocument,
  plan: NativePrintPlan | undefined, format: 'hwp' | 'hwpx' = 'hwpx'): HwpDocument {
  if (!plan?.tableLeads.length || plan.pages < 2 || plan.tableLeads.length > plan.pages ||
      source.pageCount() !== plan.pages + 1) return source;
  let reopened: HwpDocument | undefined;
  try {
    // Round-trip the normal save first. The mutable paste document and its
    // reopened representation can report different cached glyph positions.
    reopened = new Core(source.exportHwpx());
    if (reopened.pageCount() !== plan.pages + 1) return source;
    for (const proof of plan.tableLeads) {
      const candidate = tryTableLead(Core,reopened,plan,proof);
      if (!candidate) continue;
      let retained = false;
      try {
        retained = nativeLossReport(source,format) === nativeLossReport(candidate,format);
        if (retained) return candidate;
      } finally {
        if (!retained) candidate.free();
      }
    }
  } catch {
    // Without the same content and physical proof, keep the normal save.
  } finally { reopened?.free(); }
  return source;
}
