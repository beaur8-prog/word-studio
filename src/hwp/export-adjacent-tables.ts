import type { HwpDocument } from '@rhwp/core';
import { cssPixels } from '../hwp-cell-line';
import { hiddenEmptyLineClass } from '../page-empty-lines';
import { hwpWrapFingerprint } from '../hwp-text-advances';
import { hwpInlineTableTop, tableAnchor } from '../table-anchor-data';
import { nativeBodyRuns, nativePaintedControls, nativePaintedPayload, type NativePaintedControl } from './export-binary-pages';

type TableProof = { signature: string; width: number; height: number; rows: number; columns: number };
export type AdjacentNativeTables = { label: TableProof; main: TableProof; spacerFont: number };
const text = (value: string) => value.replace(/\s/gu, '');
const close = (left: number, right: number, tolerance = .16) => Math.abs(left - right) <= tolerance;

function tableProof(table: HTMLTableElement): TableProof | undefined {
  const literal = text(table.textContent ?? '');
  const columns = table.querySelectorAll(':scope > colgroup > col').length;
  if (!literal || !columns || literal.length > 16000 || table.style.zoom ||
      table.querySelector('table,img,svg,[data-note-reference],[data-document-textbox],[data-word-overlap]')) return;
  const {width,height} = table.getBoundingClientRect();
  if (width > 0 && height > 0)
    return { signature: hwpWrapFingerprint(literal), rows: table.rows.length, columns, width, height };
}

/** The browser moved the large table at a measured boundary, not an authored
 * break. The native writer may place its preceding label on a new sheet too. */
export function measuredAdjacentNativeTables(previous: Element, sheet: HTMLElement): AdjacentNativeTables | undefined {
  if (sheet.dataset.keepBlank || sheet.dataset.sectionFirst === 'true') return;
  const before = previous.querySelector(':scope > .print-content'), after = sheet.querySelector(':scope > .print-content');
  const spacer = before?.lastElementChild, label = spacer?.previousElementSibling, table = after?.firstElementChild;
  if (!before || !after || !(spacer instanceof HTMLParagraphElement) || spacer.childNodes.length ||
      !spacer.classList.contains(hiddenEmptyLineClass) || !(label instanceof HTMLTableElement) ||
      !(table instanceof HTMLTableElement) || hwpInlineTableTop(label.dataset.hwpInlineTableTop) === undefined) return;
  const anchor = tableAnchor(table.dataset.tableAnchor);
  if (anchor?.reference !== 'paragraph' || anchor.horizontal !== 'margin' || anchor.overlay || anchor.wrap ||
      table.hasAttribute('data-print-native-continued-table')) return;
  const a = tableProof(label), b = tableProof(table), frame = after.getBoundingClientRect(), prior = before.getBoundingClientRect();
  // The print-only hidden line uses zero computed leading. The native
  // paragraph still retains the authored inline font and advance.
  const font = cssPixels(spacer.style.fontSize), advance = cssPixels(spacer.style.lineHeight);
  if (!font || !advance || !a || !b || a.rows !== 1 || a.columns !== 1 || b.rows < 2 || b.columns < 2 ||
      a.height >= frame.height / 8 || (a.width < frame.width / 2 || b.width < frame.width / 2) ||
      b.height < frame.height / 2 || a.height + b.height + advance + anchor.y + anchor.bottom > frame.height ||
      !(font > 0 && font < 14 && advance > 0 && advance < 24) ||
      !close(prior.width, frame.width) || !close(prior.height, frame.height) ||
      label.getBoundingClientRect().bottom > prior.bottom + .5) return;
  return { label: a, main: b, spacerFont: Math.round(font * 75) };
}

function tableMatches(doc: HwpDocument, table: NativePaintedControl, page: number, proof: TableProof) {
  if (table.type !== 'table' || table.controlIdx || table.cellPath?.length ||
      !close(table.w, proof.width, 1.25) || !close(table.h, proof.height, 5)) return false;
  const {rowCount,colCount} = JSON.parse(doc.getTableDimensions(table.secIdx, table.paraIdx, table.controlIdx));
  return rowCount === proof.rows && colCount === proof.columns && hwpWrapFingerprint(text(nativeBodyRuns(doc, page)
    .filter(run => run.secIdx === table.secIdx && run.parentParaIdx === table.paraIdx && run.cellPath?.length === 1)
    .map(run => run.text).join(''))) === proof.signature;
}

function findPair(doc: HwpDocument, proof: AdjacentNativeTables) {
  const matches = [];
  for (let page = 1; page < doc.pageCount(); page++) {
    const before = nativePaintedControls(doc, page - 1), after = nativePaintedControls(doc, page);
    if (before.length !== 1 || after.length !== 1) continue;
    const [label] = before, [main] = after;
    if (label.secIdx !== main.secIdx || label.paraIdx + 2 !== main.paraIdx ||
        !tableMatches(doc,label,page - 1,proof.label) || !tableMatches(doc,main,page,proof.main)) continue;
    const {secIdx:section,paraIdx:paragraph} = main, spacer = paragraph - 1;
    const span = JSON.parse(doc.getParagraphLineMetrics(section, spacer, '[]'));
    const props = JSON.parse(doc.getParaPropertiesAt(section, paragraph));
    const position = JSON.parse(doc.getPageOfPosition(section, spacer));
    const font = JSON.parse(doc.getCharPropertiesAt(section, spacer, 0));
    const paper = JSON.parse(doc.getPageInfo(page));
    if (!props.pageBreakBefore || span.controlCount || doc.getParagraphLength(section, spacer) ||
        position.page !== page - 1 || font.fontSize !== proof.spacerFont ||
        !Number.isFinite(paper.footerArea?.y) || label.y + label.h + main.h >= paper.footerArea.y ||
        nativeBodyRuns(doc, page - 1).some(run => text(run.text) && !run.cellPath?.length)) continue;
    matches.push({page, label, main});
  }
  if (matches.length === 1) return matches[0];
}

export function restoreAdjacentNativeTables(Core: typeof HwpDocument, doc: HwpDocument,
  proofs: readonly AdjacentNativeTables[] = [], printPages = 0): HwpDocument {
  if (!proofs.length || doc.pageCount() <= printPages) return doc;
  for (const proof of proofs) {
    let candidate: HwpDocument | undefined;
    try {
      const pair = findPair(doc, proof);
      if (!pair) continue;
      const {page,label,main} = pair, section = main.secIdx;
      candidate = new Core(doc.exportHwpx());
      if (candidate.pageCount() !== doc.pageCount() || !findPair(candidate, proof)) continue;
      const edit = JSON.parse(candidate.applyParaFormat(section, main.paraIdx, JSON.stringify({pageBreakBefore:false})));
      if (!edit.ok || candidate.pageCount() > doc.pageCount() || candidate.pageCount() < doc.pageCount() - 1) continue;
      const shapes = nativePaintedControls(candidate, page - 1);
      if (shapes.length !== 2 || shapes[0].paraIdx !== label.paraIdx || shapes[1].paraIdx !== main.paraIdx) continue;
      const [title, table] = shapes, paper = JSON.parse(candidate.getPageInfo(page - 1));
      if (!close(title.y,label.y) || !close(title.x,label.x) || !close(table.x,main.x) ||
          !Number.isFinite(paper.footerArea?.y) || table.y < title.y + title.h || table.y + table.h > paper.footerArea.y + .16 ||
          !tableMatches(candidate,table,page - 1,proof.main) || nativePaintedPayload(candidate) !== nativePaintedPayload(doc)) continue;
      const ink = nativeBodyRuns(candidate, page - 1);
      if (ink.some(run => text(run.text) && (!run.cellPath?.length || run.y < paper.marginTop + paper.marginHeader - .16 ||
          run.y + run.h > paper.footerArea.y + .16 || run.x < 0 || run.x + run.w > paper.width + .16))) continue;
      const accepted = candidate;
      candidate = undefined;
      return accepted;
    } catch { /* The normal measured layout remains valid if the proof fails. */ }
    finally { candidate?.free(); }
  }
  return doc;
}
