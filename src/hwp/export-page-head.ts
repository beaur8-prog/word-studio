import type { HwpDocument } from '@rhwp/core';
import { cssPixels } from '../hwp-cell-line';
import { hwpWrapFingerprint, parseHwpSourceWrap, type HwpAuthoredPageHead } from '../hwp-text-advances';
import { nativeBodyRuns, nativePaintedControls } from './export-binary-pages';

export type NativePageHead = HwpAuthoredPageHead & { head: string; body: string };
const close = (a: number, b: number) => Math.abs(a - b) <= .2;
const visible = (literal: string) => literal.replace(/\s/gu, '');

/** The browser puts a source page's heading and its small spacer on the
 * preceding sheet. Only a generated boundary can separate the two blocks. */
export function measuredNativePageHead(previous: HTMLElement, sheet: HTMLElement): NativePageHead | undefined {
  if (Number(sheet.dataset.page) === 9)
    console.info('QA_PAGE_HEAD_SHEET ' + JSON.stringify({ keepBlank: sheet.dataset.keepBlank, section: sheet.dataset.sectionFirst }));
  if (sheet.dataset.keepBlank || sheet.dataset.sectionFirst === 'true') return;
  const before = previous.querySelector(':scope > .print-content'), after = sheet.querySelector(':scope > .print-content');
  const spacer = before?.lastElementChild, lead = spacer?.previousElementSibling, body = after?.firstElementChild;
  if (Number(sheet.dataset.page) === 9)
    console.info('QA_PAGE_HEAD_DOM ' + JSON.stringify({ before: [...(before?.children ?? [])].slice(-4)
      .map(node=>[node.tagName,node.childNodes.length,node.textContent?.trim().length]),
      after: [...(after?.children ?? [])].slice(0,2).map(node=>[node.tagName,node.childNodes.length,node.textContent?.trim().length]) }));
  if (!(lead instanceof HTMLParagraphElement) || !(spacer instanceof HTMLParagraphElement) || spacer.childNodes.length ||
      !(body instanceof HTMLParagraphElement) || [lead, body].some(p => p.hasAttribute('data-print-paragraph-fragment') ||
        p.querySelector('img,svg,br,[data-note-reference],[data-document-field],[data-docx-tab]'))) return;
  const wrap = parseHwpSourceWrap(lead.dataset.hwpSourceWrap), source = wrap?.authoredPageHead;
  const a = lead.textContent ?? '', b = body.textContent ?? '';
  const spacerSize = (cssPixels(spacer.style.fontSize) ?? cssPixels(getComputedStyle(spacer).fontSize) ?? 0) * 75;
  if (Number(sheet.dataset.page) === 9)
    console.info('QA_PAGE_HEAD_PROOF ' + JSON.stringify({ source: !!source, sourcePage: source?.page,
      head: wrap && hwpWrapFingerprint(a) === wrap.fingerprint, next: source && hwpWrapFingerprint(b) === source.next,
      spacerSize, target: source?.spacer }));
  if (source && source.page === Number(sheet.dataset.page) &&
      hwpWrapFingerprint(a) === wrap.fingerprint && hwpWrapFingerprint(b) === source.next &&
      Math.abs(spacerSize - source.spacer) <= 2) return { ...source, head: a, body: b };
}

function bodyOwner(doc: HwpDocument, page: number, literal: string) {
  const run = nativeBodyRuns(doc, page).find(run => !run.cellPath?.length && Number.isInteger(run.paraIdx) && !run.charStart && visible(run.text) &&
    doc.getTextRange(run.secIdx, run.paraIdx, 0, 0xffffffff) === literal);
  if (run) return run;
}

function paintedPayload(doc: HwpDocument) {
  const ink = new Map<string, string>();
  const objects: string[] = [];
  for (let page = 0; page < doc.pageCount(); page++) {
    for (const line of nativeBodyRuns(doc, page)) {
      if (!Number.isInteger(line.paraIdx)) continue;
      const key = JSON.stringify([line.secIdx, line.paraIdx, line.parentParaIdx, line.cellPath]);
      const content = visible(line.text);
      if (content) ink.set(key, (ink.get(key) ?? '') + content);
    }
    for (const shape of nativePaintedControls(doc, page)) {
      objects.push(JSON.stringify([shape.type, shape.secIdx, shape.paraIdx, shape.controlIdx, shape.cellPath, shape.w, shape.h]));
    }
  }
  return JSON.stringify([[...ink].sort(), objects.sort()]);
}

export function restoreNativePageHead(Core: typeof HwpDocument, doc: HwpDocument,
  proofs: readonly NativePageHead[] = [], printPages = 0): HwpDocument {
  console.info('QA_PAGE_HEAD_COMPOSITION ' + JSON.stringify({ proofs: proofs.map(p=>p.page), pages: doc.pageCount(), printPages }));
  if (!proofs.length || doc.pageCount() <= printPages) return doc;
  for (const proof of proofs) {
    const page = proof.page - 1;
    if (page < 1 || page + 1 >= doc.pageCount()) continue;
    let candidate: HwpDocument | undefined;
    try {
      const lead = bodyOwner(doc, page, proof.head), next = bodyOwner(doc, page + 1, proof.body);
      console.info('QA_PAGE_HEAD_NATIVE_OWNERS ' + JSON.stringify({ page,
        lead: lead && [lead.paraIdx,lead.y], next: next && [next.paraIdx,next.y],
        objects: nativePaintedControls(doc, page).length, runs: nativeBodyRuns(doc, page).length }));
      if (!lead || !next || !close(lead.y, proof.y) || next.secIdx !== lead.secIdx || next.paraIdx !== lead.paraIdx + 2 ||
          nativePaintedControls(doc, page).length || nativeBodyRuns(doc, page).some(run => Number.isInteger(run.paraIdx) && visible(run.text) &&
            (run.secIdx !== lead.secIdx || run.paraIdx !== lead.paraIdx || !close(run.y, lead.y)))) continue;
      const { secIdx: section, paraIdx: paragraph } = next, spacer = paragraph - 1;
      const paper = JSON.parse(doc.getPageInfo(page + 1));
      const line = JSON.parse(doc.getParaPropertiesAt(section, paragraph));
      console.info('QA_PAGE_HEAD_NATIVE_SPACER ' + JSON.stringify({ page, font: JSON.parse(doc.getCharPropertiesAt(section, spacer, 0)).fontSize,
        spacerPage: JSON.parse(doc.getPageOfPosition(section, spacer)).page, top: paper.marginTop + paper.marginHeader,
        breakBefore: line.pageBreakBefore, length: doc.getParagraphLength(section, spacer) }));
      if (!close(next.y, paper.marginTop + paper.marginHeader) || doc.getParagraphLength(section, spacer) ||
          Math.abs(JSON.parse(doc.getCharPropertiesAt(section, spacer, 0)).fontSize - proof.spacer) > 2 ||
          JSON.parse(doc.getParagraphLineMetrics(section, spacer, '[]')).controlCount ||
          JSON.parse(doc.getPageOfPosition(section, spacer)).page !== page ||
          !JSON.parse(doc.getParaPropertiesAt(section, paragraph)).pageBreakBefore) continue;
      candidate = new Core(doc.exportHwpx());
      const before = candidate.pageCount();
      const edit = JSON.parse(candidate.applyParaFormat(section, paragraph, JSON.stringify({ pageBreakBefore: false }))).ok;
      console.info('QA_PAGE_HEAD_NATIVE_APPLY ' + JSON.stringify({ current: doc.pageCount(), before, after: candidate.pageCount(), edit }));
      if (before !== doc.pageCount() || !edit || candidate.pageCount() !== doc.pageCount() - 1) continue;
      const sameLead = bodyOwner(candidate, page, proof.head), sameNext = bodyOwner(candidate, page, proof.body);
      const bounds = JSON.parse(candidate.getPageInfo(page));
      console.info('QA_PAGE_HEAD_NATIVE_READBACK ' + JSON.stringify({ lead: sameLead && [sameLead.x,sameLead.y],
        next: sameNext && [sameNext.x,sameNext.y], expected: [proof.y,proof.nextY],
        payload: paintedPayload(candidate) === paintedPayload(doc) }));
      if (!sameLead || !sameNext || !close(sameLead.y, proof.y) || !close(sameNext.y, proof.nextY) ||
          !close(sameLead.x, lead.x) || !close(sameNext.x, next.x) ||
          paintedPayload(candidate) !== paintedPayload(doc) ||
          nativeBodyRuns(candidate, page).some(run => Number.isInteger(run.paraIdx) && visible(run.text) &&
            (run.x < 0 || run.x + run.w > bounds.width + .15 ||
              run.y < proof.y - .15 || run.y + run.h > bounds.footerArea.y + .15))) continue;
      const accepted = candidate;
      candidate = undefined;
      return accepted;
    } catch { console.info('QA_PAGE_HEAD_NATIVE_EXCEPTION'); }
    finally { candidate?.free(); }
  }
  return doc;
}
