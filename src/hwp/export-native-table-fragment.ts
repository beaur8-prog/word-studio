import type { HwpDocument } from '@rhwp/core';
import { strFromU8, strToU8, zipSync } from 'fflate';
import { hwpWrapFingerprint, normalizeHwpWrapText } from '../hwp-text-advances';
import { tableAnchor, type AuthoredNativeTableFragment } from '../table-anchor-data';
import { boundedNativeParts, nativeBodyRuns, nativePaintedControls, nativePaintedPayload, type NativeBodyRun, type NativePaintedControl } from './export-binary-pages';
import { elementRanges } from './markup-ranges';

export function measuredNativeTableFragment(body: Element, sheet: HTMLElement): AuthoredNativeTableFragment | undefined {
  const table = body.lastElementChild;
  if (!(table instanceof HTMLTableElement) || table.querySelector('table,img,svg,[data-document-note],[data-note-reference],[data-document-textbox]') ||
      body.children.length < 2 || body.children.length > 3 || [...body.children].some(child => child.tagName !== 'TABLE')) return;
  const proof = tableAnchor(table.dataset.tableAnchor)?.nativeFirstFragment;
  if (!proof || Number(sheet.dataset.page) !== proof.page || table.rows.length !== proof.rows ||
      [...table.rows].some(row => row.cells.length !== proof.columns || [...row.cells].some(cell => cell.colSpan !== 1 || cell.rowSpan !== 1))) return;
  const box = table.getBoundingClientRect(), frame = body.getBoundingClientRect(), paper = sheet.getBoundingClientRect();
  if (Math.abs(frame.width - proof.bodyWidth) <= .2 && Math.abs(paper.height - proof.paperHeight) <= .2 &&
      Math.abs(box.width - proof.width) <= 1 && Math.abs(box.left - paper.left - proof.x) <= .2 &&
      box.bottom > frame.bottom - 16 && box.bottom <= frame.bottom + .2) return proof;
}

const sameOwner = (a: NativePaintedControl, b: NativePaintedControl) => a.secIdx === b.secIdx &&
  a.paraIdx === b.paraIdx && a.controlIdx === b.controlIdx;
const close = (a: number, b: number) => Math.abs(a - b) <= .16;
const patch = (tag: string, values: Record<string, number | string>) => Object.entries(values).reduce((text, [key, value]) =>
  text.replace(new RegExp(`\\b${key}="[^"]*"`), `${key}="${value}"`), tag);

function ownInk(doc: HwpDocument, page: number, control: NativePaintedControl) {
  return nativeBodyRuns(doc, page).filter(run => normalizeHwpWrapText(run.text) &&
    run.secIdx === control.secIdx && run.parentParaIdx === control.paraIdx && run.cellPath?.length === 1 &&
    (run.cellPath[0] as { controlIndex: number }).controlIndex === control.controlIdx);
}

function visibleRight(run: NativeBodyRun) {
  const end = run.text.trimEnd().length, advances = run.charX;
  if (end > 0 && end < run.text.length && advances?.length === run.text.length + 1 &&
      advances.every((advance, index) => Number.isFinite(advance) && advance >= (index ? advances[index - 1] : 0))) {
    // The layout box includes the advance of trailing source whitespace.
    // The last visible glyph can fit while that non-printing advance does not.
    return run.x + advances[end];
  }
  return run.x + run.w;
}

function restoreFragment(Core: typeof HwpDocument, doc: HwpDocument, proof: AuthoredNativeTableFragment) {
  const page = proof.page - 1;
  if (page + 1 >= doc.pageCount()) return;
  const [target, ...peers] = nativePaintedControls(doc, page + 1);
  if (!target || peers.length || target.type !== 'table' || target.cellPath?.length || target.controlIdx !== 0 ||
      Math.abs(target.w - proof.width) > 1) return;
  const paper = JSON.parse(doc.getPageInfo(page));
  if (!close(paper.height, proof.paperHeight) || !close(paper.bodyRight - paper.bodyLeft, proof.bodyWidth) ||
      proof.x < paper.bodyLeft || proof.x + target.w > paper.bodyRight || proof.y + target.h > paper.footerArea.y + .08) return;
  const current = nativePaintedControls(doc, page);
  if (!current.length || current.length > 2 || current.some(item => item.type !== 'table' ||
      item.y + item.h > proof.y - .15)) return;
  if (nativeBodyRuns(doc, page).some(run => normalizeHwpWrapText(run.text) &&
    (!run.cellPath?.length || run.y + run.h >= proof.y - .15))) return;
  const { secIdx: section, paraIdx: paragraph, controlIdx: control } = target;
  const shape = JSON.parse(doc.getTableDimensions(section, paragraph, control));
  const attributes = JSON.parse(doc.getTableProperties(section, paragraph, control));
  const paragraphFormat = JSON.parse(doc.getParaPropertiesAt(section, paragraph));
  const line = JSON.parse(doc.getParagraphLineMetrics(section, paragraph, '[]'));
  const ink = ownInk(doc, page + 1, target);
  const visible = nativeBodyRuns(doc, page + 1).filter(run => normalizeHwpWrapText(run.text));
  if (shape.rowCount !== proof.rows || shape.colCount !== proof.columns || !attributes.treatAsChar ||
      attributes.hasCaption || attributes.pageBreak !== 0 || paragraphFormat.pageBreakBefore ||
      line.controlCount !== 1 || doc.getParagraphLength(section, paragraph) || visible.length !== ink.length ||
      hwpWrapFingerprint(ink.map(run => run.text).join('')) !== proof.fingerprint) return;
  const parts = boundedNativeParts(doc.exportHwpx());
  const name = `Contents/section${section}.xml`, bytes = parts[name];
  if (!bytes) return;
  const xml = strFromU8(bytes), range = elementRanges(xml, 'hp:p')[paragraph];
  if (!range) return;
  const [table, extra] = elementRanges(range.inner, 'hp:tbl');
  if (!table || extra || !/<hp:pos\b[^>]*treatAsChar="1"/.test(table.inner)) return;
  const prefix = /^<hp:sz\b[^>]*\/><hp:pos\b[^>]*\/><hp:outMargin\b[^>]*\/>/.exec(table.inner)?.[0];
  if (!prefix) return;
  const position = { treatAsChar: 0, affectLSpacing: 0, allowOverlap: 0, flowWithText: 0,
    vertRelTo: 'PAPER', horzRelTo: 'PAPER', vertAlign: 'TOP', horzAlign: 'LEFT',
    vertOffset: Math.round(proof.y * 75) - proof.top, horzOffset: Math.round(proof.x * 75) };
  const margins = { top: proof.top, bottom: proof.bottom, left: proof.left, right: proof.right };
  const changed = prefix.replace(/<hp:pos\b[^>]*\/>/, tag => patch(tag, position))
    .replace(/<hp:outMargin\b[^>]*\/>/, tag => patch(tag, margins));
  if (changed === prefix) return;
  const at = range.start + range.open.length + table.start + table.open.length;
  parts[name] = strToU8(xml.slice(0, at) + changed + xml.slice(at + prefix.length));
  return new Core(zipSync(parts, {level:1}));
}

export function restoreNativeTableFragment(Core: typeof HwpDocument, doc: HwpDocument,
  proofs: readonly AuthoredNativeTableFragment[] = [], printPages = 0) {
  if (!proofs.length || doc.pageCount() <= printPages) return doc;
  for (const proof of proofs) {
    let candidate: HwpDocument | undefined;
    try {
      candidate = restoreFragment(Core, doc, proof);
      if (!candidate || candidate.pageCount() !== doc.pageCount() - 1) continue;
      const page = proof.page - 1;
      const before = nativePaintedControls(doc, page), [target] = nativePaintedControls(doc, page + 1);
      const after = nativePaintedControls(candidate, page);
      const moved = after.find(control => sameOwner(control, target));
      const paper = JSON.parse(candidate.getPageInfo(page));
      if (!moved || !close(moved.x, proof.x) || !close(moved.y, proof.y) || moved.y + moved.h > paper.footerArea.y + .08 ||
          JSON.stringify(after.filter(item => item !== moved)) !== JSON.stringify(before) ||
          ownInk(candidate, page, moved).some(run => run.x < moved.x - .15 ||
            visibleRight(run) > Math.min(paper.bodyRight, moved.x + moved.w) + .15 ||
            run.y < moved.y - .15 || run.y + run.h > paper.footerArea.y + .15) ||
          nativePaintedPayload(doc) !== nativePaintedPayload(candidate)) continue;
      const accepted = candidate;
      candidate = undefined;
      return accepted;
    } catch { /* Stay in ordinary flow if the authored first fragment cannot fit. */ }
    finally { candidate?.free(); }
  }
  return doc;
}
