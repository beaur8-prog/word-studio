import type { HwpDocument } from '@rhwp/core';
import { strFromU8, strToU8, zipSync } from 'fflate';
import documentLimits from '../../electron/document-limits.json';
import { retainedTextBands } from '../hwp-line-ink';
import { boundedNativeParts, nativeBodyRuns, nativePaintedControls, type NativeBodyRun } from './export-binary-pages';
import { directHwpPageBreak, elementRanges } from './markup-ranges';

export type PastedEmptyParagraphs = { text: string; count: number; cells?: { text: string; count: number } };

function restoreEmptyBodyParagraphs(doc: HwpDocument, text: string, count: number) {
  const locations: number[] = [];
  const total = doc.getParagraphCount(0);
  for (let paragraph = 0; paragraph < total; paragraph++) {
    if (doc.getParagraphLength(0, paragraph) === text.length &&
        doc.getTextRange(0, paragraph, 0, text.length) === text) locations.push(paragraph);
  }
  if (locations.length !== count) return false;
  doc.beginBatch();
  try {
    return locations.every(paragraph => {
      const { charShapeId } = JSON.parse(doc.getCharPropertiesAt(0, paragraph, 0));
      if (!Number.isSafeInteger(charShapeId) || charShapeId < 0 ||
          !JSON.parse(doc.deleteText(0, paragraph, 0, text.length)).ok ||
          doc.getParagraphLength(0, paragraph) !== 0) return false;
      // A temporary tab-setting run can leave its preceding default style at
      // offset zero. Reuse the actual authored shape from the removed text.
      return JSON.parse(doc.getCharPropertiesAt(0, paragraph, 0)).charShapeId === charShapeId ||
        JSON.parse(doc.setCharShapeId(0, paragraph, 0, 0, charShapeId)).ok;
    });
  } finally { doc.endBatch(); }
}

export function restorePastedEmptyParagraphs(Core: typeof HwpDocument, doc: HwpDocument, { text, count, cells }: PastedEmptyParagraphs) {
  if (count && !restoreEmptyBodyParagraphs(doc, text, count)) return;
  if (!cells) return doc;
  const parts = boundedNativeParts(doc.exportHwpx());
  const literal = '<hp:t>' + cells.text + '</hp:t>';
  let restored = 0;
  for (const [name, bytes] of Object.entries(parts)) {
    if (!/^Contents\/section\d+\.xml$/.test(name)) continue;
    const xml = strFromU8(bytes), pieces = xml.split(literal);
    if (pieces.some(piece => piece.includes(cells.text))) return;
    if (pieces.length > 1) {
      restored += pieces.length - 1;
      // Keep the pasted character properties on the now empty cell line.
      parts[name] = strToU8(pieces.join('<hp:t/>'));
    }
  }
  if (restored !== cells.count) return;
  return new Core(zipSync(parts, { level: 1 }));
}

export function withoutPastedEmptyCarriers(html: string, { text, cells }: PastedEmptyParagraphs) {
  const body = html.replaceAll(text, '');
  return cells ? body.replace(/(<span\b[^>]*\bdata-hwp-empty-cell-carrier="true"[^>]*>)([^<]*)(<\/span>)/g,
    (all, open, value, close) => value === cells.text ? open + close : all) : body;
}

export type TerminalEmptyFonts = {
  page: number; heading: string; terminal: string;
  beforeHeading: number; beforeTerminal: number;
};

const hasInk = (text: string) => /[^\s\x00-\x1f\ufffc\ufeff]/u.test(text);
const same = (left: unknown, right: unknown) => JSON.stringify(left) === JSON.stringify(right);

export function authoredEmptySize(p: HTMLParagraphElement) {
  if (p.textContent !== '' || p.children.length || !p.hasAttribute('data-exact-line-height') || !p.style.fontSize) return;
  const css = getComputedStyle(p), font = Math.round(parseFloat(css.fontSize) * 75);
  const height = p.getBoundingClientRect().height;
  if (Number.isInteger(font) && font > 0 && font <= 163800 && height > 0 &&
      Math.abs(height - parseFloat(css.lineHeight)) < .1) return font;
}

/** Measure the two authored empty paragraphs beside a terminal text block.
 * They occupy their own lines even though neither contains a text run. */
export function measuredTerminalEmptyFonts(body: Element, page: number): TerminalEmptyFonts | undefined {
  const end = [...body.children].slice(-4);
  if (end.length !== 4 || !end.every((p): p is HTMLParagraphElement => p instanceof HTMLParagraphElement)) return;
  const [empty, heading, between, terminal] = end;
  const beforeHeading = authoredEmptySize(empty), beforeTerminal = authoredEmptySize(between);
  const title = heading.textContent ?? '', content = terminal.textContent ?? '';
  if (!beforeHeading || !beforeTerminal || !hasInk(title) || !hasInk(content) ||
      title.length > 256 || content.length > 1024 || /[\x00-\x1f\ufffc]/u.test(title+content) ||
      end.some(p => p.querySelector('img,table,br,svg,[data-docx-tab],[data-note-reference],[data-document-textbox]')) ||
      heading.matches('[data-print-paragraph-fragment]') || terminal.matches('[data-print-paragraph-fragment]')) return;
  const rect = terminal.getBoundingClientRect(), frame = body.getBoundingClientRect();
  if (rect.height <= 0 || rect.top < frame.top ||
      !retainedTextBands(terminal,document.createRange(),frame)) return;
  return {page,heading:title,terminal:content,beforeHeading,beforeTerminal};
}

const pageInk = (doc: HwpDocument, page: number) => nativeBodyRuns(doc,page).filter(run => hasInk(run.text));

function tryTerminalFontRestoration(Core: typeof HwpDocument, bytes: Uint8Array, source: HwpDocument, proof: TerminalEmptyFonts) {
  const {page,heading,terminal,beforeHeading,beforeTerminal} = proof;
  const own = pageInk(source,page), spill = pageInk(source,page+1), first = spill[0];
  if (!first || nativePaintedControls(source,page+1).length) return;
  const {secIdx:section,paraIdx:paragraph} = first;
  const isBody = (run: NativeBodyRun, index: number) => run.secIdx === section && run.paraIdx === index && !run.cellPath?.length;
  const lead = paragraph - 2, preceding = own.filter(run => isBody(run,paragraph));
  if (paragraph < 3 || !preceding.length || !own.some(run => isBody(run,lead)) ||
      spill.some(run => !isBody(run,paragraph)) || first.charStart <= preceding.at(-1)!.charStart ||
      source.getTextRange(section,lead,0,0xffffffff) !== heading ||
      source.getTextRange(section,paragraph,0,0xffffffff) !== terminal) return;
  const sizes = [[lead-1,beforeHeading],[paragraph-1,beforeTerminal]] as const;
  if (sizes.some(([index,font]) => source.getParagraphLength(section,index) ||
      JSON.parse(source.getParagraphLineMetrics(section,index,'[]')).controlCount ||
      !Number.isInteger(font) || font <= 0 || font > 163800)) return;
  const candidate = new Core(bytes);
  let accepted = false;
  try {
    for (const [index,fontSize] of sizes) {
      if (!JSON.parse(candidate.applyCharFormat(section,index,0,0,JSON.stringify({fontSize}))).ok ||
          candidate.getParagraphLength(section,index) ||
          JSON.parse(candidate.getCharPropertiesAt(section,index,0)).fontSize !== fontSize) return;
    }
    const pages = candidate.pageCount();
    if (pages !== source.pageCount()-1) return;
    const after = pageInk(candidate,page), shape = JSON.parse(candidate.getPageInfo(page));
    const changedY = (run: NativeBodyRun) => isBody(run,lead) || isBody(run,paragraph);
    const stable = (runs: NativeBodyRun[]) => runs.map(run => changedY(run) ? {...run,y:0} : run);
    if (!same(stable([...own,...spill]),stable(after)) ||
        !same(nativePaintedControls(source,page),nativePaintedControls(candidate,page)) ||
        after.filter(changedY).some(run => run.y < shape.marginTop + shape.marginHeader - .15 ||
          run.y + run.h > shape.height - shape.marginBottom - shape.marginFooter + .15 ||
          run.x < 0 || run.x + run.w > shape.width + .15)) return;
    for (let i=0;i<pages;i++) {
      if (i === page) continue;
      const previous = i < page ? i : i+1;
      if (!same(pageInk(source,previous),pageInk(candidate,i)) ||
          !same(nativePaintedControls(source,previous),nativePaintedControls(candidate,i))) return;
    }
    accepted = true;
    return candidate;
  } finally {
    if (!accepted) candidate.free();
  }
}

/** Keep empty paragraph formatting when it accounts for the measured
 * one-sheet spill, with the complete text and all other sheets retained. */
export function restoreTerminalEmptyFonts(Core: typeof HwpDocument, doc: HwpDocument,
  proofs: readonly TerminalEmptyFonts[] = [], printedPages = 0): HwpDocument {
  if (!proofs.length || doc.pageCount() !== printedPages+1) return doc;
  let baseline: HwpDocument | undefined;
  try {
    const bytes = doc.exportHwpx();
    baseline = new Core(bytes);
    if (baseline.pageCount() !== doc.pageCount()) return doc;
    for (const proof of proofs) {
      if (Number.isInteger(proof.page) && proof.page >= 0 && proof.page+1 < printedPages) {
        const restored = tryTerminalFontRestoration(Core,bytes,baseline,proof);
        if (restored) return restored;
      }
    }
    return doc;
  } catch {
    return doc;
  } finally { baseline?.free(); }
}

/** Once the empty lines fit in the logical document, a binary-only spill
 * before an authored empty break owner can be caused by the flag encoding. */
export function repairBinaryTerminalEmptyPage(Core: typeof HwpDocument, source: HwpDocument,
  binary: HwpDocument, proofs: readonly TerminalEmptyFonts[] | undefined, expected: number | undefined, loss: string) {
  if (!proofs?.length || !expected || source.pageCount() !== expected || binary.pageCount() !== expected+1) return;
  try {
    for (const proof of proofs) {
      if (!Number.isInteger(proof.page) || proof.page < 0 || proof.page+1 >= expected) continue;
      const repaired = tryBinaryTerminalEmptyPage(Core,source,binary,proof,expected,loss);
      if (repaired) return repaired;
    }
  } catch {
    // Keep the already verified writer output when the alternate encoding fails.
  }
}

function tryBinaryTerminalEmptyPage(Core: typeof HwpDocument, source: HwpDocument, binary: HwpDocument,
  proof: TerminalEmptyFonts, expected: number, loss: string) {
  const {page,heading,terminal,beforeHeading,beforeTerminal} = proof;
  const printed = pageInk(source,page), spill = pageInk(binary,page+1), first = spill[0];
  if (!first || nativePaintedControls(binary,page+1).length) return;
  const {secIdx:section,paraIdx:paragraph} = first, next = paragraph+1;
  const isTerminal = (run: Pick<NativeBodyRun,'secIdx'|'paraIdx'|'cellPath'>) => run.secIdx === section && run.paraIdx === paragraph && !run.cellPath?.length;
  const prior = pageInk(binary,page), last = prior.filter(isTerminal).at(-1);
  if (!last || first.charStart <= last.charStart || !spill.every(isTerminal) || paragraph < 3 ||
      next >= source.getParagraphCount(section) || source.getParagraphLength(section,next) ||
      source.getTextRange(section,paragraph-2,0,0xffffffff) !== heading ||
      source.getTextRange(section,paragraph,0,0xffffffff) !== terminal ||
      JSON.parse(source.getCharPropertiesAt(section,paragraph-3,0)).fontSize !== beforeHeading ||
      JSON.parse(source.getCharPropertiesAt(section,paragraph-1,0)).fontSize !== beforeTerminal ||
      !JSON.parse(source.getParaPropertiesAt(section,next)).pageBreakBefore ||
      JSON.parse(source.getParagraphLineMetrics(section,next,'[]')).controlCount ||
      JSON.parse(source.getPageOfPosition(section,next)).page !== page+1 ||
      JSON.parse(binary.getPageOfPosition(section,next)).page !== page+2) return;
  const combined = [...prior,...spill];
  if (combined.filter(isTerminal).map(run=>run.text).join('') !== printed.filter(isTerminal).map(run=>run.text).join('')) return;

  const prepared = new Core(source.exportHwpx());
  let candidate: HwpDocument | undefined, check: HwpDocument | undefined;
  try {
    if (!JSON.parse(prepared.applyParaFormat(section,next,JSON.stringify({pageBreakBefore:false}))).ok) return;
    const file = `Contents/section${section}.xml`, parts = boundedNativeParts(prepared.exportHwpx());
    const xml = strFromU8(parts[file]), owners = elementRanges(xml,'hp:p');
    if (owners.length !== prepared.getParagraphCount(section) || !owners[next]) return;
    const directBreak = directHwpPageBreak(xml,owners[next]);
    if (!directBreak) return;
    parts[file] = strToU8(directBreak);
    candidate = new Core(zipSync(parts,{level:1}));
    const output = candidate.exportHwpWithReport();
    let bytes: Uint8Array;
    try {
      if (output.contentLoss() !== loss) return;
      bytes = output.takeBytes();
    } finally { output.free(); }
    if (bytes.byteLength > documentLimits.documentFileBytes) return;
    check = new Core(bytes);
    if (check.pageCount() !== expected || check.getParagraphLength(section,next) ||
        check.getTextRange(section,paragraph,0,0xffffffff) !== terminal ||
        JSON.parse(check.getPageOfPosition(section,next)).page !== page+1) return;

    const comparable = (runs: NativeBodyRun[], moving = false) =>
      runs.map(({charShapeId,y,...run})=>({...run,y:moving && isTerminal(run) ? 0 : y}));
    const retained = pageInk(check,page), paper = JSON.parse(check.getPageInfo(page));
    if (!same(comparable(combined,true),comparable(retained,true)) ||
        !same(nativePaintedControls(binary,page),nativePaintedControls(check,page)) ||
        retained.filter(isTerminal).some(run => run.y < paper.marginTop+paper.marginHeader-.15 ||
          run.y+run.h > paper.height-paper.marginBottom-paper.marginFooter+.15 || run.x < 0 || run.x+run.w > paper.width+.15)) return;
    for (let i = 0; i < expected; i++) {
      if (i === page) continue;
      const previous = i < page ? i : i+1;
      if (!same(comparable(pageInk(binary,previous)),comparable(pageInk(check,i))) ||
          !same(nativePaintedControls(binary,previous),nativePaintedControls(check,i))) return;
    }
    return {bytes,pages:expected};
  } finally {
    check?.free();
    candidate?.free();
    prepared.free();
  }
}
