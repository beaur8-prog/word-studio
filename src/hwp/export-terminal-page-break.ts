import type { HwpDocument } from '@rhwp/core';
import { strFromU8, strToU8, zipSync } from 'fflate';
import { parseHwpBodyLine } from '../hwp-body-line';
import { retainedTextBands } from '../hwp-line-ink';
import { hwpSourceWrapAttribute, hwpWrapFingerprint, normalizeHwpWrapText, parseHwpSourceWrap, type HwpAuthoredNativeLine } from '../hwp-text-advances';
import { paragraphLineOffsets } from '../paragraph-fragments';
import { hiddenEmptyLineClass } from '../page-empty-lines';
import { authoredEmptySize } from './export-empty-paragraph-format';
import { boundedNativeParts, nativeBodyRuns, nativeGlyphBaselines, nativePaintedControls, nativePaintedPayload, type NativeBodyRun } from './export-binary-pages';
import { directHwpPageBreak, elementRanges } from './markup-ranges';

import { nativeLineTypography } from './script-fonts';

type EmptyPosition = { font: number; height: number; origin: number };
export type TerminalPageBreak = {
  page: number; literal: string; starts: number[]; font: number; pitch: number; percent: number;
  origin: number; bodyWidth: number; before: EmptyPosition; after: EmptyPosition;
};

const units = (pixels: number) => Math.round(pixels * 75);
const graphemes = new Intl.Segmenter(undefined, { granularity: 'grapheme' });
const visible = (literal: string) => literal.replace(/[\s\x00-\x1f\ufffc\ufeff]/gu,'');
const same = (a: unknown,b: unknown) => JSON.stringify(a) === JSON.stringify(b);
const display = ({secIdx,paraIdx,parentParaIdx,cellPath,charStart,text,x,y,w,h,fontSize}: NativeBodyRun) =>
  [secIdx,paraIdx,parentParaIdx,cellPath,charStart,text,x,y,w,h,fontSize];
const ink = (doc: HwpDocument,page: number) => nativeBodyRuns(doc,page).filter(run => visible(run.text));

/** Source wrap offsets count normalized visible text. Keep literal whitespace
 * at the end of its preceding line when translating to native text positions. */
function authoredTerminalStarts(paragraph: HTMLParagraphElement, literal: string, page: number,
  width: number, font: number, pitch: number, lineCount: number) {
  const wrap = parseHwpSourceWrap(paragraph.getAttribute(hwpSourceWrapAttribute));
  const reference = wrap?.authoredBodyLines;
  if (!reference || wrap.fingerprint !== hwpWrapFingerprint(literal) ||
      wrap.ends.length !== lineCount || wrap.ends.at(-1) !== normalizeHwpWrapText(literal).length ||
      reference.origin.page !== page+1 || Math.abs(reference.width-width) > .15 ||
      Math.abs(reference.size-font) > .08 || Math.abs(reference.pitch-pitch) > .08) return;
  const style = getComputedStyle(paragraph);
  if (!style.textIndent.endsWith('px') || !style.marginLeft.endsWith('px') ||
      Math.abs(Number.parseFloat(style.textIndent)+reference.hanging) > .15 ||
      Math.abs(Number.parseFloat(style.marginLeft)-reference.hanging) > .15) return;
  const starts = [0];
  for (const {index,segment} of graphemes.segment(literal)) {
    if (starts.length === lineCount) break;
    if (segment.trim() && normalizeHwpWrapText(literal.slice(0,index)).length === wrap.ends[starts.length-1]) starts.push(index);
  }
  return starts.length === lineCount ? starts : undefined;
}

/** A printed last body paragraph and its actual empty neighbors fit before
 * the next explicit sheet boundary. Keep each measured origin and line end. */
export function measuredTerminalPageBreak(body: Element,page: number): TerminalPageBreak | undefined {
  const end = [...body.children].slice(-3);
  if (end.length !== 3 || !end.every((p): p is HTMLParagraphElement => p instanceof HTMLParagraphElement)) return;
  const [a, terminal, b] = end;
  const first = authoredEmptySize(a), last = authoredEmptySize(b), literal = terminal.textContent ?? '';
  if (!first || !last || !visible(literal) || literal.length > 512 ||
      /[\x00-\x1f\ufffc\uD800-\uDFFF]/u.test(literal) ||
      terminal.matches('[data-print-paragraph-fragment]') ||
      terminal.querySelector('img,svg,br,table,[data-docx-tab],[data-note-reference],[data-document-field]')) return;
  const format = parseHwpBodyLine(terminal.dataset.hwpBodyLine);
  if (!format || format.font <= 0 || format.percent <= 100 || format.base <= format.font ||
      Math.abs(format.font * format.percent / 100 - format.base) > .08) return;
  const frame = body.getBoundingClientRect();
  const printedStarts = [0,...paragraphLineOffsets(terminal)];
  const positioned = (element: HTMLElement,font: number): EmptyPosition => {
    const r = element.getBoundingClientRect();
    return {font,height:units(r.height),origin:units(r.top-frame.top)};
  };
  const before = positioned(a,first), after = positioned(b,last), origin = units(terminal.getBoundingClientRect().top-frame.top);
  if (printedStarts.length < 2 || printedStarts.length > 16 || printedStarts.at(-1)! >= literal.length ||
      !retainedTextBands(terminal,document.createRange(),frame) || a.getBoundingClientRect().top < frame.top ||
      b.getBoundingClientRect().bottom > frame.bottom || before.height < first || after.height < last ||
      origin <= before.origin || after.origin <= origin) return;
  const starts = authoredTerminalStarts(terminal,literal,page,frame.width,format.font,format.base,printedStarts.length) ?? printedStarts;
  return {page,literal,starts,origin,bodyWidth:frame.width,before,after,
    font:units(format.font),pitch:units(format.base),percent:format.percent};
}

function lineCache(starts: readonly number[], origin: number, font: number, pitch: number, ascent: number, width: number) {
  const lines = starts.map((offset,i) => `<hp:lineseg textpos="${offset}" vertpos="${origin+i*pitch}" vertsize="${font}" textheight="${font}" baseline="${ascent}" spacing="${pitch-font}" horzpos="0" horzsize="${width}" flags="${i?0x160000:0x60000}"/>`);
  return '<hp:linesegarray>'+lines.join('')+'</hp:linesegarray>';
}

function recreateTerminalPackage(prepared: HwpDocument, proof: TerminalPageBreak, section: number, paragraph: number, ascent: number, width: number) {
  for (const [index,fontSize] of [[paragraph-1,proof.before.font],[paragraph+1,proof.after.font]]) {
    const set = JSON.parse(prepared.applyCharFormat(section,index,0,0,JSON.stringify({fontSize})));
    if (!set.ok || prepared.getParagraphLength(section,index)) return;
  }
  // Preserve the existing forced break as a paragraph flag. The break in
  // its paragraph shape would strand the preceding empty line.
  const changed = JSON.parse(prepared.applyParaFormat(section,paragraph+2,JSON.stringify({pageBreakBefore:false})));
  if (!changed.ok) return;
  const parts = boundedNativeParts(prepared.exportHwpx()), file = `Contents/section${section}.xml`;
  const body = strFromU8(parts[file]), owners = elementRanges(body,'hp:p'), next = owners[paragraph+2];
  if (!next || owners.length !== prepared.getParagraphCount(section)) return;
  let xml = directHwpPageBreak(body,next);
  if (!xml) return;
  for (const [index,spec] of [[paragraph+1,proof.after],[paragraph,undefined],[paragraph-1,proof.before]] as const) {
    const owner = owners[index];
    if (!owner || /<hp:(?:linesegarray|ctrl|tbl|pic)\b/.test(owner.inner)) return;
    const offset = owner.end-'</hp:p>'.length;
    const cache = spec
      ? lineCache([0],spec.origin,spec.font,spec.height,Math.round(ascent*spec.font/proof.font),width)
      : lineCache(proof.starts,proof.origin,proof.font,proof.pitch,ascent,width);
    xml = xml.slice(0,offset)+cache+xml.slice(offset);
  }
  parts[file] = strToU8(xml);
  return zipSync(parts,{level:1});
}

function readbackTerminal(Core: typeof HwpDocument, source: HwpDocument, bytes: Uint8Array, proof: TerminalPageBreak) {
  const { page } = proof;
  const priorInk = ink(source,page), duplicate = ink(source,page+1), first = duplicate[0];
  if (!first || nativePaintedControls(source,page+1).length) return;
  const { secIdx: section, paraIdx: paragraph } = first;
  const belongs = (run: NativeBodyRun) => run.secIdx === section && run.paraIdx === paragraph && !run.cellPath?.length;
  const terminal = priorInk.filter(belongs), last = terminal.at(-1);
  if (paragraph < 1 || !last || paragraph+2 >= source.getParagraphCount(section) ||
      !duplicate.every(belongs) || !same(duplicate.map(r=>display({...r,y:0})),
        terminal.filter(r=>r.y===last.y).map(r=>display({...r,y:0}))) ||
      !same(visible(terminal.map(r=>r.text).join('')),visible(proof.literal)) ||
      source.getTextRange(section,paragraph,0,0xffffffff) !== proof.literal ||
      source.getParagraphLength(section,paragraph-1) || source.getParagraphLength(section,paragraph+1) ||
      JSON.parse(source.getPageOfPosition(section,paragraph+1)).page !== page+1 ||
      JSON.parse(source.getPageOfPosition(section,paragraph+2)).page !== page+2 ||
      !JSON.parse(source.getParaPropertiesAt(section,paragraph+2)).pageBreakBefore) return;
  const paper = JSON.parse(source.getPageInfo(page)), width = units(paper.bodyRight-paper.bodyLeft);
  const bodyStyle = JSON.parse(source.getParaPropertiesAt(section,paragraph));
  const measured = JSON.parse(source.getParagraphLineMetrics(section,paragraph,'[]'));
  if (paper.columns.length !== 1 || Math.abs(width/75-proof.bodyWidth) > .15 || measured.stored || measured.controlCount ||
      bodyStyle.lineSpacingType !== 'Percent' || bodyStyle.lineSpacing !== proof.percent ||
      JSON.parse(source.getCharPropertiesAt(section,paragraph,0)).fontSize !== proof.font) return;
  const baseline = nativeGlyphBaselines(source.renderPageSvg(page));
  const ascents = terminal.flatMap(run => {
    if (units(run.fontSize) !== proof.font || units(run.h) !== proof.font) return [];
    const y = baseline(run);
    return y === undefined ? [] : [units(y-run.y)];
  }).sort((a,b)=>a-b);
  const ascent = ascents[Math.floor(ascents.length/2)];
  if (!ascent || ascent <= 0 || ascent > proof.font) return;
  const prepared = new Core(bytes);
  let candidate: HwpDocument | undefined;
  try {
    const rebuilt = recreateTerminalPackage(prepared,proof,section,paragraph,ascent,width);
    if (!rebuilt) return;
    candidate = new Core(rebuilt);
    const pages = candidate.pageCount(), after = ink(candidate,page), moved = after.filter(belongs);
    const right = Math.max(...terminal.map(r=>r.x+r.w));
    const starts = [...new Set(moved.map(r=>r.y))].map(y=>moved.find(r=>r.y===y)!.charStart);
    if (pages !== source.pageCount()-1 || !same(starts,proof.starts) ||
        candidate.getTextRange(section,paragraph,0,0xffffffff) !== proof.literal ||
        visible(moved.map(r=>r.text).join('')) !== visible(proof.literal) ||
        JSON.parse(candidate.getPageOfPosition(section,paragraph+1)).page !== page ||
        JSON.parse(candidate.getPageOfPosition(section,paragraph+2)).page !== page+1 ||
        moved.some(r=>r.y < paper.marginTop+paper.marginHeader || r.y+r.h > paper.height-paper.marginBottom-paper.marginFooter+.15 ||
          r.x < 0 || r.x+r.w > Math.min(paper.width,right+.15)) ||
        !same(priorInk.filter(r=>!belongs(r)).map(display),after.filter(r=>!belongs(r)).map(display)) ||
        !same(nativePaintedControls(source,page),nativePaintedControls(candidate,page))) return;
    for (let i=0;i<pages;i++) {
      if (i === page) continue;
      const old = i < page ? i : i+1;
      if (!same(ink(source,old).map(display),ink(candidate,i).map(display)) ||
          !same(nativePaintedControls(source,old),nativePaintedControls(candidate,i))) return;
    }
    const result = candidate;
    candidate = undefined;
    return result;
  } finally {
    prepared.free();
    candidate?.free();
  }
}

export function restoreTerminalPageBreak(Core: typeof HwpDocument, doc: HwpDocument,
  proofs: readonly TerminalPageBreak[] = [], printedPages = 0): HwpDocument {
  if (!proofs.length || doc.pageCount() !== printedPages+1) return doc;
  let baseline: HwpDocument | undefined;
  try {
    const bytes = doc.exportHwpx();
    baseline = new Core(bytes);
    if (baseline.pageCount() !== doc.pageCount()) return doc;
    for (const proof of proofs) {
      if (Number.isInteger(proof.page) && proof.page >= 0 && proof.page+1 < printedPages) {
        const restored = readbackTerminal(Core,baseline,bytes,proof);
        if (restored) return restored;
      }
    }
    return doc;
  } catch {
    return doc;
  } finally {
    baseline?.free();
  }
}

export type AuthoredTerminalWrap = {
  page: number; literal: string; starts: number[]; x: number; y: number;
  width: number; hanging: number; font: number; pitch: number; percent: number;
};

/** When the last blank fits only as a page-edge carrier, the complete
 * preceding body paragraph still owns its original line starts. */
export function measuredAuthoredTerminalWrap(body: Element, page: number): AuthoredTerminalWrap | undefined {
  const blank = body.lastElementChild, terminal = blank?.previousElementSibling;
  if (!(blank instanceof HTMLParagraphElement) || blank.childNodes.length ||
      !blank.classList.contains(hiddenEmptyLineClass) ||
      !(terminal instanceof HTMLParagraphElement) || terminal.hasAttribute('data-print-paragraph-fragment') ||
      !(terminal.previousElementSibling instanceof HTMLParagraphElement) ||
      terminal.previousElementSibling.childNodes.length ||
      terminal.querySelector('img,svg,br,table,[data-docx-tab],[data-note-reference],[data-document-field]')) return;
  const literal = terminal.textContent ?? '';
  if (!visible(literal) || literal.length > 512 || /[\x00-\x1f\ufffc\uD800-\uDFFF]/u.test(literal)) return;
  const format = parseHwpBodyLine(terminal.dataset.hwpBodyLine);
  const reference = parseHwpSourceWrap(terminal.dataset.hwpSourceWrap)?.authoredBodyLines;
  if (!format || !reference || format.font <= 0 || format.percent <= 100 ||
      Math.abs(format.font * format.percent / 100 - format.base) > .08 ||
      blank.getBoundingClientRect().height !== 0) return;
  const frame = body.getBoundingClientRect(), count = paragraphLineOffsets(terminal).length + 1;
  const starts = authoredTerminalStarts(terminal, literal, page, frame.width, format.font, format.base, count);
  if (!starts || starts.length < 2 || starts.length > 16 || !retainedTextBands(terminal, document.createRange(), frame)) return;
  return { page, literal, starts, x: reference.origin.x, y: reference.origin.y, width: frame.width, hanging: reference.hanging,
    font: units(format.font), pitch: units(format.font * format.percent / 100), percent: format.percent };
}

function applyAuthoredTerminalWrap(Core: typeof HwpDocument, source: HwpDocument, proof: AuthoredTerminalWrap) {
  const { page } = proof;
  const before = ink(source, page), spill = ink(source, page + 1), first = spill[0];
  if (!first || first.cellPath?.length || nativePaintedControls(source, page + 1).length) return;
  const { secIdx: section, paraIdx: paragraph } = first;
  const owns = (run: NativeBodyRun) => run.secIdx === section && run.paraIdx === paragraph && !run.cellPath?.length;
  if (!spill.every(owns) || paragraph < 1 || paragraph + 2 >= source.getParagraphCount(section) ||
      source.getParagraphLength(section, paragraph - 1) || source.getParagraphLength(section, paragraph + 1) ||
      source.getTextRange(section, paragraph, 0, 0xffffffff) !== proof.literal ||
      JSON.parse(source.getPageOfPosition(section, paragraph + 2)).page !== page + 2 ||
      !JSON.parse(source.getParaPropertiesAt(section, paragraph + 2)).pageBreakBefore) return;
  const terminal = before.filter(owns), lead = terminal[0];
  if (!lead || lead.charStart !== 0 || before.at(-1) !== terminal.at(-1) ||
      Math.abs(lead.x - proof.x) > .15 || Math.abs(lead.y - proof.y) > .15 ||
      visible([...terminal, ...spill].map(run => run.text).join('')) !== visible(proof.literal)) return;
  const paper = JSON.parse(source.getPageInfo(page));
  const pageDef = JSON.parse(source.getPageDef(section));
  const width = (pageDef.landscape ? pageDef.height : pageDef.width) -
    pageDef.marginLeft - pageDef.marginRight - pageDef.marginGutter;
  const style = JSON.parse(source.getParaPropertiesAt(section, paragraph));
  const { stored, controlCount } = JSON.parse(source.getParagraphLineMetrics(section, paragraph, '[]'));
  if (paper.columns.length !== 1 || !Number.isSafeInteger(width) || width <= 0 || Math.abs(width / 75 - proof.width) > .15 ||
      stored || controlCount || style.lineSpacingType !== 'Percent' || style.lineSpacing !== proof.percent ||
      Math.abs(style.indent + proof.hanging) > .15 ||
      JSON.parse(source.getCharPropertiesAt(section, paragraph, 0)).fontSize !== proof.font) return;

  const baseline = nativeGlyphBaselines(source.renderPageSvg(page))(lead);
  if (baseline === undefined) return;
  const ascent = units(baseline - lead.y), origin = units(baseline - paper.marginTop - paper.marginHeader) - ascent;
  if (ascent <= 0 || ascent > proof.font || origin < 0) return;
  const bodyBottom = paper.height - paper.marginBottom - paper.marginFooter;
  if (lead.y + ((proof.starts.length - 1) * proof.pitch + proof.font) / 75 > bodyBottom + .15) return;

  const file = `Contents/section${section}.xml`, parts = boundedNativeParts(source.exportHwpx());
  const xml = strFromU8(parts[file]), owners = elementRanges(xml, 'hp:p'), owner = owners[paragraph];
  if (!owner || owners.length !== source.getParagraphCount(section) || /<hp:(?:linesegarray|ctrl|tbl|pic)\b/.test(owner.inner)) return;
  const offset = owner.end - '</hp:p>'.length;
  const cache = lineCache(proof.starts, origin, proof.font, proof.pitch, ascent, width);
  parts[file] = strToU8(xml.slice(0, offset) + cache + xml.slice(offset));
  const candidate = new Core(zipSync(parts, { level: 1 }));
  let accepted = false;
  try {
    const merged = ink(candidate, page), restored = merged.filter(owns);
    const starts = [...new Set(restored.map(run => run.y))].map(y => restored.find(run => run.y === y)!.charStart);
    if (candidate.pageCount() + 1 !== source.pageCount() || !same(starts, proof.starts) ||
        candidate.getTextRange(section, paragraph, 0, 0xffffffff) !== proof.literal ||
        visible(restored.map(run => run.text).join('')) !== visible(proof.literal) ||
        JSON.parse(candidate.getPageOfPosition(section, paragraph + 1)).page !== page ||
        JSON.parse(candidate.getPageOfPosition(section, paragraph + 2)).page !== page + 1 ||
        restored.some(run => run.x < 0 || run.x + run.w > paper.width + .15 ||
          run.y < lead.y - .15 || run.y + run.h > bodyBottom + .15) ||
        !same(before.filter(run => !owns(run)).map(display), merged.filter(run => !owns(run)).map(display)) ||
        !same(nativePaintedControls(source, page), nativePaintedControls(candidate, page))) return;
    for (let i = 0; i < candidate.pageCount(); i++) {
      if (i === page) continue;
      const previous = i < page ? i : i + 1;
      if (!same(ink(source, previous).map(display), ink(candidate, i).map(display)) ||
          !same(nativePaintedControls(source, previous), nativePaintedControls(candidate, i))) return;
    }
    accepted = true;
    return candidate;
  } finally {
    if (!accepted) candidate.free();
  }
}

export function restoreAuthoredTerminalWrap(Core: typeof HwpDocument, doc: HwpDocument,
  proofs: readonly AuthoredTerminalWrap[] = [], printedPages = 0): HwpDocument {
  if (!proofs.length || doc.pageCount() <= printedPages) return doc;
  for (const proof of proofs) {
    if (proof.page < 0 || proof.page + 2 >= doc.pageCount()) continue;
    try {
      const restored = applyAuthoredTerminalWrap(Core, doc, proof);
      if (restored) return restored;
    } catch { /* Keep the ordinary native composition without a complete proof. */ }
  }
  return doc;
}

export type AuthoredSingleLine = HwpAuthoredNativeLine & { literal: string };

export function measuredAuthoredSingleLines(body: Element, page: number): AuthoredSingleLine[] {
  const result: AuthoredSingleLine[] = [];
  for (const p of body.querySelectorAll<HTMLParagraphElement>(':scope > p[data-hwp-source-wrap*="authoredNativeLine"]')) {
    const wrap = parseHwpSourceWrap(p.dataset.hwpSourceWrap), proof = wrap?.authoredNativeLine;
    const literal = p.textContent ?? '';
    if (!proof || proof.origin.page !== page + 1 || literal.length > 512 ||
        wrap.fingerprint !== hwpWrapFingerprint(literal) || p.hasAttribute('data-print-paragraph-fragment') ||
        p.querySelector('img,svg,br,[data-note-reference],[data-document-field],[data-docx-tab],[data-symbol-font]')) continue;
    const font = getComputedStyle(p);
    if (font.textAlign === 'justify' && Math.abs(parseFloat(font.fontSize) - proof.size) <= .08 &&
        Math.abs(parseFloat(font.textIndent) + proof.hanging) <= .15 &&
        Math.abs(body.getBoundingClientRect().width - proof.width) <= .15) result.push({ ...proof, literal });
  }
  return result;
}

function oneLineCandidate(Core: typeof HwpDocument, doc: HwpDocument, proof: AuthoredSingleLine) {
  const page = proof.origin.page - 1;
  if (page < 0 || page + 2 >= doc.pageCount() || nativePaintedControls(doc,page + 1).length) return;
  const prior = nativeBodyRuns(doc,page), leaf = ink(doc,page + 1), spill = leaf[0], terminal = ink(doc,page).at(-1);
  if (!spill || !terminal || terminal.cellPath?.length || terminal.secIdx !== spill.secIdx || terminal.paraIdx !== spill.paraIdx ||
      leaf.some(run => run.cellPath?.length || run.secIdx !== spill.secIdx || run.paraIdx !== spill.paraIdx)) return;

  const owners = prior.filter(run => !run.cellPath?.length && run.charStart === 0 &&
    run.secIdx === spill.secIdx && run.paraIdx < spill.paraIdx && Math.abs(run.x - proof.origin.x) <= .15 &&
    Math.abs(run.y - proof.origin.y) <= proof.pitch && doc.getTextRange(run.secIdx,run.paraIdx,0,0xffffffff) === proof.literal);
  if (owners.length !== 1) return;
  const {secIdx:section,paraIdx:paragraph} = owners[0], firstY = owners[0].y;
  const owns = (run: NativeBodyRun) => run.secIdx === section && run.paraIdx === paragraph && !run.cellPath?.length;
  const lines = prior.filter(owns);
  if (new Set(lines.map(run => run.y)).size !== 2 || lines.map(run => run.text).join('') !== proof.literal) return;
  const props = JSON.parse(doc.getParaPropertiesAt(section,paragraph)), paper = JSON.parse(doc.getPageInfo(page));
  const cached = JSON.parse(doc.getParagraphLineMetrics(section,paragraph,'[]'));
  const marginTop = paper.marginTop + paper.marginHeader;
  const pageDef = JSON.parse(doc.getPageDef(section));
  const width = (pageDef.landscape ? pageDef.height : pageDef.width) - pageDef.marginLeft - pageDef.marginRight - pageDef.marginGutter;
  const { size,pitch,top,ascent } = proof;
  if (paper.columns.length !== 1 || cached.stored || cached.controlCount || !Number.isFinite(paper.footerArea?.y) ||
      !(width > 0 && Math.abs(width / 75 - proof.width) <= .15) || Math.abs(firstY - marginTop - top) > pitch ||
      props.alignment !== 'justify' || props.lineSpacingType !== 'Percent' || props.marginLeft || props.marginRight ||
      Math.abs(props.indent + proof.hanging) > .15 || Math.abs(props.lineSpacing * size / 100 - pitch) > .08 ||
      JSON.parse(doc.getCharPropertiesAt(section, paragraph, 0)).fontSize !== units(size) ||
      nativeLineTypography(doc, section, paragraph, proof.literal.length) !== proof.typography) return;

  const file = `Contents/section${section}.xml`, parts = boundedNativeParts(doc.exportHwpx());
  const xml = strFromU8(parts[file]), paragraphs = elementRanges(xml,'hp:p'), owner = paragraphs[paragraph];
  if (!owner || paragraphs.length !== doc.getParagraphCount(section) || /<hp:(?:linesegarray|ctrl|tbl|pic)\b/.test(owner.inner)) return;
  const at = owner.end - '</hp:p>'.length;
  parts[file] = strToU8(xml.slice(0,at) + lineCache([0],units(firstY - marginTop),units(size),units(pitch),units(ascent),width) + xml.slice(at));
  const candidate = new Core(zipSync(parts,{level:1}));
  let accepted = false;
  try {
    if (candidate.pageCount() + 1 !== doc.pageCount() || nativePaintedPayload(candidate) !== nativePaintedPayload(doc)) return;
    const restored = nativeBodyRuns(candidate,page).filter(owns);
    if (!restored.length || restored.map(run => run.text).join('') !== proof.literal ||
        restored.some(run => Math.abs(run.y - firstY) > .15 || run.x < paper.bodyLeft - .15 ||
          run.x + run.w > paper.bodyRight + .15 || run.y + run.h > paper.footerArea.y + .15)) return;
    accepted = true;
    return candidate;
  } finally { if (!accepted) candidate.free(); }
}

export function restoreAuthoredNativeSingleLine(Core: typeof HwpDocument, doc: HwpDocument,
  proofs: readonly AuthoredSingleLine[] = [], printPages = 0): HwpDocument {
  if (!proofs.length || doc.pageCount() <= printPages) return doc;
  for (const proof of proofs) {
    try {
      const restored = oneLineCandidate(Core,doc,proof);
      if (restored) return restored;
    } catch { /* The output can always use the native wrapped line without a complete fit. */ }
  }
  return doc;
}
