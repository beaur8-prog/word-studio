import type { HwpDocument } from '@rhwp/core';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import documentLimits from '../../electron/document-limits.json';
import { elementRanges } from './markup-ranges';

export type NativeBodyRun = {
  secIdx: number; paraIdx: number; charStart: number; parentParaIdx?: number;
  text: string; x: number; y: number; w: number; h: number; fontSize: number; charShapeId: number;
  headerFooter?: object; cellPath?: object[]; charX?: number[];
};
export type NativePaintedControl = { headerFooter?: object; type: string; secIdx: number; paraIdx: number;
  controlIdx: number; x: number; y: number; w: number; h: number; rowStart?: number; rowEnd?: number;
  cells?: object[]; cellPath?: object[] };
const text = (literal: string) => literal.replace(/\s/gu, '');
export const nativeBodyRuns = (doc: HwpDocument, page: number): NativeBodyRun[] =>
  (JSON.parse(doc.getPageTextLayout(page)).runs as NativeBodyRun[]).filter(run => !run.headerFooter && run.paraIdx !== 0xffffffff);
export const nativePaintedControls = (doc: HwpDocument, page: number): NativePaintedControl[] =>
  (JSON.parse(doc.getPageControlLayout(page)).controls as NativePaintedControl[]).filter(control => !control.headerFooter);
/* Compare visible text by native owner and retain each painted control's size
 * across a layout-only readback. */
export function nativePaintedPayload(doc: HwpDocument) {
  const paragraphs = new Map<string, string>();
  const controls: string[] = [];
  for (let page = 0; page < doc.pageCount(); page++) {
    for (const run of nativeBodyRuns(doc, page)) {
      if (!Number.isInteger(run.paraIdx)) continue;
      const key = JSON.stringify([run.secIdx, run.parentParaIdx, run.paraIdx, run.cellPath]);
      const value = run.text.replace(/\s/gu, '');
      if (value) paragraphs.set(key, (paragraphs.get(key) ?? '') + value);
    }
    for (const item of nativePaintedControls(doc, page))
      controls.push(JSON.stringify([item.secIdx,item.paraIdx,item.controlIdx,item.type,item.w,item.h,item.cellPath]));
  }
  return JSON.stringify([[...paragraphs].sort(),controls.sort()]);
}

const same = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b);
const units = (value: number) => Math.round(value * 75);
const cacheTag = '<hp:linesegarray>';

function duplicateTerminal(binary: HwpDocument, pageCount: number) {
  let found: { page: number; section: number; paragraph: number; prior: NativeBodyRun[] } | undefined;
  for (let page = 1; page + 1 < pageCount; page++) {
    const remainder = nativeBodyRuns(binary, page).filter(run => text(run.text));
    const first = remainder[0];
    if (!first || remainder.some(run => run.cellPath?.length || run.secIdx !== first.secIdx ||
        run.paraIdx !== first.paraIdx || run.y !== first.y)) continue;
    const { secIdx: section, paraIdx: paragraph } = first;
    if (paragraph + 1 >= binary.getParagraphCount(section) ||
        !JSON.parse(binary.getParaPropertiesAt(section, paragraph + 1)).pageBreakBefore ||
        nativePaintedControls(binary, page).length) continue;
    const prior = nativeBodyRuns(binary, page - 1).filter(run => run.secIdx === section && run.paraIdx === paragraph && !run.cellPath?.length);
    const final = prior.filter(run => run.y === prior.at(-1)?.y && text(run.text));
    const boundary = (run: NativeBodyRun) => [run.charStart, run.text, run.x, run.w, run.h];
    if (!final.length || !same(final.map(boundary), remainder.map(boundary)) ||
        text(prior.map(run => run.text).join('')) !== text(binary.getTextRange(section, paragraph, 0, 0xffffffff))) continue;
    if (found) return;
    found = { section, paragraph, page, prior };
  }
  return found;
}

export function nativeGlyphBaselines(svg: string) {
  const coordinates = [...svg.matchAll(/<text\s[^>]*>/g)].map(([tag]) => {
    const numeric = (part: string) => Number(tag.match(new RegExp(`\\b${part}="([-+\\d.]+)"`))?.[1]);
    return { x: numeric('x'), y: numeric('y'), size: numeric('font-size') };
  });
  return (run: NativeBodyRun) => {
    const baselines = new Set(coordinates.filter(at => Math.abs(at.size - run.fontSize) <= .065 &&
      at.x >= run.x - .065 && at.x <= run.x + run.w + .065 && at.y > run.y && at.y <= run.y + run.h).map(at => at.y));
    return baselines.size === 1 ? [...baselines][0] : undefined;
  };
}

function owns(run: NativeBodyRun, section: number, paragraph: number) {
  return run.secIdx === section && run.paraIdx === paragraph && !run.cellPath?.length;
}
function ownerRuns(doc: HwpDocument, page: number, section: number, paragraph: number) {
  return nativeBodyRuns(doc, page).filter(run => owns(run, section, paragraph));
}
type MeasuredSheet = { page: number; runs: NativeBodyRun[] };

function measuredCache(source: HwpDocument, section: number, paragraph: number, sheets: readonly MeasuredSheet[]) {
  const details = JSON.parse(source.getParagraphLineMetrics(section, paragraph, '[]'));
  const format = JSON.parse(source.getParaPropertiesAt(section, paragraph));
  const paper = JSON.parse(source.getPageInfo(sheets[0].page));
  if (details.stored || details.controlCount || format.lineSpacingType !== 'Percent' ||
      JSON.parse(source.getColumnDef(section)).columnCount !== 1) return;
  const font = JSON.parse(source.getCharPropertiesAt(section, paragraph, 0)).fontSize;
  if (!Number.isInteger(font) || font <= 0) return;
  const pitch = font + Math.round(font * (format.lineSpacing / 100 - 1));
  const measured: { start: number; top: number; baseline: number }[] = [];
  let last = -1;
  for (const { page, runs } of sheets) {
    const sheetStart = measured.length;
    if (runs.some(run => Math.abs(run.h - font / 75) > .065)) return;
    const lines = new Map<number, NativeBodyRun>();
    for (const run of runs) if (!lines.has(run.y)) lines.set(run.y, run);
    const baselineAt = nativeGlyphBaselines(source.renderPageSvg(page));
    for (const [top, run] of lines) {
      if (run.charStart <= last || (last < 0 && run.charStart !== 0) || run.w <= 0) return;
      const baseline = baselineAt(run);
      if (baseline === undefined) return;
      const sheetLine = measured.length - sheetStart;
      if (sheetLine && Math.abs(units(baseline - measured[sheetStart].baseline) - sheetLine * pitch) > 1) return;
      measured.push({ start: run.charStart, top, baseline });
      last = run.charStart;
    }
  }
  if (measured.length < 2) return;
  const gap = pitch - font;
  const width = units(paper.width - paper.marginLeft - paper.marginRight);
  const first = measured[0].baseline;
  // The native JSON rounds glyph tops. Preserve exact pitch from the
  // unrounded SVG baselines rather than quantizing each line independently.
  const offsets = measured.map(line => units(line.baseline - line.top)).sort((a,b) => a-b);
  const ascent = offsets[Math.floor(offsets.length / 2)];
  const origin = units(first - paper.marginTop - paper.marginHeader) - ascent;
  if (![gap, width, ascent, origin].every(Number.isSafeInteger) || gap < 0 || ascent <= 0 || origin < 0 || width <= 0) return;
  const cache = measured.map(({ start },i) => {
    // HWP continuation lines use the hanging indent, unlike the first line.
    const flags = i ? 0x160000 : 0x60000;
    return `<hp:lineseg textpos="${start}" vertpos="${origin + i * pitch}" vertsize="${font}" textheight="${font}" baseline="${ascent}" spacing="${gap}" horzpos="0" horzsize="${width}" flags="${flags}"/>`;
  });
  return cacheTag + cache.join('') + '</hp:linesegarray>';
}

export function boundedNativeParts(bytes: Uint8Array) {
  let used = 0;
  return unzipSync(bytes, { filter: entry => {
    if ((used += entry.originalSize) > documentLimits.documentInflatedBytes) throw Error('Native line validation too large');
    return true;
  } });
}

export function sameNativeLineContent(a: HwpDocument, b: HwpDocument, file: string, paragraph: number) {
  const before = boundedNativeParts(a.exportHwpx()), after = boundedNativeParts(b.exportHwpx());
  const names = Object.keys(before);
  if (!same(names,Object.keys(after))) return false;
  for (const name of names) {
    const x = before[name], y = after[name];
    if (name !== file) {
      if (x.length !== y.length || x.some((byte,i) => byte !== y[i])) return false;
      continue;
    }
    const xml = strFromU8(y), owner = elementRanges(xml, 'hp:p')[paragraph];
    if (!owner) return false;
    const cleaned = owner.inner.replace(/<hp:linesegarray\b[^>]*>[\s\S]*?<\/hp:linesegarray>/, '');
    if (cleaned === owner.inner || strFromU8(x) !== xml.slice(0,owner.start) + owner.open + cleaned + '</hp:p>' + xml.slice(owner.end))
      return false;
  }
  return true;
}

type Replay = { section: number; paragraph: number; page: number; cache: string; relocated?: NativeBodyRun[] };
const visible = (runs: NativeBodyRun[]) => runs.filter(run => text(run.text));
const ink = ({ text, charStart, x, w, h, fontSize }: NativeBodyRun) => [text, charStart, x, w, h, fontSize];

function displacedTerminal(source: HwpDocument, binary: HwpDocument, pageCount: number): Replay | undefined {
  let result: Replay | undefined;
  for (let page = 1; page + 1 < pageCount; page++) {
    const tail = visible(nativeBodyRuns(binary, page)), first = tail[0];
    if (!first || nativePaintedControls(binary, page).length) continue;
    const { secIdx: section, paraIdx: paragraph } = first;
    if (tail.some(run => !owns(run, section, paragraph) || run.y !== first.y) ||
        paragraph + 1 >= binary.getParagraphCount(section) ||
        !JSON.parse(binary.getParaPropertiesAt(section, paragraph + 1)).pageBreakBefore) continue;
    const before = visible(nativeBodyRuns(binary, page - 1));
    const last = before.at(-1);
    if (!last || before.some(run => !owns(run, section, paragraph)) ||
        first.charStart <= last.charStart ||
        binary.getTextRange(section, paragraph, 0, 0xffffffff) !== source.getTextRange(section, paragraph, 0, 0xffffffff)) continue;
    const relocated = visible(ownerRuns(source, page - 1, section, paragraph));
    if (!same(relocated.map(ink),[...before, ...tail].map(ink)) ||
        !same(relocated.slice(0,before.length),before)) continue;
    const paper = JSON.parse(source.getPageInfo(page - 1));
    if (relocated.slice(before.length).some(run => run.y <= last.y ||
        run.y + run.h > paper.footerArea.y || run.x < paper.bodyLeft - .15 || run.x + run.w > paper.bodyRight + .15)) continue;
    const earlier = page > 1 ? ownerRuns(source, page - 2, section, paragraph) : [];
    const sourceSheets: MeasuredSheet[] = [{ page: page - 1, runs: relocated }];
    if (earlier.length) sourceSheets.unshift({ page: page - 2, runs: earlier });
    const cache = measuredCache(source, section, paragraph, sourceSheets);
    if (!cache) continue;
    if (result) return;
    result = { section, paragraph, page, cache, relocated };
  }
  return result;
}

/** Replay the proven source line layout if binary HWP adds a one-line sheet.
 * The retry must retain the document bytes other than that line cache and
 * reproduce the same ink and controls everywhere outside the repaired page. */
export function repairBinaryExtraPage(Core: typeof HwpDocument, source: HwpDocument, binary: HwpDocument,
  pageCount: number, loss: string): { bytes: Uint8Array; pages: number } | undefined {
  try {
    const expected = source.pageCount();
    if (pageCount !== expected + 1) return;
    const duplicate = duplicateTerminal(binary, pageCount);
    const replays: Replay[] = [];
    if (duplicate) {
      const { section, paragraph, page, prior } = duplicate;
      const runs = ownerRuns(source, page - 1, section, paragraph);
      if (same(runs.map(run => [...ink(run),run.y]),prior.map(run => [...ink(run),run.y]))) {
        const cache = measuredCache(source, section, paragraph, [{ page: page - 1, runs }]);
        if (cache) replays.push({ section, paragraph, page, cache });
      }
    }
    const displaced = displacedTerminal(source, binary, pageCount);
    if (displaced) replays.push(displaced);
    for (const { section, paragraph, page, cache, relocated } of replays) {
      const file = `Contents/section${section}.xml`, files = boundedNativeParts(source.exportHwpx());
      const xml = strFromU8(files[file]), owners = elementRanges(xml, 'hp:p'), owner = owners[paragraph];
      if (!owner || owners.length !== source.getParagraphCount(section) || owner.inner.includes('hp:linesegarray')) continue;
      files[file] = strToU8(xml.slice(0,owner.end - '</hp:p>'.length) + cache + xml.slice(owner.end - '</hp:p>'.length));
      const withCache = new Core(zipSync(files, { level: 1 }));
      let check: HwpDocument | undefined;
      try {
        const output = withCache.exportHwpWithReport();
        let bytes: Uint8Array;
        try {
          if (output.contentLoss() !== loss) continue;
          bytes = output.takeBytes();
        } finally { output.free(); }
        if (bytes.byteLength > documentLimits.documentFileBytes) continue;
        check = new Core(bytes);
        if (check.pageCount() !== expected || !sameNativeLineContent(binary, check, file, paragraph)) continue;
        let agrees = true;
        for (let i = 0; i < expected; i++) {
          const previous = i >= page ? i + 1 : i;
          const actual = nativeBodyRuns(check, i);
          const textMatches = relocated && i === page - 1 ? same(visible(actual), relocated) :
            same(nativeBodyRuns(binary, previous), actual);
          if (!textMatches || !same(nativePaintedControls(binary, previous), nativePaintedControls(check, i))) {
            agrees = false;
            break;
          }
        }
        if (agrees) return { bytes, pages: expected };
      } finally { check?.free(); withCache.free(); }
    }
  } catch { /* Existing HWP remains available if exact reproduction cannot be proved. */ }
}
