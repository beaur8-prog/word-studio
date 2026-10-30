import type { HwpDocument } from '@rhwp/core';
import type { CellPath } from './cell-content';
import { encodeHwpCellLine } from '../hwp-cell-line';

const escape = (text: string) => text.replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]!);

export const cellParagraphKey = (section: number, paragraph: number, path: CellPath) =>
  `${section}:${paragraph}:` + path.map(c => `${c.controlIndex}:${c.cellIndex}:${c.cellParaIndex}`).join('/');

export function physicalLineHeight(properties: Record<string, any>, char?: Record<string, any>): number {
  const fontPx = Number(char?.fontSize) / 75;
  const line = Number(properties.lineSpacing);
  return properties.lineSpacingType === 'Percent' ? fontPx * line / 100
    : properties.lineSpacingType === 'Fixed' ? line
    : properties.lineSpacingType === 'SpaceOnly' ? fontPx + line
    : properties.lineSpacingType === 'Minimum' ? Math.max(fontPx, line) : NaN;
}

export function paragraphCss(properties: Record<string, any>, char?: Record<string, any>, cap?: number, includeSpacing = true): string {
  const fontPx = Number(char?.fontSize) / 75;
  const line = Number(properties.lineSpacing);
  // The query API mixes units: margins/indent and fixed leading are
  // already physical pixels, but before/after retain ParaShape's 2x scale.
  // See build_para_properties_json and style_resolver at the corpus revision.
  const height = physicalLineHeight(properties, char);
  const indent = Number(properties.indent) || 0;
  const left = (Number(properties.marginLeft) || 0) + Math.max(0, -indent);
  const spacing = Number.isFinite(cap) && Number.isFinite(height) && height > cap! && cap! >= fontPx ? `${cap}px`
    : properties.lineSpacingType === 'Percent' ? `${line}%` : Number.isFinite(height) ? `${height}px` : '';
  return `margin-left:${left}px;margin-right:${(Number(properties.marginRight) || 0)}px;text-indent:${indent}px;` +
    (['left', 'center', 'right', 'justify'].includes(properties.alignment) ? `text-align:${properties.alignment};` : '') +
    (includeSpacing ? `margin-top:${(Number(properties.spacingBefore) || 0) / 2}px;margin-bottom:${(Number(properties.spacingAfter) || 0) / 2}px;` : '') +
    (char?.fontSize ? `font-size:${char.fontSize / 100}pt;font-family:${escape(JSON.stringify(char.fontFamily || 'serif'))};` : '') +
    (spacing ? `line-height:${spacing};` : '') +
    // The native numeric API uses 0 for Korean word boundaries and 1 for
    // character boundaries. Do not infer this from HWPX's Latin enum names.
    (properties.koreanBreakUnit === 0 ? 'word-break:keep-all;' : '');
}

export function sourceLinePitches(doc: HwpDocument): Map<string, number> {
  const lines = new Map<string, { y: number; h: number; page: number; charStart?: number }[]>();
  const texts = new Map<string, () => string>();
  const pitches = new Map<string, number>();
  for (let page = 0; page < doc.pageCount(); page++) for (const run of JSON.parse(doc.getPageTextLayout(page)).runs ?? []) {
    if (run.parentParaIdx === undefined && typeof run.secIdx === 'number' && typeof run.paraIdx === 'number') pitches.set(`flow:${run.secIdx}:${run.paraIdx}`, 1);
    if (run.headerFooter || run.paraIdx === 0xffffffff || typeof run.secIdx !== 'number' || typeof run.paraIdx !== 'number') continue;
    const key = run.parentParaIdx === undefined ? `${run.secIdx}:${run.paraIdx}` : cellParagraphKey(run.secIdx, run.parentParaIdx, run.cellPath?.length ? run.cellPath : [{ controlIndex: run.controlIdx, cellIndex: run.cellIdx, cellParaIndex: run.cellParaIdx }]);
    if (!texts.has(key)) texts.set(key, () => run.parentParaIdx === undefined
      ? doc.getTextRange(run.secIdx, run.paraIdx, 0, doc.getParagraphLength(run.secIdx, run.paraIdx))
      : doc.getTextInCellByPath(run.secIdx, run.parentParaIdx, JSON.stringify(run.cellPath?.length ? run.cellPath : [{ controlIndex: run.controlIdx, cellIndex: run.cellIdx, cellParaIndex: run.cellParaIdx }]), 0, 0xffffffff));
    if (!pitches.has(`top:${key}`)) { pitches.set(`top:${key}`, run.y); pitches.set(`page:${key}`, page); }
    pitches.set(`height:${key}`, Math.max(pitches.get(`height:${key}`) ?? 0, run.h));
    if (!run.text?.trim()) continue;
    const group = lines.get(key) ?? [], previous = group.at(-1);
    if (!previous || previous.page !== page || Math.abs(previous.y - run.y) > Math.min(previous.h, run.h) * .45) group.push({ y: run.y, h: run.h, page, charStart: run.charStart });
    lines.set(key, group);
  }
  for (const [key, group] of lines) {
    pitches.set(`lines:${key}`, group.length);
    let text: string[] | undefined;
    const values = group.slice(1).flatMap((line, i) => {
      const previous = group[i];
      if (line.page !== previous.page) return [];
      const distance = line.y - previous.y, height = Math.max(line.h, previous.h);
      // Retain the existing rejection of large control/object gaps before
      // normalizing blank lines; they must not become new pitch candidates.
      if (distance < height * .9 || distance >= height * 4) return [];
      let steps = 1;
      if (line.charStart !== undefined && previous.charStart !== undefined && line.charStart > previous.charStart) {
        try {
          // Paint-run character starts use native Unicode scalar offsets.
          // JS string slicing would stop early after supplementary Unicode
          // and mistake several explicit line pitches for one enlarged line.
          text ??= Array.from(texts.get(key)?.() ?? '');
          // Empty explicit lines have no painted run. Their combined distance
          // is several line pitches, rather than an enlarged single pitch.
          steps = Math.max(1, text.slice(previous.charStart, line.charStart).filter(char => char === '\n').length);
        } catch { /* Optional text metrics retain the measured fallback. */ }
      }
      const value = distance / steps;
      return value >= height * .9 && value < height * 4 ? [value] : [];
    }).sort((a, b) => a - b);
    if (values.length) pitches.set(key, Math.round(values[Math.floor(values.length / 2)] * 100) / 100);
  }
  return pitches;
}

/** Shared cell line-box semantics, independent of the table nesting depth. */
export function cellParagraphCss(properties: Record<string, any>, char: Record<string, any>, options: {
  pitch?: number; inkHeight?: number; cap?: number; final: boolean; lines?: number; simple: boolean;
}): string {
  let pitch = options.pitch;
  let policy = '';
  if (options.final && options.lines !== undefined && options.simple) {
    const inkHeight = options.inkHeight ?? Number(char.fontSize) / 75;
    const requested = pitch ?? physicalLineHeight(properties, char);
    if (inkHeight > 0 && requested > inkHeight) {
      if (options.lines <= 1) pitch = inkHeight;
      const font = Number(char.fontSize) / 75;
      if (['Percent', 'SpaceOnly', 'Minimum', 'Fixed'].includes(properties.lineSpacingType) && Math.abs(inkHeight - font) < .5) {
        pitch = options.lines <= 1 ? inkHeight : requested;
        policy = `--hwp-cell-line:${encodeHwpCellLine({ rule: properties.lineSpacingType, font, advance: requested, end: inkHeight, base: options.lines <= 1 ? inkHeight : requested })};`;
      }
    }
  }
  return paragraphCss(properties, char, options.cap, false) + (pitch ? `line-height:${pitch}px;` : '') + policy;
}
