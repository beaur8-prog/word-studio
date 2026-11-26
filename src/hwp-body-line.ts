import { cssPixels } from './hwp-cell-line';
import { hwpWrapFingerprint } from './hwp-text-advances';
import { imageFloating } from './image-floating';

/** HWP percentage leading follows each visual line's font sizes. The saved
 * physical spacing is only an import baseline, not a fixed paragraph strut. */
// font is the original largest em size; base is the physical import pitch.
export type HwpTerminalLineFit = { baseline: number; remainder: number; width: number; fingerprint: string; preceding: string };
export type HwpColumnEndFit = HwpTerminalLineFit & { hang?: true };
export type HwpStoredMixedLinePair = readonly [first: { pitch: number; baseline: number }, second: { pitch: number; baseline: number }];
export type HwpPaperHeading = { fingerprint: string; width: number };
export interface HwpBodyLine { percent: number; font: number; base: number; maxFont?: number; align?: 'top' | 'mixed-top' | 'column-end'; cell?: true; empty?: true; storedSingleLine?: true; storedUniformPair?: true; mixedLinePair?: HwpStoredMixedLinePair; precisePitch?: true; columnFit?: HwpColumnEndFit; bodyEndFit?: HwpTerminalLineFit; paperHeading?: HwpPaperHeading }
export const hwpBodyLineNamespace = 'urn:wordstudio:hwp-body-line:1';
const alignments = [undefined, 'top', 'mixed-top', 'column-end'] as const;

function validStoredUniformPair(percent: number, font: number, base: number) {
  return percent > 100 && percent <= 200 && font > 0 && base > font && base <= 240 &&
    Math.abs(font * percent / 100 - base) <= .06;
}
export function parseHwpBodyLine(value: unknown): HwpBodyLine | null {
  if (typeof value !== 'string' || value.length > 160) return null;
  const [version, ...parts] = value.trim().split(/\s+/);
  if (version === '13') {
    const [percent, font, base, width] = parts.slice(0, 4).map(Number);
    const fingerprint = parts[4];
    return parts.length === 5 && validStoredUniformPair(percent, font, base) && Number.isFinite(width) &&
      width > font && width <= 10000 && /^[0-9a-f]{16}$/.test(fingerprint)
      ? { percent, font, base, paperHeading: { width, fingerprint } } : null;
  }
  if (version === '12') {
    const [percent, font, base] = parts.map(Number);
    return parts.length === 3 && validStoredUniformPair(percent, font, base)
      ? { percent, font, base, align: 'mixed-top', storedUniformPair: true } : null;
  }
  if (version === '10') {
    const [percent, font, base, firstPitch, firstBaseline, secondPitch, secondBaseline] = parts.map(Number);
    if (parts.length !== 7 || ![percent, font, base, firstPitch, firstBaseline, secondPitch, secondBaseline]
      .every(value => Number.isFinite(value) && value > 0 && value <= 2000) ||
      percent <= 100 || base <= font || Math.abs(firstPitch - base) > .08 || firstPitch <= secondPitch ||
      firstBaseline > font || firstBaseline <= secondBaseline || secondBaseline >= secondPitch) return null;
    return { percent, font, base, mixedLinePair: [
      { pitch: firstPitch, baseline: firstBaseline }, { pitch: secondPitch, baseline: secondBaseline },
    ] };
  }
  if (version === '8' || version === '11') {
    const body = version === '11';
    if (parts.length !== 9 && (body || parts.length !== 10)) return null;
    const [percent, font, base, maxFont, baseline, remainder, width] = parts.slice(0, 7).map(Number);
    const [fingerprint, preceding, hang] = parts.slice(7);
    if (![percent, font, base, maxFont, baseline, remainder, width].every(n => Number.isFinite(n) && n > 0 && n <= 2000) ||
        percent <= 100 || base <= font || maxFont < font || baseline > font || remainder < font - .16 ||
        remainder >= base || width < font * 4 || !/^[0-9a-f]{16}$/.test(fingerprint) ||
        !/^[0-9a-f]{16}$/.test(preceding) || hang !== undefined && hang !== '1') return null;
    const fit = { baseline, remainder, width, fingerprint, preceding };
    return body ? { percent, font, base, maxFont, align: 'mixed-top', bodyEndFit: fit }
      : { percent, font, base, maxFont, align: 'column-end', columnFit: { ...fit, ...(hang ? { hang: true } : {}) } };
  }
  const values = parts.map(Number);
  if (!/^[1-79]$/.test(version) || values.length !== 3 && !(version === '4' && values.length === 4) ||
      values.some(n => !Number.isFinite(n) || n <= 0 || n > 2000)) return null;
  const [percent, font, base, maxFont] = values;
  if (version === '9') return percent > 100 && base > font && Math.abs(base - font * percent / 100) <= .05
    ? { percent, font, base, precisePitch: true } : null;
  if (version === '5') return percent > 100 && base > font ? { percent, font, base, cell: true } : null;
  if (version === '6') return percent > 100 && percent <= 200 && base > font && base <= font * 2
    ? { percent, font, base, empty: true } : null;
  if (version === '7') return percent > 100 && base > font ? { percent, font, base, storedSingleLine: true } : null;
  const align = alignments[Number(version) - 1];
  if (align) {
    if (percent <= 100 || base <= font || maxFont !== undefined && maxFont < font) return null;
    return { percent, font, base, align, ...(maxFont === undefined ? {} : { maxFont }) };
  }
  return { percent, font, base };
}
const rounding = (n: number) => Math.round(n * 10000) / 10000;
export function encodeHwpBodyLine(line: HwpBodyLine) {
  if (line.paperHeading) return [13, ...[line.percent, line.font, line.base, line.paperHeading.width].map(rounding),
    line.paperHeading.fingerprint].join(' ');
  const column = line.align === 'column-end' ? line.columnFit : undefined;
  const fit = column ?? (line.align === 'mixed-top' ? line.bodyEndFit : undefined);
  if (fit && line.maxFont !== undefined) {
    return [column ? '8' : '11', ...[line.percent, line.font, line.base, line.maxFont, fit.baseline, fit.remainder, fit.width].map(rounding),
      fit.fingerprint, fit.preceding, ...(column?.hang ? ['1'] : [])].join(' ');
  }
  if (line.mixedLinePair) {
    const [first, second] = line.mixedLinePair;
    return ['10', ...[line.percent, line.font, line.base, first.pitch, first.baseline, second.pitch, second.baseline].map(rounding)].join(' ');
  }
  if (line.storedUniformPair) {
    const percent = rounding(line.percent), font = rounding(line.font), base = rounding(line.base);
    if (validStoredUniformPair(percent, font, base)) return [12, percent, font, base].join(' ');
  }
  const version = line.precisePitch ? 9 : line.storedSingleLine ? 7 : line.empty ? 6 : line.cell ? 5 : alignments.indexOf(line.align) + 1;
  return [version, ...[line.percent, line.font, line.base, ...(line.maxFont === undefined ? [] : [line.maxFont])]
    .map(rounding)].join(' ');
}
export function activeHwpBodyLine(attrs?: Record<string, any>) {
  const line = parseHwpBodyLine(attrs?.hwpBodyLine), base = cssPixels(attrs?.lineSpacing);
  if (line?.precisePitch && attrs?.wordLineRule && attrs.wordLineRule !== 'exact') return null;
  return line && base !== null && Math.abs(base - line.base) < (line.precisePitch ? .001 : .06) ? line : null;
}
export function hwpBodyTextLine(node: { attrs?: Record<string, any>; content?: readonly any[] }): HwpBodyLine | null {
  const line = activeHwpBodyLine(node.attrs);
  if (!line || line.empty || !node.content?.some(n => n.text?.trim())) return null;
  const inherited = cssPixels(node.attrs?.paragraphSize) ?? line.font;
  const valid = node.content.every(n => {
    if (n.type === 'inlineImage' && line.paperHeading) return Boolean(imageFloating(n.attrs?.floating)?.hwpPaperDrawing);
    if (n.type !== 'text') return false;
    const style = n.marks?.find((mark: any) => mark.type === 'textStyle')?.attrs;
    const size = cssPixels(style?.fontSize) ?? inherited;
    return !style?.lineHeight && (line.align !== 'top' || Math.abs(size - line.font) < .06);
  });
  if (!valid) return null;
  if (line.paperHeading && (!node.content.some(n => n.type === 'inlineImage') ||
      hwpWrapFingerprint(node.content.map(n => n.text ?? '').join('')) !== line.paperHeading.fingerprint)) return null;
  const fit = line.columnFit ?? line.bodyEndFit;
  if (fit && hwpWrapFingerprint(node.content.map(n => n.text).join('')) !== fit.fingerprint) {
    const { columnFit: _column, bodyEndFit: _body, ...ordinary } = line;
    return ordinary;
  }
  return line;
}
export function hwpBodyEmptyLine(node: { type?: string; attrs?: Record<string, any>; content?: readonly any[] }) {
  if (node.type !== 'paragraph' || node.content?.length) return null;
  const line = activeHwpBodyLine(node.attrs), font = cssPixels(node.attrs?.paragraphSize);
  return line?.empty && font !== null && Math.abs(font - line.font) < .06 ? line : null;
}

/** Keep the imported compatibility pitch until a font edit grows beyond the
 * original maximum. A Word exact line must then accommodate the larger glyphs. */
export function hwpBodyFallback(node: { attrs?: Record<string, any>; content?: readonly any[] }, line: HwpBodyLine) {
  const inherited = cssPixels(node.attrs?.paragraphSize) ?? line.font;
  const sizes = (node.content ?? []).filter(n => n.type === 'text' && n.text?.length).map(n =>
    cssPixels(n.marks?.find((m: any) => m.type === 'textStyle')?.attrs?.fontSize) ?? inherited);
  const font = sizes.reduce((maximum, size) => Math.max(maximum, size), 0);
  return font > (line.maxFont ?? line.font) + .1 ? Math.max(line.base, font * line.percent / 100) : line.base;
}
