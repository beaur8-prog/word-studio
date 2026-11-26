import { cssPixels } from './hwp-cell-line';

/** HWP percentage leading follows each visual line's font sizes. The saved
 * physical spacing is only an import baseline, not a fixed paragraph strut. */
// font is the original largest em size; base is the physical import pitch.
export interface HwpBodyLine { percent: number; font: number; base: number }
export const hwpBodyLineNamespace = 'urn:wordstudio:hwp-body-line:1';
export function parseHwpBodyLine(value: unknown): HwpBodyLine | null {
  if (typeof value !== 'string' || value.length > 160) return null;
  const [version, ...parts] = value.trim().split(/\s+/), values = parts.map(Number);
  if (version !== '1' || values.length !== 3 || values.some(n => !Number.isFinite(n) || n <= 0 || n > 2000)) return null;
  const [percent, font, base] = values;
  return { percent, font, base };
}
export const encodeHwpBodyLine = (line: HwpBodyLine) =>
  `1 ${[line.percent, line.font, line.base].map(n => Math.round(n * 10000) / 10000).join(' ')}`;
export function activeHwpBodyLine(attrs?: Record<string, any>) {
  const line = parseHwpBodyLine(attrs?.hwpBodyLine), base = cssPixels(attrs?.lineSpacing);
  return line && base !== null && Math.abs(base - line.base) < .06 ? line : null;
}
export function hwpBodyTextLine(node: { attrs?: Record<string, any>; content?: readonly any[] }) {
  const line = activeHwpBodyLine(node.attrs);
  return line && node.content?.some(n => n.text?.trim()) && node.content.every(n => n.type === 'text' &&
    !n.marks?.some((mark: any) => mark.type === 'textStyle' && mark.attrs?.lineHeight)) ? line : null;
}
/** Keep the imported compatibility pitch until a font edit grows beyond the
 * original maximum. A Word exact line must then accommodate the larger glyphs. */
export function hwpBodyFallback(node: { attrs?: Record<string, any>; content?: readonly any[] }, line: HwpBodyLine) {
  const inherited = cssPixels(node.attrs?.paragraphSize) ?? line.font;
  const sizes = (node.content ?? []).filter(n => n.type === 'text' && n.text?.length).map(n =>
    cssPixels(n.marks?.find((m: any) => m.type === 'textStyle')?.attrs?.fontSize) ?? inherited);
  const font = sizes.reduce((maximum, size) => Math.max(maximum, size), 0);
  return font > line.font + .1 ? Math.max(line.base, font * line.percent / 100) : line.base;
}
