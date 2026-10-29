/** HWP cells omit the final line's extra leading. Keep that terminal height
 * separate from the baseline advance so editing a formerly single line does
 * not apply its compact height to every new line. Values are physical pixels. */
export interface HwpCellLine {
  rule: 'Percent' | 'SpaceOnly' | 'Minimum';
  font: number;
  advance: number;
  end: number;
  base: number;
}
export const hwpCellLineNamespace = 'urn:word-studio:hwp-cell-lines:1';
const rules = ['Percent', 'SpaceOnly', 'Minimum'];
export function parseHwpCellLine(value: unknown): HwpCellLine | null {
  if (typeof value !== 'string' || value.length > 200) return null;
  const [version, rule, ...values] = value.trim().split(/\s+/);
  const numbers = values.map(Number);
  if (version !== '1' || !rules.includes(rule) || numbers.length !== 4 ||
    numbers.some(n => !Number.isFinite(n) || n <= 0 || n > 2000)) return null;
  const [font, advance, end, base] = numbers;
  if (end > advance || Math.abs(base - end) > .01 && Math.abs(base - advance) > .01) return null;
  return { rule: rule as HwpCellLine['rule'], font, advance, end, base };
}
export function encodeHwpCellLine(line: HwpCellLine): string {
  return `1 ${line.rule} ${[line.font, line.advance, line.end, line.base].map(n => Math.round(n * 10000) / 10000).join(' ')}`;
}
export function cssPixels(value: unknown): number | null {
  if (typeof value !== 'string' || !/^\d+(?:\.\d+)?(?:px|pt)$/.test(value)) return null;
  const number = parseFloat(value) * (value.endsWith('pt') ? 4 / 3 : 1);
  return number > 0 && number <= 2000 ? number : null;
}
/** Changing line spacing explicitly retires the imported line policy. */
export function activeHwpCellLine(attrs?: Record<string, any>): HwpCellLine | null {
  const line = parseHwpCellLine(attrs?.hwpCellLine), base = cssPixels(attrs?.lineSpacing);
  return line && base !== null && Math.abs(line.base - base) < .06 ? line : null;
}
export function hwpCellAdvance(line: HwpCellLine, font: number): number {
  return line.rule === 'Percent' ? line.advance * font / line.font
    : line.rule === 'SpaceOnly' ? font + line.advance - line.font : Math.max(font, line.advance);
}
/** Read the document's font sizes, not prior view decorations. A cleared text
 * mark must inherit the stored paragraph size rather than a stale view size. */
export function hwpCellFont(node: { attrs?: Record<string, any>; content?: readonly any[] }, line: HwpCellLine): number {
  const inherited = cssPixels(node.attrs?.paragraphSize) ?? line.font, sizes: number[] = [];
  for (const run of node.content ?? []) if (run.type === 'text' && run.text?.trim()) {
    const style = run.marks?.find((m: any) => m.type === 'textStyle');
    sizes.push(cssPixels(style?.attrs?.fontSize) ?? inherited);
  }
  return Math.max(...(sizes.length ? sizes : [inherited]));
}
