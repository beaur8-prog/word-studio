/** HWP cells omit the final line's extra leading. Keep that terminal height
 * separate from the baseline advance so editing a formerly single line does
 * not apply its compact height to every new line. Values are physical pixels. */
export interface HwpCellLine {
  rule: 'Percent' | 'SpaceOnly' | 'Minimum' | 'Fixed';
  font: number;
  advance: number;
  end: number;
  base: number;
  authoredTerminal?: HwpAuthoredCellTerminal;
}
export type HwpAuthoredCellTerminal = { text: string; width: number; row: number; middle?: true };
export const hwpCellLineNamespace = 'urn:word-studio:hwp-cell-lines:1';
const rules = ['Percent', 'SpaceOnly', 'Minimum', 'Fixed'];
export function parseHwpCellLine(value: unknown): HwpCellLine | null {
  if (typeof value !== 'string' || value.length > 200) return null;
  const [version, rule, ...values] = value.trim().split(/\s+/);
  const numbers = values.slice(0, 4).map(Number);
  const middle = values[7] === 'middle';
  const expectedValues = version === '2' ? (middle ? 8 : 7) : 4;
  if (version !== '1' && version !== '2' || !rules.includes(rule) || values.length !== expectedValues ||
    numbers.some(n => !Number.isFinite(n) || n <= 0 || n > 2000)) return null;
  const [font, advance, end, base] = numbers;
  if (end > advance || Math.abs(base - end) > .01 && Math.abs(base - advance) > .01) return null;
  const result: HwpCellLine = { rule: rule as HwpCellLine['rule'], font, advance, end, base };
  if (version === '2') {
    const [text, width, row] = [values[4], Number(values[5]), Number(values[6])] as const;
    if (!/^[\da-f]{16}$/.test(text) || !(width > 0 && width <= 2000) || !(row >= end && row <= 2000) ||
        Math.abs(base - end) > .01) return null;
    result.authoredTerminal = { text, width, row, ...(middle ? { middle: true as const } : {}) };
  }
  return result;
}
export function encodeHwpCellLine(line: HwpCellLine): string {
  const rounded = (n: number) => Math.round(n * 10000) / 10000;
  const policy = [line.font, line.advance, line.end, line.base].map(rounded).join(' ');
  const authored = line.authoredTerminal;
  return authored ? `2 ${line.rule} ${policy} ${authored.text} ${rounded(authored.width)} ${rounded(authored.row)}` +
    (authored.middle ? ' middle' : '') : `1 ${line.rule} ${policy}`;
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
    : line.rule === 'SpaceOnly' ? font + line.advance - line.font
    : line.rule === 'Fixed' ? line.advance : Math.max(font, line.advance);
}
/** Read the document's font sizes, not prior view decorations. A cleared text
 * mark must inherit the stored paragraph size rather than a stale view size. */
export function hwpCellFontProfile(node: { attrs?: Record<string, any>; content?: readonly any[] }, line: HwpCellLine) {
  const inherited = cssPixels(node.attrs?.paragraphSize) ?? line.font, sizes: number[] = [];
  for (const run of node.content ?? []) if (run.type === 'text' && run.text?.length) {
    const style = run.marks?.find((m: any) => m.type === 'textStyle');
    sizes.push(cssPixels(style?.attrs?.fontSize) ?? inherited);
  }
  const font = sizes.reduce((max, size) => Math.max(max, size), sizes.length ? 0 : inherited);
  return { inherited, font, uniform: sizes.every(size => Math.abs(size - font) <= .05) };
}
export const hwpCellFont = (node: { attrs?: Record<string, any>; content?: readonly any[] }, line: HwpCellLine) => hwpCellFontProfile(node, line).font;
