export const hwpFootnoteLayoutAttribute = 'data-hwp-footnote-layout';
export const hwpFootnoteMarkAttribute = 'data-hwp-footnote-mark';
export const hwpFootnoteNamespace = 'urn:wordstudio:hwp-footnote-layout:1';

export type HwpFootnoteLayout = {
  mark: string;
  above: number;
  below: number;
  between: number;
  separator: number;
  stroke: number;
};

export function hwpFootnoteMark(value: unknown): string | undefined {
  return typeof value === 'string' && /^[1-9]\d{0,4}\)$/.test(value) &&
    Number.parseInt(value) <= 65535 ? value : undefined;
}

export function parseHwpFootnoteLayout(value: unknown): HwpFootnoteLayout | undefined {
  if (typeof value !== 'string' || value.length > 256) return;
  try {
    const layout = JSON.parse(value) as HwpFootnoteLayout;
    if (!layout || typeof layout !== 'object' || !hwpFootnoteMark(layout.mark)) return;
    const between = (n: unknown, min: number, max: number) =>
      typeof n === 'number' && Number.isFinite(n) && n >= min && n <= max;
    if (!(between(layout.above, 0, 160) && between(layout.below, 0, 160) &&
          between(layout.between, 0, 160) && between(layout.separator, 1, 1200) &&
          between(layout.stroke, .1, 6))) return;
    return { mark: layout.mark, above: layout.above, below: layout.below, between: layout.between,
      separator: layout.separator, stroke: layout.stroke };
  } catch { /* The standard note remains available when its layout data is invalid. */ }
}
