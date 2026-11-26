export type WordParagraphBottomBorder =
  | { style: 'single'; size: number; space?: number; color?: string }
  | { style: 'none' | 'nil' };

/** Only a bounded OOXML line description enters the editor model. */
export function parseWordParagraphBottomBorder(value: unknown): WordParagraphBottomBorder | null {
  if (typeof value !== 'string' || value.length > 128) return null;
  try {
    const border: WordParagraphBottomBorder = JSON.parse(value);
    if (!border || typeof border !== 'object' || Array.isArray(border)) return null;
    if (border.style === 'none' || border.style === 'nil')
      return Object.keys(border).length === 1 ? { style: border.style } : null;
    if (border.style !== 'single' ||
        Object.keys(border).some(key => !['style', 'size', 'space', 'color'].includes(key)) ||
        !Number.isInteger(border.size) || border.size < 2 || border.size > 96 ||
        (border.space !== undefined && (!Number.isInteger(border.space) || border.space < 0 || border.space > 31)) ||
        (border.color !== undefined && (typeof border.color !== 'string' || !/^(?:[\da-f]{6}|auto)$/i.test(border.color)))) return null;
    return { style: 'single', size: border.size, ...(border.space !== undefined ? { space: border.space } : {}),
      ...(border.color ? { color: border.color.toLowerCase() === 'auto' ? 'auto' : border.color.toUpperCase() } : {}) };
  } catch { return null; }
}
