import { parseWordParagraphBottomBorder } from './word-paragraph-border';

export const paragraphBorderBoxAttribute = 'data-paragraph-border-box';
export const paragraphBorderBoxStart = 'data-paragraph-border-box-start';
export const paragraphBorderBoxEnd = 'data-paragraph-border-box-end';

export type ParagraphBorderBox = { size: number; color: string; leftSpace: number; rightSpace: number };

export function parseParagraphBorderBox(value: unknown): ParagraphBorderBox | null {
  if (typeof value !== 'string' || value.length > 128) return null;
  try {
    const box: ParagraphBorderBox = JSON.parse(value);
    if (!box || typeof box !== 'object' || Array.isArray(box) ||
        Object.keys(box).some(key => !['size', 'color', 'leftSpace', 'rightSpace'].includes(key))) return null;
    const line = parseWordParagraphBottomBorder(JSON.stringify({ style: 'single', size: box.size, color: box.color }));
    const validSpace = (space: number) => Number.isInteger(space) && space >= 0 && space <= 31;
    if (line?.style !== 'single' || !validSpace(box.leftSpace) || !validSpace(box.rightSpace)) return null;
    return { size: line.size, color: line.color ?? 'auto', leftSpace: box.leftSpace, rightSpace: box.rightSpace };
  } catch { return null; }
}

export function paragraphBorderBoxOOXML(box: ParagraphBorderBox) {
  const edge = { style: 'single' as const, size: box.size, color: box.color, space: 0 };
  return { top: edge, left: { ...edge, space: box.leftSpace }, bottom: edge, right: { ...edge, space: box.rightSpace },
    between: { style: 'nil' as const } };
}
