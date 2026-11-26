/** Hancom's explicit section option hides at most two otherwise empty lines
 * at an automatic page boundary. Decorations and authored spacing still
 * occupy paper; do not mistake them for an invisible blank line. */
export const hiddenEmptyLineClass = 'pagination-hidden-empty-line';
export const emptyLineDocxProperty = 'WordStudio.HwpEmptyLines';
export const emptyLineCompatibilityWarning = '원본의 쪽 시작 빈 줄 감추기를 적용합니다. DOCX를 다른 앱에서 열면 숨긴 빈 줄이 다시 나타나 쪽 수가 늘 수 있습니다.';

export function canHideEmptyPageLine(paragraph: HTMLElement): boolean {
  if (paragraph.tagName !== 'P' || paragraph.textContent?.trim() ||
    paragraph.querySelector('img,svg,table,hr,br:not(.ProseMirror-trailingBreak),[data-note-reference],[data-page-number]')) return false;
  for (const element of [paragraph, ...paragraph.querySelectorAll<HTMLElement>('*')]) {
    const css = getComputedStyle(element);
    if (css.backgroundColor !== 'rgba(0, 0, 0, 0)' || css.backgroundImage !== 'none' ||
      ['paddingTop', 'paddingBottom', 'borderTopWidth', 'borderBottomWidth', 'marginTop', 'marginBottom']
        .some(key => Math.abs(parseFloat(css[key as 'paddingTop']) || 0) > .01)) return false;
  }
  return true;
}

/** After the two allowed hidden lines, a short empty paragraph can still
 * fit at the bottom by its em. Leave its full advance in document flow. */
export function nativeEmptyLineContentHeight(paragraph: HTMLElement, scale: number): number | null {
  if (paragraph.textContent || !canHideEmptyPageLine(paragraph) || !(scale > 0)) return null;
  const style = getComputedStyle(paragraph);
  const font = parseFloat(style.fontSize), advance = parseFloat(style.lineHeight);
  if (!(font > 0 && advance > font && advance <= font * 2) ||
    Math.abs(paragraph.getBoundingClientRect().height / scale - advance) > .07) return null;
  return font;
}
