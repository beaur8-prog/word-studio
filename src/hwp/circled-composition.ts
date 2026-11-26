/** The Hancom spread circle uses two positional forms of each decimal digit.
 * Only compositions with a single corresponding Unicode circle are editable
 * as ordinary text in both Studio and standard DOCX. */
export function circledComposition(value: string): string | undefined {
  const digits = [...value].map(character => {
    const code = character.codePointAt(0)!;
    if (code >= 0xf0289 && code <= 0xf0291) return code - 0xf0288;
    if (code >= 0xf0292 && code <= 0xf029b) return code - 0xf0292;
    return -1;
  });
  if (digits.length !== 2 || digits[0] < 1 || digits[1] < 0) return;
  const number = 10 * digits[0] + digits[1];
  if (number <= 20) return String.fromCodePoint(0x2460 + number - 1);
  if (number <= 35) return String.fromCodePoint(0x3251 + number - 21);
  if (number <= 50) return String.fromCodePoint(0x32b1 + number - 36);
}

export function circledCompositionHtml(circle: string, nativeSize: unknown) {
  if (typeof nativeSize === 'number' && Number.isFinite(nativeSize) && nativeSize > 0 && nativeSize <= 28800)
    return `<span style="font-size:${nativeSize / 100}pt">${circle}</span>`;
  return '<span>' + circle + '</span>';
}
