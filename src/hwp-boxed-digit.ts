export const hwpBoxedDigitAttribute = 'data-hwp-boxed-digit';
export const hwpBoxedDigitNamespace = 'urn:word-studio:hwp-boxed-digit:1';

export function parseHwpBoxedDigit(value: unknown) {
  const match = typeof value === 'string' && /^([1-9]\d{0,2}):([0-9])$/.exec(value);
  if (!match) return;
  const size = Number(match[1]);
  if (size < 8 || size > 576) return;
  return { size, digit: match[2] };
}

export function boxedDigitHtml(digit: string | undefined, nativeSize: unknown) {
  if (!digit || typeof nativeSize !== 'number' || !Number.isFinite(nativeSize)) return;
  const encoded = String(Math.round(nativeSize / 50)) + ':' + digit;
  if (parseHwpBoxedDigit(encoded)) return '<span ' + hwpBoxedDigitAttribute + '="' + encoded + '">' + digit + '</span>';
}

export function hwpBoxedDigitBorder(size: number) {
  return { style: 'single' as const, size: Math.max(2, Math.min(96, Math.round(size * .16))), space: Math.round(size * .06) };
}
