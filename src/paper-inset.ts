/** Paper insets are CSS pixels. An authored zero is a real edge, not an
 * absent value; keep fallback behaviour only for missing/invalid metadata. */
export function paperInset(value: string | undefined, fallback: number): number {
  const number = value === undefined || value.trim() === '' ? NaN : Number(value);
  return Number.isFinite(number) && number >= 0 ? number : fallback;
}
