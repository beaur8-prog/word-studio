/** Standard OOXML vertical alignment of mixed-size runs on a paragraph line. */
export function paragraphTextAlignment(value: unknown): string | null {
  return typeof value === 'string' && ['auto', 'baseline', 'top', 'center', 'bottom'].includes(value) ? value : null;
}
