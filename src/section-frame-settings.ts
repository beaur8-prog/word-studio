import type { PageSettings } from './types';

function canonical(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(canonical);
  if (value && typeof value === 'object') return Object.fromEntries(Object.entries(value)
    .filter(([, value]) => value !== undefined).sort(([a], [b]) => a.localeCompare(b))
    .map(([key, value]) => [key, canonical(value)]));
  return value;
}

/** A continuous boundary can share an editable paper frame only when every
 * effective setting, including stories and columns, remains the same. Keep
 * the source boundary in the model and DOCX; this is a view grouping rule. */
export function sameSectionFrameSettings(a: PageSettings, b: PageSettings) {
  const { breakType: _a, ...first } = a, { breakType: _b, ...second } = b;
  return JSON.stringify(canonical(first)) === JSON.stringify(canonical(second));
}
