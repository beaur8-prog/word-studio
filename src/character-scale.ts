export const characterScaleProperty = '--word-character-scale';

/** Word's usable ST_TextScale range. Missing or invalid values have no override. */
export function parseCharacterScale(value: unknown): number | null {
  if (typeof value === 'string') {
    const decimal = value.trim();
    if (!/^\+?\d{1,8}$/.test(decimal)) return null;
    value = Number(decimal);
  }
  return typeof value === 'number' && Number.isInteger(value) && value >= 1 && value <= 600 ? value : null;
}
