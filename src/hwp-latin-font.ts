export const hwpLatinFontAttribute = 'data-hwp-latin-font';
export const hwpLatinFontNamespace = 'urn:wordstudio:hwp-latin-font:1';

function primaryFace(value: unknown) {
  if (typeof value !== 'string') return '';
  const first = value.split(',')[0].trim();
  return first.replace(/^(["'])(.*)\1$/s, '$2');
}

export function hwpLatinFont(hangul: unknown, latin: unknown) {
  const faces = [primaryFace(hangul), primaryFace(latin)];
  if (faces[0] !== faces[1] &&
      faces.every(face => face && face.length <= 500 && !/[\x00-\x1f\x7f]/.test(face)))
    return JSON.stringify(faces);
}

/** A current face change takes precedence over the source Latin slot. */
export function activeHwpLatinFont(value: unknown, family: unknown): string | undefined {
  if (typeof value !== 'string' || value.length > 2000) return;
  try {
    const faces = JSON.parse(value);
    if (Array.isArray(faces) && faces.length === 2 && hwpLatinFont(faces[0], faces[1]) === value &&
        primaryFace(family) === faces[0]) return faces[1];
  } catch { /* Ignore unrecognized source typography. */ }
}
