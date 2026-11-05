import { legacySymbolText } from './document-symbols';

const fonts = new Map<string, string>();
// Mapped Unicode must use a Unicode font, never the original legacy encoding
// a second time. Existing bundled fallbacks cover the complete symbol maps.
const unicodeFonts = 'Arial,"Studio Symbols 2","Studio Symbols","Studio Math","Studio Text","Studio Devanagari","Studio Emoji",sans-serif';

export function wordListFontName(value: unknown): string | null {
  return typeof value === 'string' && value.trim().length > 0 && value.length <= 200 &&
    !/[;{}\\\u0000-\u001f]/.test(value) && !/url\s*\(|var\s*\(/i.test(value) ? value : null;
}

/** Rendering is separate from the saved lvlText/rFonts. Unknown legacy codes
 * remain unchanged rather than inventing a replacement or losing their font. */
export function wordListMarker(label: string, sourceFont: unknown) {
  const font = wordListFontName(sourceFont);
  if (!font) return { glyph: null, fontFamily: '' };
  const glyph = legacySymbolText(font, label);
  let family = fonts.get(font);
  if (!family) {
    const style = document.createElement('span').style;
    style.fontFamily = JSON.stringify(font) + ",Arial,sans-serif"; family = style.fontFamily;
    if (fonts.size >= 128) fonts.delete(fonts.keys().next().value!);
    fonts.set(font, family);
  }
  return { glyph, fontFamily: glyph ? unicodeFonts : family };
}
