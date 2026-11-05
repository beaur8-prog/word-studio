/** The original installed face always wins. Optional fallback faces are view
 * resources: requested font names, source point sizes and DOCX stay unchanged. */
export type FontFallbackFamily = {
  family: string;
  faces: readonly {
    names: readonly string[]; weight: string; style: string; url: string;
    ascent: string; descent: string;
  }[];
};

// Unmodified TeX Gyre Adventor 2.501, GFL/LPPL. Advance widths were compared
// against real Word Century Gothic output, including a separate four-face
// probe. CSS overrides reproduce its hhea strut, not a page-count multiplier:
// regular/italic 1989/2048, bold 2032/2048, bold italic 2060/2048; descent 451/2048.
// See electron/licenses/TeXGyreAdventor-README.txt for scope and known differences.
const families: readonly FontFallbackFamily[] = [{
  family: 'Century Gothic',
  faces: [
    { names: ['Century Gothic', 'CenturyGothic'], weight: '400', style: 'normal',
      url: new URL('./assets/fonts/texgyreadventor-regular.otf', import.meta.url).href,
      ascent: '97.119140625%', descent: '22.021484375%' },
    { names: ['Century Gothic Bold', 'CenturyGothic-Bold'], weight: '700', style: 'normal',
      url: new URL('./assets/fonts/texgyreadventor-bold.otf', import.meta.url).href,
      ascent: '99.21875%', descent: '22.021484375%' },
    { names: ['Century Gothic Italic', 'CenturyGothic-Italic'], weight: '400', style: 'italic',
      url: new URL('./assets/fonts/texgyreadventor-italic.otf', import.meta.url).href,
      ascent: '97.119140625%', descent: '22.021484375%' },
    { names: ['Century Gothic Bold Italic', 'CenturyGothic-BoldItalic'], weight: '700', style: 'italic',
      url: new URL('./assets/fonts/texgyreadventor-bolditalic.otf', import.meta.url).href,
      ascent: '100.5859375%', descent: '22.021484375%' },
  ],
}];

// Only our measured fallback faces opt into fractional Word line metrics.
// Native faces and unrelated web fonts keep their existing normal line boxes.
const metricFaces = new Map<FontFace, { family: string; weight: string; style: string; ratio: number }>();
export const wordFontMetricsChanged = 'studio-word-font-metrics-changed';
export function wordFallbackLineHeight(css: CSSStyleDeclaration, text: string) {
  if (!/^[\u0020-\u024f\u2000-\u206f]*$/.test(text)) return null;
  const family = css.fontFamily.split(',')[0].replace(/["']/g, '').trim().toLowerCase();
  for (const [face, metric] of metricFaces) {
    if (metric.family === family && metric.weight === css.fontWeight && metric.style === css.fontStyle &&
      face.status === 'loaded' && document.fonts.has(face)) return parseFloat(css.fontSize) * metric.ratio;
  }
  return null;
}

export async function initializeWordFontFallbacks(profiles: readonly FontFallbackFamily[] = families) {
  let released = false;
  const changed = () => { if (released) document.dispatchEvent(new Event(wordFontMetricsChanged)); };
  const initialized = Promise.all(profiles.flatMap(profile => profile.faces.map(async face => {
    const descriptor = { weight: face.weight, style: face.style };
    try {
      // FontFaceSet.check() also succeeds for a missing system family. Loading
      // a local-only probe really tests availability without OS font permission.
      const native = new FontFace(profile.family,
        face.names.map(name => `local(${JSON.stringify(name)})`).join(','), descriptor);
      await native.load();
      // A CSS family with only a missing bold face would otherwise shadow its
      // installed regular sibling. Register the unchanged native face too.
      document.fonts.add(native); changed();
    } catch {
      try {
        const fallback = new FontFace(profile.family, `url(${JSON.stringify(face.url)})`, {
          ...descriptor, display: 'swap', ascentOverride: face.ascent,
          descentOverride: face.descent, lineGapOverride: '0%',
        });
        // Register before the editor mounts, but download only if used. The
        // existing loadingdone handlers remeasure lines and page furniture.
        metricFaces.set(fallback, { family: profile.family.toLowerCase(), weight: face.weight, style: face.style,
          ratio: (parseFloat(face.ascent) + parseFloat(face.descent)) / 100 });
        document.fonts.add(fallback); changed();
      } catch { /* An optional font cannot prevent opening or recovering documents. */ }
    }
  })));
  // A stalled optional local-font probe must not leave the startup window
  // blank. Late results still register and notify the measurement owners.
  let timer: ReturnType<typeof setTimeout> | undefined;
  try { await Promise.race([initialized, new Promise<void>(resolve => { timer = setTimeout(resolve, 1500); })]); }
  finally { clearTimeout(timer); released = true; }
}
