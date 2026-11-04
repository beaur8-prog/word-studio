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

export async function initializeWordFontFallbacks(profiles: readonly FontFallbackFamily[] = families) {
  await Promise.all(profiles.flatMap(profile => profile.faces.map(async face => {
    const descriptor = { weight: face.weight, style: face.style };
    try {
      // FontFaceSet.check() also succeeds for a missing system family. Loading
      // a local-only probe really tests availability without OS font permission.
      const native = new FontFace(profile.family,
        face.names.map(name => `local(${JSON.stringify(name)})`).join(','), descriptor);
      await native.load();
      // A CSS family with only a missing bold face would otherwise shadow its
      // installed regular sibling. Register the unchanged native face too.
      document.fonts.add(native);
    } catch {
      try {
        const fallback = new FontFace(profile.family, `url(${JSON.stringify(face.url)})`, {
          ...descriptor, display: 'swap', ascentOverride: face.ascent,
          descentOverride: face.descent, lineGapOverride: '0%',
        });
        // Register before the editor mounts, but download only if used. The
        // existing loadingdone handlers remeasure lines and page furniture.
        document.fonts.add(fallback);
      } catch { /* An optional font cannot prevent opening or recovering documents. */ }
    }
  })));
}
