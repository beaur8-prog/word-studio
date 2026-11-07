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

/** A partial family maps only measured regular/italic Latin faces. Other
 * weights retain the browser's installed fallback, instead of synthesizing
 * bold from our regular face. An original installed face disables this alias. */
export type RegularFontFallbackFamily = {
  family: string;
  nativeNames: readonly string[];
  faces: readonly { style: 'normal'|'italic'; source: string; ascent: string; descent: string; unicodeRange: string }[];
};
const regularFamilies: readonly RegularFontFallbackFamily[] = [{
  family: 'Gill Sans MT',
  nativeNames: ['Gill Sans MT', 'GillSansMT', 'Gill Sans MT Bold', 'GillSansMT-Bold',
    'Gill Sans MT Italic', 'GillSansMT-Italic', 'Gill Sans MT Bold Italic', 'GillSansMT-BoldItalic'],
  // Measured separately against the retained Microsoft PDF: regular/italic
  // advances match; the macOS bold cut does not. Word hhea is1903/-472 at2048UPM.
  faces: [
    { style: 'normal', source: 'local("GillSans")', ascent: '92.919921875%', descent: '23.046875%', unicodeRange: 'U+20-024F,U+2000-206F' },
    { style: 'italic', source: 'local("GillSans-Italic")', ascent: '92.919921875%', descent: '23.046875%', unicodeRange: 'U+20-024F,U+2000-206F' },
  ],
}];
const registeredRegularFamilies = new Map<string, FontFace[]>();

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

export async function initializeRegularFontFallback(profile: RegularFontFallbackFamily, changed?: () => void): Promise<'original'|'fallback'|'unavailable'> {
  const key = profile.family.toLowerCase();
  const previous = registeredRegularFamilies.get(key);
  if (previous?.every(face => document.fonts.has(face))) return 'fallback';
  if (previous) for (const face of previous) { document.fonts.delete(face); metricFaces.delete(face); }
  if ([...document.fonts].some(face => face.family.replace(/["']/g, '').toLowerCase() === key)) return 'original';
  const originals = await Promise.all(profile.nativeNames.map(async name => {
    try { await new FontFace('Studio native face check', `local(${JSON.stringify(name)})`).load(); return true; }
    catch { return false; }
  }));
  if (originals.some(Boolean)) return 'original';
  // Both measured styles must be usable before anything shadows the family.
  if (profile.faces.length !== 2 || !['normal','italic'].every(style => profile.faces.some(face => face.style === style))) return 'unavailable';
  const candidates = await Promise.all(profile.faces.map(async data => {
    try {
      const face = new FontFace(profile.family, data.source, { weight: '400', style: data.style,
        ascentOverride: data.ascent, descentOverride: data.descent, lineGapOverride: '0%', unicodeRange: data.unicodeRange });
      await face.load(); return { data, face };
    } catch { return undefined; }
  }));
  if (candidates.some(face => !face)) return 'unavailable';
  const faces = candidates.filter((face): face is NonNullable<typeof face> => !!face);
  const guards: FontFace[] = [];
  try {
    // CSS matches weight first, then its character map. A selected face with
    // no character proceeds to the next family, not another weight. The
    // NUL-only guards keep400 from becoming an unmeasured synthetic700 face.
    for (const { data } of faces) for (const weight of ['1 399','401 1000'])
      guards.push(new FontFace(profile.family, data.source, { weight, style: data.style, unicodeRange: 'U+0' }));
  } catch { return 'unavailable'; }
  const registered = [...faces.map(item => item.face), ...guards];
  for (const face of registered) document.fonts.add(face);
  for (const { face, data } of faces) metricFaces.set(face, { family: key, weight: '400', style: data.style,
    ratio: (parseFloat(data.ascent) + parseFloat(data.descent)) / 100 });
  registeredRegularFamilies.set(key, registered);
  if (changed) changed(); else document.dispatchEvent(new Event(wordFontMetricsChanged));
  return 'fallback';
}

// CSS family lookup can resolve an unavailable Windows name through a
// platform fontconfig alias. A local-only face probe admits actual installed
// Korean serif faces; Linux otherwise reaches the bundled Noto face directly.
async function initializeNativeKoreanSerif(changed: () => void) {
  try {
    const face = new FontFace('Studio Korean Serif', 'local("Batang"),local("바탕"),local("AppleMyungjo")');
    await face.load(); document.fonts.add(face); changed();
  } catch { /* No native serif: the CSS chain uses bundled Noto Serif KR. */ }
}

export async function initializeWordFontFallbacks(profiles: readonly FontFallbackFamily[] = families) {
  let released = false;
  const changed = () => { if (released) document.dispatchEvent(new Event(wordFontMetricsChanged)); };
  const initialized = Promise.all([...profiles.flatMap(profile => profile.faces.map(async face => {
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
  })), ...(profiles === families ? [initializeNativeKoreanSerif(changed), ...regularFamilies.map(profile => initializeRegularFontFallback(profile, changed))] : [])]);
  // A stalled optional local-font probe must not leave the startup window
  // blank. Late results still register and notify the measurement owners.
  let timer: ReturnType<typeof setTimeout> | undefined;
  try { await Promise.race([initialized, new Promise<void>(resolve => { timer = setTimeout(resolve, 1500); })]); }
  finally { clearTimeout(timer); released = true; }
}
