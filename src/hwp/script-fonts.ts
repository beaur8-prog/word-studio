import { hwpLatinFont, hwpLatinFontAttribute } from '../hwp-latin-font';
import { characterScaleProperty } from '../character-scale';
import { hwpSingleBodyFace, hwpSourceSpacesAttribute, hwpSourceWrapAttribute, hwpWrapFingerprint, normalizeHwpWrapText, parseHwpSourceWrap, type HwpSourceOrigin } from '../hwp-text-advances';

/** The native HTML bridge combines script-specific faces into a CSS fallback
 * list and emits the base size. Recover the source character shape's face
 * and dimensions for the script of each run. */
export type HwpScriptTypography = {
  fonts: readonly string[];
  base?: number;
  sizes?: readonly number[];
  tracking?: readonly (number | null)[];
  widths?: readonly number[];
};
export type HwpFontRun = HwpScriptTypography & { start: number; end: number; charShapeId?: number };

/** Query only each distinct character shape, including shapes after the first
 * run of a paragraph. Native offsets remain absolute until this boundary. */
export function sourceFontRuns(serialized: string, start: number, source: string, read: (offset: number) => Record<string, unknown>,
  cache = new Map<number, HwpScriptTypography>(), includeSizes = false): HwpFontRun[] {
  // Native character offsets count Unicode scalars, while HTML/JS offsets
  // count UTF-16 units. A supplementary character must not shift later faces.
  const offsets = [0];
  for (const char of source) offsets.push(offsets[offsets.length - 1] + char.length);
  return JSON.parse(serialized).map((run: { startOffset: number; endOffset: number; charShapeId: number }) => {
    let typography = cache.get(run.charShapeId);
    if (!typography) {
      const char = read(run.startOffset);
      const fonts = Array.isArray(char.fontFamilies) ? char.fontFamilies.map(font => typeof font === 'string' ? font : '') : [];
      const base = typeof char.fontSize === 'number' && Number.isSafeInteger(char.fontSize) ? char.fontSize : 0;
      const sizes = includeSizes && base > 0 && Array.isArray(char.relativeSizes) ? char.relativeSizes.slice(0, 7).map(percent =>
        Number.isInteger(percent) && percent >= 10 && percent <= 250 && percent !== 100 ? base * percent / 10000 : 0) : undefined;
      const spacings = Array.isArray(char.spacings) ? char.spacings : [];
      // The clipboard has one tracking value for the whole character
      // shape. A non-Hangul run can specify a different percentage.
      const tracking = includeSizes && base > 0 ? Array.from({ length: 7 }, (_, script) => {
        const percent = spacings[script];
        if (!Number.isInteger(percent) || Math.abs(percent) > 50 ||
          percent === spacings[0] && (!percent || !sizes?.[script])) return null;
        return percent;
      }) : undefined;
      const widths = Array.isArray(char.ratios) ? char.ratios.slice(0, 7).map(percent =>
        Number.isInteger(percent) && percent >= 50 && percent <= 200 && percent !== 100 ? percent : 0) : undefined;
      typography = { fonts, base, sizes, tracking, widths };
      cache.set(run.charShapeId, typography);
    }
    return { ...typography, charShapeId: run.charShapeId, start: offsets[run.startOffset - start] ?? source.length, end: offsets[run.endOffset - start] ?? source.length };
  });
}

/** Mark ordinary spaces only when every character shape in this complete
 * HWP paragraph explicitly uses the native space advance. */
export function preserveHwpSourceSpaces(html: string, text: string, runs: readonly HwpFontRun[],
  fontSpace?: readonly boolean[], readLines?: () => string, origin?: HwpSourceOrigin, recordBodyWidth = false,
  allowSingleBodyFallback = false): string {
  if (!fontSpace || !text.includes(' ') || !/[가-힣]/u.test(text) || !runs.length ||
      /[\t\r\n]/.test(text) || /<(?:img|table|br|svg)\b|<!--\s*rhwp:/i.test(html)) return html;
  let end = 0;
  for (const run of runs) {
    if (run.start !== end || run.end <= end || run.charShapeId === undefined || fontSpace[run.charShapeId] !== false) return html;
    end = run.end;
  }
  if (end !== text.length || !/^<p\b/.test(html)) return html;
  let reference = '';
  if (readLines && text.length <= 20000 && !/[\uD800-\uDFFF\ufffc]/u.test(text)) {
    try {
      const stored = JSON.parse(readLines());
      const lines = stored.lines as { textStart: number; width?: number }[] | undefined;
      const positions = lines?.map(line => line.textStart);
      const uniformBody = allowSingleBodyFallback && origin && recordBodyWidth && positions?.length === 1 &&
        text.length <= 160 && !/<(?:u|s|b|strong|em|i|a|sup|sub)\b/i.test(html) &&
        runs.every(run => run.charShapeId === runs[0].charShapeId && run.fonts.every(font => font === hwpSingleBodyFace) &&
          !run.sizes?.some(Boolean) && !run.widths?.some(Boolean) && !run.tracking?.some(Boolean));
      if (stored.stored && stored.controlCount === 0 && positions && (uniformBody || positions.length > 1) &&
          positions.length <= 80 && positions[0] === 0 &&
          positions.every((at, index) => Number.isInteger(at) && at < text.length && (!index || at > positions[index - 1]))) {
        const ends = [...positions.slice(1).map(at => normalizeHwpWrapText(text.slice(0, at)).length), normalizeHwpWrapText(text).length];
        if (ends.every((value, index) => value > (index ? ends[index - 1] : 0)) && (!uniformBody || ends[0] >= 12)) {
          // The stored line segment width is the body measure before a
          // hanging indent. Retain it only if every original line agrees.
          const width = lines?.[0]?.width;
          const bodyWidth = recordBodyWidth && width !== undefined && width > 0 && width <= 10000 &&
            lines?.every(line => line.width === width) ? width : undefined;
          if (!uniformBody || bodyWidth)
            reference = ` ${hwpSourceWrapAttribute}="${escape(JSON.stringify({ fingerprint: hwpWrapFingerprint(text), ends,
              ...(origin && { origin }), ...(bodyWidth && { bodyWidth }), ...(uniformBody && { singleBodyFallback: true }) }))}"`;
        }
      }
    } catch { /* A natural fallback remains available without stored layout. */ }
  }
  return html.replace(/^<p\b/, `<p ${hwpSourceSpacesAttribute}="true"${reference}`);
}

function language(char: string): number | undefined {
  const cp = char.codePointAt(0)!;
  if (cp === 0x318d) return 6;
  if (/[\u1100-\u11ff\u3130-\u318f\ua960-\ua97f\uac00-\ud7ff]/u.test(char)) return 0;
  if (/[A-Za-z0-9\u00c0-\u024f\u1e00-\u1eff]/u.test(char)) return 1;
  if (/[\u3400-\u9fff\uf900-\ufaff\u{20000}-\u{2a6df}]/u.test(char)) return 2;
  if (/[\u3040-\u30ff\u31f0-\u31ff]/u.test(char)) return 3;
  if (/[\u2190-\u27bf\u3000-\u303f]/u.test(char)) return 5;
  // Spaces and ordinary punctuation follow the preceding script. Other
  // scripts use HWP's OTHER face; do not treat their letters as punctuation.
  if (/\p{L}/u.test(char)) return 4;
  return undefined;
}

const escape = (text: string) => text.replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]!);
const decode = (text: string) => text.replace(/&(?:amp|lt|gt|quot|apos|nbsp|#\d+|#x[\da-f]+);/gi, value => {
  const named: Record<string, string> = { '&amp;': '&', '&lt;': '<', '&gt;': '>', '&quot;': '"', '&apos;': "'", '&nbsp;': '\u00a0' };
  if (named[value.toLowerCase()]) return named[value.toLowerCase()];
  const n = value[2].toLowerCase() === 'x' ? parseInt(value.slice(3, -1), 16) : Number(value.slice(2, -1));
  return n > 0 && n <= 0x10ffff ? String.fromCodePoint(n) : '\ufffd';
});

function textTokens(html: string) {
  const tokens: { raw: string; text?: string }[] = [];
  let spanDepth = 0;
  for (const match of html.matchAll(/<!--[\s\S]*?-->|<[^>]+>|[^<]+|</g)) {
    const raw = match[0];
    if (raw[0] === '<') {
      if (/^<span\b/i.test(raw)) spanDepth++;
      if (/^<\/span\b/i.test(raw)) spanDepth--;
      tokens.push({ raw, ...(/^<br\b/i.test(raw) ? { text: '\n' } : {}) });
    } else tokens.push({ raw, ...(!spanDepth && /^[\t\r\n ]*[\r\n][\t\r\n ]*$/.test(raw) ? {} : { text: decode(raw) }) });
  }
  return tokens;
}

type BodyLineWitness = { origin: HwpSourceOrigin; hanging: number; painted: readonly { y: number; charStart?: number }[] };
type ColumnPaintLine = { y: number; h: number; page: number; charStart?: number };

/** Record a complete authored paragraph from one physical text area. The
 * conversion already represents certain source space codes in its own form. */
function preserveHwpMeasuredLineEnds(html: string, text: string, readLines: () => string,
  body?: BodyLineWitness, columnPaint?: readonly ColumnPaintLine[]) {
  if (text.length > 20000 || /[\u0000-\u001f\uD800-\uDFFF\ufffc]/u.test(text) ||
      !/^<p\b[^>]*>[\s\S]*<\/p>$/i.test(html) || /<(?!\/?(?:p|span|strong|b|i|em|u)\b)/i.test(html)) return html;
  const literal = textTokens(html).map(token => token.text ?? '').join('');
  const normalized = normalizeHwpWrapText(literal);
  if (!normalized || normalized !== normalizeHwpWrapText(text)) return html;
  try {
    const { stored, controlCount, lines } = JSON.parse(readLines()) as {
      stored: boolean; controlCount: number;
      lines?: { textStart: number; verticalPosition: number; width: number; height: number; textHeight: number; spacing: number }[];
    };
    if (!stored || controlCount || !lines || lines.length < 2 || lines.length > 80 || lines[0].textStart !== 0) return html;
    const width = lines[0].width;
    if (!(width > 0 && width <= 10000) ||
        lines.some((line, index) => line.width !== width || !Number.isInteger(line.textStart) ||
          line.textStart >= text.length || (index > 0 && (line.textStart <= lines[index - 1].textStart ||
            line.verticalPosition <= lines[index - 1].verticalPosition)))) return html;
    const ends = [...lines.slice(1).map(line => normalizeHwpWrapText(text.slice(0, line.textStart)).length), normalized.length];
    if (ends.some((end, index) => end <= (index ? ends[index - 1] : 0))) return html;
    const fingerprint = hwpWrapFingerprint(literal);
    const place = new RegExp(` ${hwpSourceWrapAttribute}="([^"]*)"`);
    const open = html.slice(0, html.indexOf('>') + 1);
    const previous = open.match(place);
    let geometry;
    if (body) {
      const size = lines[0].height, pitch = size + lines[0].spacing;
      if (body.painted.length !== lines.length || !(size > 0 && pitch > size) ||
          lines.some((line, i) => line.textStart !== body.painted[i].charStart ||
            Math.abs(line.height - size) > .04 || Math.abs(line.textHeight - size) > .04 ||
            Math.abs(line.height + line.spacing - pitch) > .04 ||
            i > 0 && Math.abs(body.painted[i].y - body.painted[i - 1].y - pitch) > .2)) return html;
      const earlier = previous && parseHwpSourceWrap(decode(previous[1]));
      // The new proof carries its own physical position. It must leave any
      // already supported HWP source measurement available as it was.
      if (previous && (!earlier || earlier.cellWidth || earlier.columnWidth !== undefined ||
          earlier.fingerprint !== fingerprint || earlier.ends.length !== ends.length ||
          earlier.ends.some((end, index) => end !== ends[index]))) return html;
      geometry = { ...earlier, authoredBodyLines: { origin: body.origin, width, hanging: body.hanging, size, pitch } };
    } else {
      geometry = { columnWidth: width };
      const [first, last] = lines;
      if (columnPaint?.length === 2 && lines.length === 2 && first.height > last.height &&
          columnPaint[0].page === columnPaint[1].page && lines.every((line, index) =>
            line.textStart === columnPaint[index].charStart &&
            Math.abs(line.height - line.textHeight) <= .04 &&
            Math.abs(line.height - columnPaint[index].h) <= .16 && line.spacing >= 0)) {
        const firstPitch = first.height + first.spacing, secondPitch = last.height + last.spacing;
        if (firstPitch > secondPitch && secondPitch >= 4 && firstPitch <= 480 &&
            Math.abs(last.verticalPosition - first.verticalPosition - firstPitch) <= .04 &&
            Math.abs(columnPaint[1].y - columnPaint[0].y - firstPitch) <= .16)
          geometry = { ...geometry, columnPitches: [firstPitch, secondPitch] };
      }
    }
    const proof = escape(JSON.stringify({ fingerprint, ends, ...geometry }));
    if (previous) return html.replace(open, open.replace(place, ` ${hwpSourceWrapAttribute}="${proof}"`));
    return html.replace(/^<p\b/, `<p ${hwpSourceWrapAttribute}="${proof}"`);
  } catch {
    return html;
  }
}

export function preserveHwpColumnLineEnds(html: string, text: string, readLines: () => string, painted?: readonly ColumnPaintLine[]) {
  return preserveHwpMeasuredLineEnds(html, text, readLines, undefined, painted);
}

export function preserveHwpBodyLineEnds(html: string, text: string, readLines: () => string, witness: BodyLineWitness) {
  return preserveHwpMeasuredLineEnds(html, text, readLines, witness);
}

/** Recover the Latin metrics of a literal single-shape story whose bridge
 * span still carries the base font size and inter-character spacing. */
export function restoreUniformLatinMetrics(html: string, text: string, run: HwpFontRun): { html: string; size: number } | undefined {
  if (run.start !== 0 || run.end !== text.length || !(run.base && run.base > 0) ||
      run.widths?.[1] || !run.fonts[1] || !/[A-Za-z]/.test(text) ||
      !/^[ -~\u00a0\u2007\u2010-\u201f\u2026]+$/.test(text) || /<(?!\/?(?:p|span)\b)/i.test(html)) return;
  const literal = textTokens(html).map(token => token.text ?? '').join('');
  // A serialized fixed space is already projected to its quarter-em
  // counterpart. Its one-character position must stay exactly where it was.
  if (literal.length !== text.length || [...text].some((char, i) =>
      char !== literal[i] && !(char === '\u2007' && literal[i] === '\u2005'))) return;
  const percent = run.tracking?.[1];
  if (percent === undefined || percent === null) return;
  const spans = [...html.matchAll(/(<span\b[^>]*\bstyle=")([^"]*)"/gi)];
  const original = spans.filter(span => /\bletter-spacing:[-+\d.]+px[; ]/.test(span[2]));
  if (original.length !== 1 || spans.some(span => span !== original[0] && /\b(?:font-size|letter-spacing):/.test(span[2]))) return;
  const [span] = original;
  const base = run.base / 100;
  const oldSize = /\bfont-size:([\d.]+)pt\b/.exec(span[2]);
  if (!oldSize || Number(oldSize[1]) !== base) return;
  const size = Math.round((run.sizes?.[1] || base) * 2) / 2;
  if (!(size >= .5 && size <= 120)) return;
  const twips = Math.round(size * percent / 5);
  const style = span[2].replace(/\bfont-size:[\d.]+pt\b/, 'font-size:' + size + 'pt')
    .replace(/\bletter-spacing:[-+\d.]+px\b/, 'letter-spacing:' + twips / 20 + 'pt');
  if (style !== span[2]) return {
    html: html.slice(0, span.index) + span[1] + style + '"' + html.slice(span.index + span[0].length),
    size,
  };
}

export function restoreScriptFonts(html: string, source: string, runs: readonly HwpFontRun[], inColumns = false,
  latinFaces?: readonly (string | undefined)[]): string {
  const nativeLatin = (run: HwpFontRun) => latinFaces?.[run.charShapeId ?? -1] || run.fonts[1];
  const differentFaces = runs.some(run => run.fonts.some(font => font && font !== run.fonts[0]) ||
    !!nativeLatin(run) && nativeLatin(run) !== run.fonts[0]);
  const columnMetrics = inColumns && runs.every(run => run.base === runs[0].base);
  const canRestoreMetrics = (!inColumns || columnMetrics) && !/<(?:sup|sub)\b/i.test(html);
  const differentSizes = canRestoreMetrics && runs.some(run => run.sizes?.some(Boolean));
  const differentTracking = canRestoreMetrics && runs.some(run => run.tracking?.some(rate => rate !== null));
  const differentWidths = runs.some(run => run.widths?.some(Boolean));
  if (!source || !(differentFaces || differentSizes || differentTracking || differentWidths)) return html;
  // Such fragments contain non-text positions and possibly very large image
  // payloads. Text-only ranges on either side are restored separately.
  if (/<(?:img|table)\b|<!--\s*rhwp:/i.test(html)) return html;
  const tokens = textTokens(html);
  // Controls/field labels can have different logical positions from HTML.
  // Keep those fragments untouched rather than assigning a face by a guess.
  if (tokens.map(token => token.text || '').join('') !== source) return html;
  let offset = 0, index = 0, previousLanguage = 0;
  return tokens.map(token => {
    if (token.text === undefined) return token.raw;
    if (/^<br\b/i.test(token.raw)) { offset++; previousLanguage = 0; return token.raw; }
    let output = '', pending = '', selected = '', selectedSize = 0, selectedWidth = 0, selectedTracking: number | null = null;
    let selectedLatin: string | undefined;
    const flush = () => {
      if (!pending) return;
      // Quote each face so a document font name cannot inject a CSS property.
      const face = selected ? selected.split(',').map(font => JSON.stringify(font.trim().replace(/^(["'])(.*)\1$/s, '$2'))).join(',') : '';
      const css = (face ? `font-family:${escape(face)};` : '') +
        (selectedTracking === null ? '' : `letter-spacing:${selectedTracking}pt;`) +
        (selectedWidth ? `${characterScaleProperty}:${selectedWidth};` : '');
      const style = css ? ` style="${css}"` : '';
      const size = selectedSize ? ` data-hwp-script-size="${selectedSize}pt"` : '';
      const latin = selectedLatin ? ` ${hwpLatinFontAttribute}="${escape(selectedLatin)}"` : '';
      output += style || size || latin ? `<span${style}${size}${latin}>${escape(pending)}</span>` : escape(pending);
      pending = '';
    };
    for (const char of token.text) {
      const script = language(char) ?? previousLanguage;
      previousLanguage = script;
      while (index < runs.length && runs[index].end <= offset) index++;
      const run = runs[index]?.start <= offset ? runs[index] : undefined;
      const font = differentFaces && run ? run.fonts[script] || run.fonts[0] : '';
      // Select the known base face too. The native bridge combines Korean and
      // Latin into a CSS fallback list; inheriting that list makes a missing
      // Korean face fall back to the Latin face before a DOCX save, while the
      // saved single face uses the application's Korean fallback. These faces
      // come from the source character shape, not an authored CSS fallback list.
      const next = font && font.length <= 500 && !/[\r\n\0]/.test(font) ? font : '';
      const sourceSize = (differentSizes && run?.sizes?.[script]) || 0;
      // Use the same half-point and twip values when importing, editing and
      // saving an ordinary column. Otherwise a small size quantization at
      // the first save can create or remove a line at a forced boundary.
      const size = columnMetrics && sourceSize ? Math.max(.5, Math.round(sourceSize * 2) / 2) : sourceSize;
      const rate = differentTracking ? run?.tracking?.[script] ?? null : null;
      const spacingSize = size || (run?.base ?? 0) / 100;
      const tracking = rate === null ? null : Math.round(spacingSize * rate / 5) / 20;
      const width = (differentWidths && run?.widths?.[script]) || 0;
      const latin = script === 0 && next && run ? hwpLatinFont(run.fonts[0], nativeLatin(run)) : undefined;
      if (next !== selected || size !== selectedSize || tracking !== selectedTracking || width !== selectedWidth || latin !== selectedLatin) {
        flush(); selected = next; selectedSize = size; selectedTracking = tracking; selectedWidth = width; selectedLatin = latin;
      }
      pending += char; offset += char.length;
    }
    flush(); return output;
  }).join('');
}
