/** HWP character shapes have seven script-specific faces. The native HTML
 * bridge combines the Korean and Latin faces into one CSS fallback list,
 * which cannot express that distinction (and DOCX kept only its first face).
 * Restore the selected face as ordinary inline formatting, without changing
 * text, size, tracking, or line breaks. */
export type HwpFontRun = { start: number; end: number; fonts: readonly string[] };

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

export function restoreScriptFonts(html: string, source: string, runs: readonly HwpFontRun[]): string {
  if (!source || !runs.some(run => run.fonts.some(font => font && font !== run.fonts[0]))) return html;
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
  // Controls/field labels can have different logical positions from HTML.
  // Keep those fragments untouched rather than assigning a face by a guess.
  if (tokens.map(token => token.text || '').join('') !== source) return html;
  let offset = 0, index = 0, previousLanguage = 0;
  return tokens.map(token => {
    if (token.text === undefined) return token.raw;
    if (/^<br\b/i.test(token.raw)) { offset++; previousLanguage = 0; return token.raw; }
    let output = '', pending = '', selected = '';
    const flush = () => {
      if (!pending) return;
      // Quote each face so a document font name cannot inject a CSS property.
      const css = selected.split(',').map(face => JSON.stringify(face.trim().replace(/^(["'])(.*)\1$/s, '$2'))).join(',');
      output += selected ? `<span style="font-family:${escape(css)};">${escape(pending)}</span>` : escape(pending);
      pending = '';
    };
    for (const char of token.text) {
      const script = language(char) ?? previousLanguage;
      if (language(char) !== undefined) previousLanguage = script;
      while (index < runs.length && runs[index].end <= offset) index++;
      const run = runs[index], font = run && run.start <= offset ? run.fonts[script] || run.fonts[0] : '';
      const next = font && font !== run.fonts[0] && font.length <= 500 && !/[\r\n\0]/.test(font) ? font : '';
      if (next !== selected) { flush(); selected = next; }
      pending += char; offset += char.length;
    }
    flush(); return output;
  }).join('');
}
