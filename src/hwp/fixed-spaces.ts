/** The reader exposes HWP fwSpace as U+2007, whose digit-width glyph is
 * wider than the native quarter-em control. Use the matching Unicode space
 * only at serialized fwSpace positions; literal figure spaces stay intact.
 * Offsets are UTF-16, as in the serialized metadata, not native scalar indices. */
export function restoreHwpFixedSpaces(html: string, source: string, positions: readonly number[]): string {
  if (!positions.length || /<(?:img|table|svg)\b|<!--\s*rhwp:/i.test(html)) return html;
  const selected = new Set(positions);
  if ([...selected].some(i => !Number.isInteger(i) || i < 0 || source[i] !== '\u2007')) return html;
  const decode = (text: string) => text.replace(/&(?:amp|lt|gt|quot|apos|nbsp|#\d+|#x[\da-f]+);/gi, entity => {
    const names: Record<string, string> = { '&amp;': '&', '&lt;': '<', '&gt;': '>', '&quot;': '"', '&apos;': "'", '&nbsp;': '\u00a0' };
    if (names[entity.toLowerCase()]) return names[entity.toLowerCase()];
    const hex = entity[2].toLowerCase() === 'x', n = Number.parseInt(entity.slice(hex ? 3 : 2, -1), hex ? 16 : 10);
    return n > 0 && n <= 0x10ffff ? String.fromCodePoint(n) : '\ufffd';
  });
  const tokens: { raw: string; text?: string }[] = [];
  let depth = 0;
  for (const match of html.matchAll(/<!--[\s\S]*?-->|<[^>]+>|&(?:#\d+|#x[\da-f]+|\w+);|[^<&]+|[<&]/gi)) {
    const raw = match[0];
    if (raw.startsWith('<')) {
      if (/^<span\b/i.test(raw)) depth++;
      if (/^<\/span\b/i.test(raw)) depth--;
      tokens.push({ raw, ...(/^<br\b/i.test(raw) ? { text: '\n' } : {}) });
    } else tokens.push({ raw, ...(!depth && /^[\t\r\n ]*[\r\n][\t\r\n ]*$/.test(raw) ? {} : { text: decode(raw) }) });
  }
  // Object labels or unsupported field positions can differ from native text.
  // Never move a space to a guessed offset or touch the surrounding content.
  if (tokens.map(token => token.text ?? '').join('') !== source) return html;
  const ordered = [...selected].sort((a, b) => a - b);
  let offset = 0, next = 0;
  return tokens.map(token => {
    if (token.text === undefined) return token.raw;
    const start = offset; offset += token.text.length;
    if (next === ordered.length || ordered[next] >= offset) return token.raw;
    while (next < ordered.length && ordered[next] < offset) next++;
    if (token.raw.startsWith('&')) return selected.has(start) ? '\u2005' : token.raw;
    let value = '', at = start;
    for (const char of token.raw) { value += selected.has(at) ? '\u2005' : char; at += char.length; }
    return value;
  }).join('');
}
