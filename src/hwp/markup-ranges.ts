type Range = { start: number; end: number; open: string; inner: string };
/** Match complete elements, including paragraphs containing nested tables.
 * Inputs are sanitized HTML or the engine's bounded, normalized HWPX. */
export function elementRanges(source: string, name: string): Range[] {
  const result: Range[] = [], stack: { start: number; inner: number; open: string }[] = [];
  for (const m of source.matchAll(new RegExp(`<!--[\\s\\S]*?-->|<!\\[CDATA\\[[\\s\\S]*?\\]\\]>|<\\/?${name}\\b[^>]*>`, 'g'))) {
    const tag = m[0]; if (tag.startsWith('<!')) continue;
    if (tag.startsWith('</')) {
      const start = stack.pop();
      if (start && !stack.length) result.push({ start: start.start, end: m.index! + tag.length, open: start.open, inner: source.slice(start.inner, m.index) });
    } else if (!tag.endsWith('/>')) stack.push({ start: m.index!, inner: m.index! + tag.length, open: tag });
  }
  return result;
}

/** The binary writer can apply shape-level and paragraph-level page breaks
 * differently. Change only the flag on the already identified paragraph. */
export function directHwpPageBreak(source: string, owner: Range) {
  if (!owner.open.includes(' pageBreak="0"')) return;
  return source.slice(0,owner.start) + owner.open.replace(' pageBreak="0"',' pageBreak="1"') +
    source.slice(owner.start+owner.open.length);
}

/** Return complete sibling blocks from the sanitized native paste HTML.
 * Reject fragments whose text or inline elements could share a paragraph. */
export function topLevelHwpPasteBlocks(source: string): string[] | undefined {
  const block = /^(?:p|h[1-6]|div|table|blockquote|ol|ul|hr)$/;
  const voidElement = /^(?:area|base|br|col|embed|hr|img|input|link|meta|param|source|track|wbr)$/;
  const stack: string[] = [], parts: string[] = [];
  let at = 0, from = 0;
  for (const token of source.matchAll(/<!--[\s\S]*?-->|<(\/?)([a-z][a-z\d-]*)\b(?:[^<>"']|"[^"]*"|'[^']*')*>/gi)) {
    if (!stack.length && source.slice(at, token.index).trim()) return undefined;
    at = token.index! + token[0].length;
    if (!token[2]) continue;
    const tag = token[2].toLowerCase();
    if (token[1]) {
      if (tag !== stack.pop()) return undefined;
    } else {
      if (!stack.length && !block.test(tag)) return undefined;
      if (!voidElement.test(tag)) stack.push(tag);
    }
    if (!stack.length) {
      parts.push(source.slice(from, at));
      from = at;
    }
  }
  if (stack.length || source.slice(at).trim() || !parts.length) return undefined;
  parts[parts.length - 1] += source.slice(from);
  return parts;
}
