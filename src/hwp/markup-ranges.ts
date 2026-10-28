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
