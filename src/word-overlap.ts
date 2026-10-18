import { Node } from '@tiptap/core';

export function overlapParts(value: unknown): string[] {
  try {
    const parts = typeof value === 'string' && value.length < 1024 ? JSON.parse(value) : null;
    return Array.isArray(parts) && parts.length >= 2 && parts.length <= 8 &&
      parts.every(p => typeof p === 'string' && p.length > 0 && p.length <= 64 && !/[(),\\\r\n]/.test(p)) ? parts : [];
  } catch { return []; }
}
export function overlapInstruction(instruction: string) {
  const match = instruction.match(/^\s*EQ\s+\\o\s*\\ac\(([^()]*)\)\s*$/i);
  return match ? overlapParts(JSON.stringify(match[1].split(','))) : [];
}

/** Legacy Word EQ overstrike fields are text, not external field commands. */
export const WordOverlap = Node.create({
  name: 'wordOverlap', group: 'inline', inline: true, atom: true,
  addAttributes: () => ({
    parts: { default: '[]', parseHTML: el => JSON.stringify(overlapParts(el.getAttribute('data-word-overlap'))),
      renderHTML: a => ({ 'data-word-overlap': JSON.stringify(overlapParts(a.parts)) }) },
  }),
  parseHTML: () => [{ tag: 'span[data-word-overlap]', getAttrs: el => overlapParts(el.getAttribute('data-word-overlap')).length ? null : false }],
  renderHTML: ({ node, HTMLAttributes }) => ['span', { ...HTMLAttributes, class: 'word-overlap' },
    ...overlapParts(node.attrs.parts).map(part => ['span', {}, part])],
  renderText: ({ node }) => overlapParts(node.attrs.parts).join(''),
});
