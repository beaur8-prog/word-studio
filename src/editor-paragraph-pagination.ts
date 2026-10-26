import type { EditorView } from '@tiptap/pm/view';
import { Decoration } from '@tiptap/pm/view';
import { measureParagraphLines } from './editor-lines';

/** Continue ordinary body text at the same line boundaries used by print.
 * Widgets affect the view only; the paragraph and undo history stay intact. */
export function paginateParagraph(view: EditorView, paragraph: HTMLElement, g: {
  origin: number; scale: number; top: number; bottom: number; height: number; extra: number;
  removed: { top: number; height: number }[]; reserveAt?: (page: number) => number;
}) {
  if (paragraph.querySelector('img,table,[data-note-reference],[data-page-number]')) return null;
  const naturalY = (value: number) => {
    const y = (value - g.origin) / g.scale;
    return y - g.removed.filter(gap => gap.top < y - .5).reduce((sum, gap) => sum + gap.height, 0);
  };
  const lines = measureParagraphLines(view, paragraph, naturalY);
  if (lines.length < 2) return null;
  const stride = g.height + 24, decorations: Decoration[] = [], signatures: string[] = [];
  let extra = 0;
  for (const line of lines) {
    if (line.bottom - line.top > g.height - g.top - g.bottom) return null;
    const y = line.top + g.extra + extra, page = Math.max(0, Math.floor((y - g.top + .5) / stride));
    if (line.bottom + (line.after || 0) + g.extra + extra <= page * stride + g.height - g.bottom - (g.reserveAt?.(page) ?? 0) + .5) continue;
    const gap = Math.max(0, Math.ceil(((page + 1) * stride + g.top + (line.before || 0) - y) * 64) / 64);
    const pos = line.pos();
    decorations.push(Decoration.widget(pos, () => {
      const spacer = document.createElement('span'); spacer.className = 'pagination-cell-gap';
      spacer.contentEditable = 'false'; spacer.setAttribute('aria-hidden', 'true');
      spacer.style.cssText = `display:block;height:${gap}px;line-height:0;pointer-events:none`;
      return spacer;
    }, { side: -1, key: `paragraph-gap-${pos}-${gap}`, ignoreSelection: true }));
    signatures.push(`paragraph:${pos}:${gap}`); extra += gap;
  }
  return { decorations, signatures, extra };
}
