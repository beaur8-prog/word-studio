import type { EditorView } from '@tiptap/pm/view';
import { Decoration } from '@tiptap/pm/view';
import { measureParagraphLines } from './editor-lines';
import { paragraphPrefixHeight, paragraphTextOffset } from './paragraph-fragments';

/** Continue ordinary body text at the same line boundaries used by print.
 * Widgets affect the view only; the paragraph and undo history stay intact. */
export function paginateParagraph(view: EditorView, paragraph: HTMLElement, g: {
  origin: number; scale: number; top: number; bottom: number; height: number; extra: number;
  removed: { top: number; height: number }[];
  references?: { id: string; top: number }[];
  reserveAt?: (page: number, references: { id: string; page: number }[]) => number;
}) {
  if (paragraph.querySelector('img:not(.ProseMirror-separator),table,[data-page-number]')) return null;
  const naturalY = (value: number) => {
    const y = (value - g.origin) / g.scale;
    return y - g.removed.filter(gap => gap.top < y - .5).reduce((sum, gap) => sum + gap.height, 0);
  };
  const lines = measureParagraphLines(view, paragraph, naturalY);
  if (lines.length < 2) return null;
  const stride = g.height + 24, decorations: Decoration[] = [], signatures: string[] = [];
  let extra = 0;
  const listParagraph = !!paragraph.closest('li');
  const references: { id: string; page: number }[] = [];
  for (let index = 0; index < lines.length; index++) {
    const line = lines[index];
    if (line.bottom - line.top > g.height - g.top - g.bottom) return null;
    const y = line.top + g.extra + extra;
    let page = Math.max(0, Math.floor((y - g.top + .5) / stride));
    const here = (g.references ?? []).filter(ref => ref.top >= line.top - .5 && ref.top < line.bottom - .5 && !references.some(placed => placed.id === ref.id));
    const reserve = () => g.reserveAt?.(page, [...references, ...here.map(ref => ({ id: ref.id, page }))]) ?? 0;
    const boundary = page * stride + g.height - g.bottom - reserve();
    let bottom = line.bottom + (line.after || 0) + g.extra + extra;
    // Uniform lines already have an exact measured grid. Avoid repeatedly
    // cloning a growing prefix on every page of a very long list paragraph.
    const irregular = listParagraph && [lines[index - 1], lines[index + 1]].some(other => other &&
      (Math.abs((other.bottom - other.top) - (line.bottom - line.top)) > .5 ||
        Math.abs((other.before || 0) - (line.before || 0)) > .5 || Math.abs((other.after || 0) - (line.after || 0)) > .5));
    if (listParagraph && irregular && index + 1 < lines.length && Math.abs(bottom - boundary) <= line.bottom - line.top + (line.before || 0) + (line.after || 0)) {
      const point = view.domAtPos(lines[index + 1].pos());
      const offset = paragraphTextOffset(paragraph, point.node, point.offset);
      bottom = naturalY(paragraph.getBoundingClientRect().top) + paragraphPrefixHeight(paragraph, offset, g.scale) + g.extra + extra;
    }
    if (bottom <= boundary + .5) {
      references.push(...here.map(ref => ({ id: ref.id, page }))); continue;
    }
    const gap = Math.max(0, Math.ceil(((page + 1) * stride + g.top + (line.before || 0) - y) * 64) / 64);
    page++;
    if (line.bottom - line.top + (line.before || 0) + (line.after || 0) + reserve() > g.height - g.top - g.bottom + .5) return null;
    references.push(...here.map(ref => ({ id: ref.id, page })));
    const pos = line.pos();
    decorations.push(Decoration.widget(pos, () => {
      const spacer = document.createElement('span'); spacer.className = 'pagination-cell-gap';
      spacer.contentEditable = 'false'; spacer.setAttribute('aria-hidden', 'true');
      spacer.style.cssText = `display:block;height:${gap}px;line-height:0;pointer-events:none`;
      return spacer;
    }, { side: -1, key: `paragraph-gap-${pos}-${gap}`, ignoreSelection: true }));
    signatures.push(`paragraph:${pos}:${gap}`); extra += gap;
  }
  return { decorations, signatures, extra, references };
}
