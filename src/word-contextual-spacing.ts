import { Extension } from '@tiptap/core';
import { Plugin } from '@tiptap/pm/state';
import { Decoration, DecorationSet } from '@tiptap/pm/view';
import type { Node as PMNode } from '@tiptap/pm/model';

type Paragraph<T> = { target: T; style: string | null; contextual: boolean | null; before: number; after: number };
type Gap = { before?: number; after?: number };
const lists = new Set(['bulletList', 'orderedList', 'listItem']);

/** OOXML contextual spacing subtracts the flagged side from the collapsed
 * gap (not from each CSS margin independently). 10pt after + 12pt before,
 * with contextual spacing only on the first paragraph, leaves 2pt. Lists are
 * transparent; table cells, stories, sections and non-paragraph blocks isolate
 * their flows so unrelated paragraphs never suppress each other's spacing. */
function spacingChanges<T>(root: T, kind: (node: T) => string,
  children: (node: T) => T[], paragraph: (node: T) => Paragraph<T> | null) {
  const changes = new Map<T, Gap>();
  let previous: Paragraph<T> | null = null;
  const walk = (node: T) => {
    for (const child of children(node)) {
      const current = paragraph(child);
      if (current) {
        if (previous?.style && current.style === previous.style && (previous.contextual === true || current.contextual === true)) {
          const gap = Math.max(0, Math.max(previous.after, current.before)
            - (previous.contextual === true ? previous.after : 0) - (current.contextual === true ? current.before : 0));
          // Keep the remainder on the unsuppressed side. This also makes
          // static layout idempotent when a print fragment is measured again.
          changes.set(previous.target, { ...changes.get(previous.target), after: previous.contextual === true ? 0 : gap });
          changes.set(current.target, { ...changes.get(current.target), before: current.contextual === true ? 0 : gap });
        }
        previous = current;
      } else if (lists.has(kind(child))) walk(child);
      else { previous = null; walk(child); previous = null; }
    }
  };
  walk(root);
  return changes;
}
const pixels = (value: unknown) => {
  if (typeof value !== 'string' || !/^\d+(?:\.\d+)?(?:px|pt)$/.test(value)) return 0;
  const n = parseFloat(value) * (value.endsWith('pt') ? 4 / 3 : 1);
  return Number.isFinite(n) ? n : 0;
};
const gapStyle = (gap: Gap) => [gap.before !== undefined ? `margin-top:${gap.before}px` : '',
  gap.after !== undefined ? `margin-bottom:${gap.after}px` : ''].filter(Boolean).join(';');

/** Static print DOM only. The editor uses decorations, retaining the original
 * before/after values and policy in its model and its exported HTML. */
export function layoutWordContextualSpacing(root: HTMLElement) {
  const changes = spacingChanges(root,
    node => ({ UL: 'bulletList', OL: 'orderedList', LI: 'listItem' }[node.tagName] || node.tagName),
    node => Array.from(node.children).filter((el): el is HTMLElement => el instanceof HTMLElement),
    node => /^(P|H[1-6])$/.test(node.tagName) ? { target: node,
      style: node.dataset.wordParagraphStyle || null,
      contextual: node.dataset.wordContextualSpacing === 'true' ? true : node.dataset.wordContextualSpacing === 'false' ? false : null,
      before: pixels(node.style.marginTop), after: pixels(node.style.marginBottom) } : null);
  for (const [paragraph, gap] of changes) {
    if (gap.before !== undefined) paragraph.style.marginTop = `${gap.before}px`;
    if (gap.after !== undefined) paragraph.style.marginBottom = `${gap.after}px`;
  }
}
function decorations(doc: PMNode) {
  let active = false;
  doc.descendants(node => {
    if (active) return false;
    if (['paragraph', 'heading'].includes(node.type.name)) { active = node.attrs.wordContextualSpacing === true; return false; }
  });
  if (!active) return DecorationSet.empty;
  type Entry = { node: PMNode; pos: number };
  const changes = spacingChanges<Entry>({ node: doc, pos: -1 }, entry => entry.node.type.name,
    entry => { const children: Entry[] = []; entry.node.forEach((node, offset) => children.push({ node, pos: entry.pos + 1 + offset })); return children; },
    entry => ['paragraph', 'heading'].includes(entry.node.type.name) ? { target: entry,
      style: entry.node.attrs.wordParagraphStyle || null, contextual: entry.node.attrs.wordContextualSpacing,
      before: pixels(entry.node.attrs.spaceBefore), after: pixels(entry.node.attrs.spaceAfter) } : null);
  return DecorationSet.create(doc, [...changes].map(([entry, gap]) =>
    Decoration.node(entry.pos, entry.pos + entry.node.nodeSize, { style: gapStyle(gap) })));
}
export const WordContextualSpacing = Extension.create({
  name: 'wordContextualSpacing',
  addProseMirrorPlugins: () => [new Plugin<DecorationSet>({
    state: { init: (_, state) => decorations(state.doc), apply: (tr, old) => tr.docChanged ? decorations(tr.doc) : old },
    props: { decorations(state) { return this.getState(state); } },
    appendTransaction(transactions, oldState, state) {
      if (!transactions.some(tr => tr.docChanged)) return null;
      // Opening/replacing a whole document has its own style namespace. It
      // must not inherit the former document's paragraph-kind expectations.
      if (transactions.some(tr => tr.steps.some((step, index) => {
        const value = step.toJSON();
        return value.stepType === 'replace' && value.from === 0 && value.to >= tr.docs[index].content.size;
      }))) return null;
      const kinds = new Map<string, Set<string>>();
      const kind = (node: PMNode) => node.type.name + (node.type.name === 'heading' ? node.attrs.level : '');
      oldState.doc.descendants(node => {
        if (!['paragraph', 'heading'].includes(node.type.name) || !node.attrs.wordParagraphStyle) return;
        const values = kinds.get(node.attrs.wordParagraphStyle) || new Set<string>();
        values.add(kind(node)); kinds.set(node.attrs.wordParagraphStyle, values);
      });
      if (!kinds.size) return null;
      const tr = state.tr;
      state.doc.descendants((node, pos) => {
        const source = kinds.get(node.attrs.wordParagraphStyle);
        if (!['paragraph', 'heading'].includes(node.type.name) || !source || source.has(kind(node))) return;
        // A heading changed to body text (including Enter at its end) is a
        // new paragraph style. Retire only the source identity/policy; keep
        // explicit formatting and let the existing undo transaction restore it.
        tr.setNodeMarkup(pos, undefined, { ...node.attrs, wordParagraphStyle: null, wordContextualSpacing: null });
      });
      return tr.docChanged ? tr : null;
    },
  })],
});
