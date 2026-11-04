import { Extension } from '@tiptap/core';
import { Plugin } from '@tiptap/pm/state';
import { Decoration, DecorationSet } from '@tiptap/pm/view';
import type { Node as PMNode } from '@tiptap/pm/model';

import { wordNumberText } from './word-numbering';

type Attributes = { id?: unknown; level?: unknown; format?: unknown; pattern?: unknown; start?: unknown };

/** Imported labels are snapshots, not counters: Enter copies the old LI's
 * attributes. Derive current display labels without adding edits to history.
 * numId continues across separate list shells; each story has its own flow. */
function listLabels<T>(root: T, kind: (node: T) => string, children: (node: T) => T[], attrs: (node: T) => Attributes) {
  const result = new Map<T, string>(), formats = new Map<string, Map<number, string>>();
  const collect = (node: T) => {
    const a = attrs(node);
    if (kind(node) === 'listItem' && a.id) {
      const levels = formats.get(String(a.id)) || new Map<number, string>();
      levels.set(Number(a.level) || 0, String(a.format || 'decimal')); formats.set(String(a.id), levels);
    }
    children(node).forEach(collect);
  };
  collect(root);
  type Counter = { values: number[]; starts: number[] };
  const walk = (node: T, counters: Map<string, Counter>, parent?: T) => {
    if (kind(node) === 'story') counters = new Map();
    const a = attrs(node);
    if (kind(node) === 'listItem' && a.id && typeof a.pattern === 'string') {
      const id = String(a.id), level = Math.max(0, Math.min(8, Number(a.level) || 0));
      const counter = counters.get(id) || { values: [], starts: [] };
      counter.starts[level] ??= Math.max(1, Number(parent && attrs(parent).start) || 1);
      counter.values[level] = counter.values[level] === undefined ? counter.starts[level] : counter.values[level] + 1;
      counter.values.length = level + 1; counters.set(id, counter);
      result.set(node, a.pattern.replace(/%([1-9])/g, (_, n) => {
        const index = Number(n) - 1;
        return wordNumberText(counter.values[index] ?? counter.starts[index] ?? 1, formats.get(id)?.get(index) || String(a.format || 'decimal'));
      }));
    }
    children(node).forEach(child => walk(child, counters, node));
  };
  walk(root, new Map()); return result;
}
export function layoutWordListLabels(root: HTMLElement) {
  if (!root.querySelector('li[data-list-id]')) return;
  const labels = listLabels(root,
    node => node.matches('[data-document-textbox],[data-document-note]') ? 'story' : node.tagName === 'LI' ? 'listItem' : node.tagName,
    node => [...node.children].filter((c): c is HTMLElement => c instanceof HTMLElement),
    node => ({ id: node.dataset.listId, level: node.dataset.listLevel, format: node.dataset.listFormat, pattern: node.dataset.listPattern, start: node.getAttribute('start') }));
  for (const [item, label] of labels) item.dataset.listLabel = label;
}
function decorations(doc: PMNode) {
  let active = false;
  doc.descendants(node => {
    if (node.type.name === 'listItem' && node.attrs.wordListId) active = true;
    return !active && !node.isTextblock;
  });
  if (!active) return DecorationSet.empty;
  type Entry = { node: PMNode; pos: number };
  const labels = listLabels<Entry>({ node: doc, pos: -1 },
    entry => ['documentTextBox', 'documentNote'].includes(entry.node.type.name) ? 'story' : entry.node.type.name,
    entry => { const nodes: Entry[] = []; entry.node.forEach((node, offset) => nodes.push({ node, pos: entry.pos + 1 + offset })); return nodes; },
    entry => ({ id: entry.node.attrs.wordListId, level: entry.node.attrs.wordListLevel, format: entry.node.attrs.wordListFormat, pattern: entry.node.attrs.wordListPattern, start: entry.node.attrs.start }));
  return DecorationSet.create(doc, [...labels].filter(([entry, label]) => label !== entry.node.attrs.wordListLabel)
    .map(([entry, label]) => Decoration.node(entry.pos, entry.pos + entry.node.nodeSize, { 'data-list-label': label })));
}
export const WordListLabels = Extension.create({
  name: 'wordListLabels',
  addProseMirrorPlugins: () => [new Plugin<DecorationSet>({
    state: { init: (_, state) => decorations(state.doc), apply: (tr, old) => tr.docChanged ? decorations(tr.doc) : old },
    props: { decorations(state) { return this.getState(state); } },
  })],
});
