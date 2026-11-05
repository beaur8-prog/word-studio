import { Extension } from '@tiptap/core';
import { Plugin, type EditorState, type Transaction } from '@tiptap/pm/state';
import { Mapping } from '@tiptap/pm/transform';
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
  // Always own the live label. Removing a decoration when it equals the
  // saved snapshot would remove the base DOM attribute after undo.
  return DecorationSet.create(doc, [...labels].map(([entry, label]) =>
    Decoration.node(entry.pos, entry.pos + entry.node.nodeSize, { 'data-list-label': label })));
}
const listDepth = (doc: PMNode, pos: number) => {
  const resolved = doc.resolve(pos); let depth = 0;
  for (let d = resolved.depth; d > 0; d--) {
    const kind = resolved.node(d).type.name;
    if (['documentTextBox', 'documentNote'].includes(kind)) break;
    if (kind === 'listItem') depth++;
  }
  return depth;
};
const indentPixels = (value: unknown) => typeof value === 'string' && /^-?\d+(?:\.\d+)?(?:px|pt)$/.test(value)
  ? parseFloat(value) * (value.endsWith('pt') ? 4 / 3 : 1) : NaN;

/** List commands move existing items without updating imported Word metadata.
 * Map their preserved first paragraph through the real transaction so Tab and
 * Shift+Tab update level/indent together, in the same undo history event. */
function movedItems(transactions: readonly Transaction[], old: EditorState, state: EditorState) {
  if (!transactions.some(tr => tr.docChanged)) return null;
  const mapping = new Mapping(); transactions.forEach(tr => mapping.appendMapping(tr.mapping));
  type Item = { node: PMNode; pos: number; depth: number };
  const items: Item[] = [], levels = new Map<string, Map<number, PMNode>>();
  old.doc.descendants((node, pos) => {
    if (node.type.name !== 'listItem' || !node.attrs.wordListId || !node.firstChild?.isTextblock) return;
    items.push({ node, pos, depth: listDepth(old.doc, pos) });
    const id = String(node.attrs.wordListId), group = levels.get(id) || new Map<number, PMNode>();
    if (!group.has(Number(node.attrs.wordListLevel) || 0)) group.set(Number(node.attrs.wordListLevel) || 0, node);
    levels.set(id, group);
  });
  const mapped: { before: Item; node: PMNode; pos: number; delta: number }[] = [];
  for (const before of items) {
    const anchor = mapping.mapResult(before.pos + 2 + Math.min(1, before.node.firstChild!.content.size), 1);
    if (anchor.deleted) continue;
    const resolved = state.doc.resolve(Math.min(anchor.pos, state.doc.content.size));
    let d = resolved.depth; while (d > 0 && resolved.node(d).type.name !== 'listItem') d--;
    if (d) { const pos = resolved.before(d); mapped.push({ before, node: resolved.node(d), pos, delta: listDepth(state.doc, pos) - before.depth }); }
  }
  const tr = state.tr, updated = new Set<number>();
  for (const { before, node, pos, delta } of mapped) {
    const original = Number(before.node.attrs.wordListLevel) || 0;
    if (!delta || updated.has(pos) || node.attrs.wordListId !== before.node.attrs.wordListId ||
      (Number(node.attrs.wordListLevel) || 0) !== original || !node.firstChild?.isTextblock) continue;
    const level = Math.max(0, Math.min(8, original + delta)); if (level === original) continue;
    const group = levels.get(String(node.attrs.wordListId))!, source = group.get(original), target = group.get(level);
    let step = 48;
    const ordered = [...group].sort(([a], [b]) => a - b);
    for (let i = 1; i < ordered.length; i++) {
      const change = (Number(ordered[i][1].firstChild?.attrs.indent) - Number(ordered[i - 1][1].firstChild?.attrs.indent)) / (ordered[i][0] - ordered[i - 1][0]);
      if (change > 0 && Number.isFinite(change)) { step = change; break; }
    }
    const offset = source && target ? Number(target.firstChild!.attrs.indent) - Number(source.firstChild!.attrs.indent) : step * (level - original);
    const format = target?.attrs.wordListFormat || node.attrs.wordListFormat;
    let pattern = target?.attrs.wordListPattern;
    if (typeof pattern !== 'string') {
      const oldPattern = String(node.attrs.wordListPattern || `%${original + 1}.`), variables = oldPattern.match(/%[1-9]/g) || [];
      pattern = format === 'bullet' ? ['•', '◦', '▪'][level % 3]
        : variables.length > 1 && /^(?:%[1-9]\.)+$/.test(oldPattern) ? Array.from({ length: level + 1 }, (_, i) => `%${i + 1}.`).join('')
        : variables.length === 1 ? oldPattern.replace(/%[1-9]/, `%${level + 1}`) : oldPattern;
    }
    tr.setNodeMarkup(pos, undefined, { ...node.attrs, wordListLevel: typeof node.attrs.wordListLevel === 'string' ? String(level) : level,
      wordListFormat: format, wordListPattern: pattern, wordListSuffix: target ? target.attrs.wordListSuffix : node.attrs.wordListSuffix });
    const paragraph = node.firstChild, attrs: Record<string, unknown> = { ...paragraph.attrs, indent: (Number(paragraph.attrs.indent) || 0) + offset };
    const from = indentPixels(source?.firstChild?.attrs.firstLineIndent), to = indentPixels(target?.firstChild?.attrs.firstLineIndent), current = indentPixels(paragraph.attrs.firstLineIndent);
    if (Number.isFinite(from) && Number.isFinite(to) && Number.isFinite(current) && from !== to) attrs.firstLineIndent = `${current + to - from}px`;
    tr.setNodeMarkup(pos + 1, undefined, attrs); updated.add(pos);
  }
  return tr.docChanged ? tr : null;
}
export const WordListLabels = Extension.create({
  name: 'wordListLabels',
  addProseMirrorPlugins: () => [new Plugin<DecorationSet>({
    state: { init: (_, state) => decorations(state.doc), apply: (tr, old) => tr.docChanged ? decorations(tr.doc) : old },
    props: { decorations(state) { return this.getState(state); } },
    appendTransaction: movedItems,
  })],
});
