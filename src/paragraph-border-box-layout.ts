import { Extension } from '@tiptap/core';
import type { Node } from '@tiptap/pm/model';
import { Plugin } from '@tiptap/pm/state';
import { Decoration, DecorationSet } from '@tiptap/pm/view';
import { cssPixels } from './hwp-cell-line';
import { type ParagraphBorderBox, paragraphBorderBoxAttribute, paragraphBorderBoxEnd, paragraphBorderBoxStart, parseParagraphBorderBox } from './paragraph-border-box';

function borderMarks<T>(items: T[], value: (item: T) => unknown) {
  const boxes = items.map(item => parseParagraphBorderBox(value(item)));
  const keys = boxes.map(box => box && JSON.stringify(box));
  return boxes.flatMap((box, i) => box ? [{
    item: items[i], box, start: keys[i - 1] !== keys[i], end: keys[i + 1] !== keys[i],
  }] : []);
}

function rightMargin(value: unknown) {
  if (typeof value === 'string' && value.startsWith('-')) return -(cssPixels(value.slice(1)) ?? 0);
  return cssPixels(value) ?? 0;
}

function decorationsForGroup(items: { node: Node; position: number }[]) {
  return borderMarks(items, ({ node }) => node.type.name === 'paragraph' || node.type.name === 'heading' ? node.attrs.paragraphBorderBox : null)
    .flatMap(({ item: { node, position }, box, start, end }) => {
      const attrs = attributes(box, Number(node.attrs.indent) || 0, rightMargin(node.attrs.rightIndent), start, end);
      return attrs ? [Decoration.node(position, position + node.nodeSize, attrs)] : [];
    });
}

function attributes(box: ParagraphBorderBox, left: number, right: number, start: boolean, end: boolean): Record<string, string> | null {
  // Use the existing paragraph margins as the paint area. Matching padding
  // keeps the paragraph's text area and all line heights in the same place.
  const leftInset = Math.min(Math.max(0, left), box.leftSpace * 4 / 3);
  const rightInset = Math.min(Math.max(0, right), box.rightSpace * 4 / 3);
  const stroke = box.size / 8;
  // Keep the standard properties even if this paragraph gives no safe
  // margin to paint into. A stroke must not intrude into the text area.
  if (Math.min(leftInset, rightInset) < stroke * 4 / 3) return null;
  return {
    'data-paragraph-border-box-layout': 'true',
    ...(start ? { [paragraphBorderBoxStart]: 'true' } : {}),
    ...(end ? { [paragraphBorderBoxEnd]: 'true' } : {}),
    style: '--paragraph-box-color:' + (box.color === 'auto' ? 'currentColor' : '#' + box.color) +
      ';--paragraph-box-stroke:' + stroke + 'pt' +
      ';--paragraph-box-left-margin:' + (left - leftInset) + 'px' +
      ';--paragraph-box-right-margin:' + (right - rightInset) + 'px' +
      ';--paragraph-box-left-inset:' + leftInset + 'px' +
      ';--paragraph-box-right-inset:' + rightInset + 'px',
  };
}

function borderDecorations(doc: Node) {
  const decorations: Decoration[] = [];
  function visit(parent: Node, position: number) {
    const items: { node: Node; position: number }[] = [];
    parent.forEach((node, offset) => {
      const at = position + offset;
      items.push({ node, position: at });
      if (node.type.name === 'columnGroup' || node.type.name === 'sectionContent') visit(node, at + 1);
    });
    decorations.push(...decorationsForGroup(items));
  }
  visit(doc, 0);
  return DecorationSet.create(doc, decorations);
}

export function layoutParagraphBorderBoxes(root: HTMLElement) {
  if (!root.querySelector('[' + paragraphBorderBoxAttribute + ']')) return;
  for (const group of [root, ...root.querySelectorAll<HTMLElement>('[data-column-count],[data-section-content]')]) {
    if (group.closest('table,li,[data-textbox-content],[data-document-note]')) continue;
    for (const { item, box, start, end } of borderMarks([...group.children] as HTMLElement[], node =>
      node.matches('p,h1,h2,h3,h4,h5,h6') ? node.getAttribute(paragraphBorderBoxAttribute) : null)) {
      const attrs = attributes(box, parseFloat(item.style.marginLeft) || 0, rightMargin(item.style.marginRight), start, end);
      if (!attrs) continue;
      item.style.cssText += ';' + attrs.style;
      for (const name of Object.keys(attrs)) if (name !== 'style') item.setAttribute(name, attrs[name]);
    }
  }
}

export const ParagraphBorderBoxLayout = Extension.create({
  name: 'paragraphBorderBoxLayout',
  addProseMirrorPlugins: () => [new Plugin<DecorationSet>({
    state: { init: (_, state) => borderDecorations(state.doc),
      apply: (transaction, decorations) => transaction.docChanged ? borderDecorations(transaction.doc) : decorations },
    props: { decorations(state) { return this.getState(state); } },
  })],
});
