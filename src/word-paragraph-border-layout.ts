import { Extension } from '@tiptap/core';
import { Plugin } from '@tiptap/pm/state';
import type { Node as PMNode } from '@tiptap/pm/model';
import { Decoration, DecorationSet } from '@tiptap/pm/view';
import { parseWordParagraphBottomBorder } from './word-paragraph-border';

function bottomLine(value: unknown) {
  const border = parseWordParagraphBottomBorder(value);
  return border?.style === 'single' ? { size: border.size, space: border.space ?? 0, color: border.color ?? 'auto' } : null;
}

function bottomBorderGroupEnds<T>(nodes: T[], description: (node: T) => unknown) {
  const borders = nodes.map(node => bottomLine(description(node)));
  return borders.flatMap((border, i) => border && JSON.stringify(border) !== JSON.stringify(borders[i + 1]) ?
    [{ target: nodes[i], border }] : []);
}

const marker = 'data-word-bottom-border-layout';
export const wordBottomBorderEnd = 'data-word-bottom-border-end';
export const wordFloatingBottomBorder = 'data-word-bottom-border-floating';
const variables = (border: NonNullable<ReturnType<typeof bottomLine>>) => ({
  '--word-bottom-border-width': `${border.size / 8}pt`,
  '--word-bottom-border-space': `${border.space}pt`,
  '--word-bottom-border-color': border.color === 'auto' ? 'currentColor' : `#${border.color}`,
});
function markBottomBorder(target: HTMLElement, border: NonNullable<ReturnType<typeof bottomLine>>) {
  target.setAttribute(marker, 'true');
  for (const [name, value] of Object.entries(variables(border))) target.style.setProperty(name, value);
}
export function floatingBottomBorderPaint(paragraph: HTMLElement, leftInset: number) {
  const border = bottomLine(paragraph.dataset.wordParagraphBottomBorder);
  if (!border) return null;
  const properties = { ...variables(border), '--word-bottom-border-inset': `${leftInset}px` };
  return { [marker]: 'true', [wordFloatingBottomBorder]: 'true',
    style: Object.entries(properties).map(([key, value]) => `${key}:${value}`).join(';') };
}
export function floatingBottomBorderDimensions(paragraph: HTMLElement) {
  const border = bottomLine(paragraph.dataset.wordParagraphBottomBorder);
  return border ? { gap: border.space * 4 / 3, stroke: border.size / 8 * 4 / 3 } : null;
}
export function applyFloatingBottomBorder(paragraph: HTMLElement, inset: number) {
  const border = bottomLine(paragraph.dataset.wordParagraphBottomBorder);
  if (!border) return;
  markBottomBorder(paragraph, border);
  paragraph.setAttribute(wordFloatingBottomBorder, 'true');
  paragraph.style.setProperty('--word-bottom-border-inset', `${inset}px`);
}
export function clearFloatingBottomBorders(root: HTMLElement) {
  for (const p of root.querySelectorAll<HTMLElement>(`[${wordFloatingBottomBorder}]`)) {
    p.removeAttribute(marker); p.removeAttribute(wordFloatingBottomBorder);
    for (const name of ['width', 'space', 'color', 'inset']) p.style.removeProperty(`--word-bottom-border-${name}`);
  }
}

/** The static tree is still in source order when group ends are determined.
 * The print fitter can subsequently move a whole group end onto a new page. */
export function layoutWordParagraphBottomBorders(root: HTMLElement) {
  const floating = !!root.querySelector('img[data-image-floating],[data-document-textbox],table[data-table-anchor]');
  const wrappedFlow = floating && !root.querySelector('[data-note-reference],table[data-table-anchor],[data-column-count]:not([data-column-count="1"])');
  if (floating && !wrappedFlow) return;
  for (const { target, border } of bottomBorderGroupEnds(Array.from(root.children), node =>
    node.matches('p,h1,h2,h3,h4,h5,h6') ? node.getAttribute('data-word-paragraph-bottom-border') : null)) {
    if (!(target instanceof HTMLElement) || !target.textContent?.trim() ||
        target.querySelector('img,svg,[data-note-reference],[data-docx-tab],[data-document-field],[data-page-number],[data-symbol-font],[data-word-overlap],[data-drawing-group-anchor]')) continue;
    if (floating) target.setAttribute(wordBottomBorderEnd, 'true');
    else markBottomBorder(target, border);
  }
}

function bottomBorderDecorations(doc: PMNode) {
  const paragraphs: { node: PMNode; offset: number }[] = [];
  doc.forEach((node, offset) => paragraphs.push({ node, offset }));
  const ends = bottomBorderGroupEnds(paragraphs, ({ node }) =>
    node.type.name === 'paragraph' || node.type.name === 'heading' ? node.attrs.wordParagraphBottomBorder : null);
  if (!ends.length) return DecorationSet.empty;
  let floating = false, canWrap = true;
  doc.descendants(node => {
    if (floating && !canWrap) return false;
    floating ||= !!node.attrs.floating || !!node.attrs.tableAnchor || node.type.name === 'documentTextBox';
    canWrap &&= !node.attrs.tableAnchor && node.type.name !== 'noteReference' &&
      !(node.type.name === 'columnGroup' && Number(node.attrs.count) !== 1);
  });
  if (floating && !canWrap) return DecorationSet.empty;
  return DecorationSet.create(doc, ends.flatMap(({ target: { node, offset }, border }) => {
    if (!node.textContent.trim()) return [];
    let complex = false;
    node.forEach(child => {
      complex ||= !child.isText && !['hardBreak', 'documentBookmark'].includes(child.type.name);
    });
    if (complex) return [];
    if (floating) return [Decoration.node(offset, offset + node.nodeSize, { [wordBottomBorderEnd]: 'true' })];
    const style = Object.entries(variables(border)).map(([key, value]) => `${key}:${value}`).join(';');
    return [Decoration.node(offset, offset + node.nodeSize, { [marker]: 'true', style })];
  }));
}

export const WordParagraphBottomBorderLayout = Extension.create({
  name: 'wordParagraphBottomBorderLayout',
  addProseMirrorPlugins: () => [new Plugin<DecorationSet>({
    state: { init: (_, state) => bottomBorderDecorations(state.doc),
      apply: (transaction, decorations) => transaction.docChanged ? bottomBorderDecorations(transaction.doc) : decorations },
    props: { decorations(state) { return this.getState(state); } },
  })],
});
