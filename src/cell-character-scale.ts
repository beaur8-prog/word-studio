import { Extension } from '@tiptap/core';
import type { Node as ProseMirrorNode } from '@tiptap/pm/model';
import { Plugin, PluginKey } from '@tiptap/pm/state';
import { Decoration, DecorationSet } from '@tiptap/pm/view';
import { characterScaleProperty, parseCharacterScale } from './character-scale';

const inlineTags = new Set(['SPAN', 'STRONG', 'B', 'EM', 'I', 'U', 'S', 'DEL', 'MARK']);
const inlineMarks = new Set(['textStyle', 'bold', 'italic', 'underline', 'strike', 'highlight']);
const excludedCharacter = /[\t\r\n\p{Bidi_Control}\u0590-\u08ff\ufb1d-\ufeff]/u;

function uniformScale(scales: number[]): number | null {
  const scale = scales[0];
  return scale >= 50 && scale <= 200 && scale !== 100 && scales.every(value => value === scale) ? scale : null;
}

function hangingIndent(left: number, right: string, first: string): number | null {
  const indent = parseFloat(first) || 0;
  if (parseFloat(right)) return null;
  if (!left && !indent) return 0;
  // A fully hanging first line starts at the cell's unindented text edge.
  // Only pixel lengths are comparable to the model's left indent.
  return Number.isFinite(left) && left > 0 && first.endsWith('px') &&
    indent < 0 && Math.abs(left + indent) < .05 ? left : null;
}

export function cellCharacterScaleDimensions(scale: number, hanging = 0) {
  // Lay out against the width that the scaled glyphs actually have available.
  // The painted box then fits the original cell without changing the strut.
  const factor = scale / 100;
  const width = `${100 / factor}%`;
  const horizontal = hanging ? `width:calc(${width} - ${hanging / factor}px);text-indent:${-hanging / factor}px` : `width:${width}`;
  return `${horizontal};transform:scaleX(${factor});transform-origin:left center`;
}

export function authoredCellParagraphScale(paragraph: HTMLParagraphElement, preserveHardBreak = false): number | null {
  if (paragraph.matches('[data-paragraph-border-box],[data-word-paragraph-bottom-border],[data-word-paragraph-reflection]')) return null;

  const scales: number[] = [];
  const contents = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT);
  while (contents.nextNode()) {
    const child = contents.currentNode;
    if (child instanceof HTMLElement) {
      if (preserveHardBreak && child instanceof HTMLBRElement) continue;
      if (!inlineTags.has(child.tagName) || child.hasAttribute('data-word-reflection')) return null;
      continue;
    }
    if (excludedCharacter.test(child.textContent || '')) return null;
    let parent = child.parentElement;
    let scale: number | null = null;
    while (parent && parent !== paragraph) {
      scale = parseCharacterScale(parent.style.getPropertyValue(characterScaleProperty));
      if (scale !== null) break;
      parent = parent.parentElement;
    }
    if (scale === null || scale === 100) return null;
    scales.push(scale);
  }
  return uniformScale(scales);
}

function hasExplicitGrid(table: HTMLTableElement) {
  const columns = [...table.querySelectorAll<HTMLTableColElement>(':scope > colgroup > col')];
  return columns.length > 0 && columns.every(column =>
    column.style.width.endsWith('px') && parseFloat(column.style.width) > 0);
}

export function layoutCellCharacterScale(root: HTMLElement) {
  // An explicit table width keeps its grid even when Word's saved layout is
  // autofit. A different cell's span does not change this cell's column.
  for (const cell of root.querySelectorAll<HTMLTableCellElement>('td')) {
    if (!cell.closest('tr') || cell.rowSpan !== 1 || !cell.children.length ||
        !cell.firstElementChild?.querySelector(`span[style*="${characterScaleProperty}"]`) ||
        ![...cell.children].every(child => child instanceof HTMLParagraphElement)) continue;
    const table = cell.closest('table');
    if (!table || table.dataset.wordTableWidthAuto === 'true') continue;
    const owner = table.parentElement?.closest<HTMLTableCellElement>('td,th');
    if (owner ? owner.tagName !== 'TD' || owner.rowSpan !== 1 || owner.colSpan !== 1 ||
      !!table.dataset.cellTablePosition || !!table.dataset.tableAnchor : cell.colSpan !== 1) continue;
    const paragraphs = [...cell.children] as HTMLParagraphElement[];
    if ((owner || paragraphs.length > 1) && !hasExplicitGrid(table)) continue;
    // Each independent paragraph wraps against the same cell width. Require
    // a uniform authored scale throughout a cell with several paragraphs.
    const indents = paragraphs.map(({ style }) => {
      const left = style.marginLeft;
      if (parseFloat(left) && !left.endsWith('px')) return null;
      return hangingIndent(parseFloat(left) || 0, style.marginRight, style.textIndent);
    });
    if (indents.some(indent => indent === null)) continue;
    const hanging = indents.find(indent => indent) || 0;
    if (hanging && (owner || paragraphs.length !== 1 || !hasExplicitGrid(table))) continue;
    const scale = uniformScale(paragraphs.map(paragraph => authoredCellParagraphScale(paragraph) ?? 0));
    if (scale !== null && (!hanging || scale < 100)) {
      for (const paragraph of paragraphs) paragraph.style.cssText += ';' + cellCharacterScaleDimensions(scale, hanging);
    }
  }
}

function nodeScale(paragraph: ProseMirrorNode): number | null {
  const attrs = paragraph.attrs;
  if (attrs.paragraphBorderBox || attrs.wordParagraphBottomBorder || attrs.wordParagraphReflection) return null;
  const scales: number[] = [];
  for (let index = 0; index < paragraph.childCount; index++) {
    const child = paragraph.child(index);
    if (!child.isText || excludedCharacter.test(child.text || '') || child.marks.some(mark => !inlineMarks.has(mark.type.name))) return null;
    const scale = parseCharacterScale(child.marks.find(mark => mark.type.name === 'textStyle')?.attrs.characterScale);
    if (scale === null || scale === 100) return null;
    scales.push(scale);
  }
  return uniformScale(scales);
}

function hasExplicitModelGrid(table: ProseMirrorNode) {
  const firstRow = table.firstChild;
  return !!firstRow?.childCount && firstRow.content.content.every(cell => {
    const widths = cell.attrs.colwidth;
    return Array.isArray(widths) && widths.length === cell.attrs.colspan &&
      widths.every(width => Number.isFinite(width) && width > 0);
  });
}

function cellDecorations(doc: ProseMirrorNode) {
  const decorations: Decoration[] = [];
  doc.descendants((row, position, table) => {
    if (row.type.name !== 'tableRow') return;
    if (!table || table.attrs.wordTableWidthAuto) return false;
    const context = doc.resolve(position);
    const owner = context.depth > 1 ? context.node(context.depth - 1) : null;
    const nested = owner?.type.name === 'tableCell' || owner?.type.name === 'tableHeader';
    if (nested && (owner.type.name !== 'tableCell' || owner.attrs.rowspan !== 1 || owner.attrs.colspan !== 1 ||
      table.attrs.cellTablePosition || table.attrs.tableAnchor || !hasExplicitModelGrid(table))) return false;
    row.forEach((cell, offset) => {
      if (cell.type.name !== 'tableCell' || cell.attrs.rowspan !== 1 || !nested && cell.attrs.colspan !== 1 ||
        !cell.childCount || (cell.childCount > 1 && !hasExplicitModelGrid(table)) ||
        !cell.content.content.every(paragraph => paragraph.type.name === 'paragraph')) return;
      const indents = cell.content.content.map(({ attrs }) =>
        hangingIndent(attrs.indent, attrs.rightIndent || '', attrs.firstLineIndent || ''));
      if (indents.some(indent => indent === null)) return;
      const hanging = indents.find(indent => indent) || 0;
      if (hanging && (nested || cell.childCount !== 1 || !hasExplicitModelGrid(table))) return;
      const scale = uniformScale(cell.content.content.map(paragraph => nodeScale(paragraph) ?? 0));
      if (scale === null || (hanging && scale >= 100)) return;
      cell.forEach((paragraph, innerOffset) => {
        const from = position + offset + innerOffset + 2;
        decorations.push(Decoration.node(from, from + paragraph.nodeSize, { style: cellCharacterScaleDimensions(scale, hanging) }));
      });
    });
  });
  return DecorationSet.create(doc, decorations);
}

const key = new PluginKey<DecorationSet>('cellCharacterScale');
export const CellCharacterScale = Extension.create({
  name: 'cellCharacterScale',
  addProseMirrorPlugins: () => [new Plugin({
    key,
    state: { init: (_, state) => cellDecorations(state.doc), apply: (transaction, previous) =>
      transaction.docChanged ? cellDecorations(transaction.doc) : previous },
    props: { decorations: state => key.getState(state) },
  })],
});
