import { hwpBoxedDigitAttribute, parseHwpBoxedDigit } from './hwp-boxed-digit';
import {DocumentDrawingAnchor} from './drawing-group-anchor';
import { Node, Extension, Mark } from "@tiptap/core";
import StarterKit from "@tiptap/starter-kit";
import Underline from '@tiptap/extension-underline';
import Strike from '@tiptap/extension-strike';
import { TextStyleKit } from "@tiptap/extension-text-style";
import TextAlign from "@tiptap/extension-text-align";
import { TableKit } from "@tiptap/extension-table";
import { BlockImage, InlineImage } from './document-image';
import Highlight from "@tiptap/extension-highlight";
import Subscript from "@tiptap/extension-subscript";
import Superscript from "@tiptap/extension-superscript";
import Placeholder from "@tiptap/extension-placeholder";
import { CellAppearance } from './table-format';
import { TableRowResize } from './table-row-resize';
import { Pagination } from './pagination';
import { CellCharacterScale } from './cell-character-scale';
import { ImportedFormat, StudioTableView } from './imported-format';
import { parseWordReflection, wordReflectionAttribute } from './word-reflection';
import { DocumentTab, DocumentTabs } from './document-tabs';
import { WordLineMetrics } from './word-line-metrics';
import { WordCellSpacing } from './word-cell-spacing';
import { WordContextualSpacing } from './word-contextual-spacing';
import { WordParagraphBottomBorderLayout } from './word-paragraph-border-layout';
import { ParagraphBorderBoxLayout } from './paragraph-border-box-layout';
import { WordListLabels } from './word-list-labels';
import { WordHeadingDefaults } from './word-heading-defaults';
import { HwpSourceAdvanceEditor } from './hwp-text-advance-editor';
import { HwpCellLineMetrics } from './hwp-cell-line-metrics';
import { HwpBodyLineMetrics } from './hwp-body-line-metrics';
import { WordOverlap } from './word-overlap';
import { DocumentBookmark } from './document-anchors';
import { NoteReference, DocumentNote, DocumentField, ColumnGroup, SectionContent } from './document-stories';
import { DocumentSymbol } from './document-symbols';
import { ColumnLayout } from './column-layout';
import { FloatingLayout } from './floating-layout';
import { TableAnchorLayout } from './table-anchor';
import { TableNavigation } from './table-navigation';
import { SelectionNavigation } from './selection-navigation';
import { DocumentTextBox, TextBoxLayout } from './textbox-frame';
import { DOMParser as ProseMirrorDOMParser } from '@tiptap/pm/model';
import { getSchema } from '@tiptap/core';

const HwpBoxedDigit = Node.create({
  name: 'hwpBoxedDigit', group: 'inline', inline: true, atom: true,
  addAttributes: () => ({
    box: { default: '', renderHTML: attrs => ({ [hwpBoxedDigitAttribute]: attrs.box }) },
  }),
  parseHTML: () => [{
    tag: 'span[' + hwpBoxedDigitAttribute + ']',
    getAttrs: element => {
      const value = element.getAttribute(hwpBoxedDigitAttribute);
      return parseHwpBoxedDigit(value)?.digit === element.textContent ? { box: value } : false;
    },
  }],
  renderHTML: ({ node, HTMLAttributes }) => {
    const box = parseHwpBoxedDigit(node.attrs.box);
    return ['span', { ...HTMLAttributes, class: 'hwp-boxed-digit', style: 'font-size:' + (box ? box.size / 2 : 11) + 'pt' },
      ['span', {}, box?.digit ?? '']];
  },
  renderText: ({ node }) => parseHwpBoxedDigit(node.attrs.box)?.digit ?? '',
});

export const PageBreak = Node.create({
  name: "pageBreak",
  group: "block",
  atom: true,
  selectable: true,
  parseHTML: () => [{ tag: "div[data-page-break]" }],
  renderHTML: () => [
    "div",
    { "data-page-break": "true", class: "page-break" },
    ["span", {}, "페이지 나누기"],
  ],
});
export const ColumnBreak = Node.create({
  name: 'columnBreak', group: 'block', atom: true,
  parseHTML: () => [{ tag: 'div[data-column-break]' }],
  renderHTML: () => ['div', { 'data-column-break': 'true', class: 'column-break', title: '단 나누기' }],
});
export const CommentAnchor = Mark.create({
  name: 'commentAnchor', inclusive: false,
  addAttributes: () => ({ ids: { default: '', parseHTML: el => el.getAttribute('data-comment-ids'), renderHTML: a => ({ 'data-comment-ids': a.ids }) } }),
  parseHTML: () => [{ tag: 'span[data-comment-ids]' }],
  renderHTML: ({ HTMLAttributes }) => ['span', { ...HTMLAttributes, class: 'comment-anchor' }, 0],
});
const WordReflection = Mark.create({
  name: 'wordReflection',
  // CommentAnchor consumes its span. Read both marks on that same span.
  priority: 101,
  inclusive: false,
  addAttributes: () => ({
    effect: {
      default: null,
      parseHTML: element => {
        const effect = parseWordReflection(element.getAttribute(wordReflectionAttribute));
        return effect ? JSON.stringify(effect) : null;
      },
      renderHTML: attrs => {
        const effect = parseWordReflection(attrs.effect);
        return effect ? { [wordReflectionAttribute]: JSON.stringify(effect) } : {};
      },
    },
  }),
  parseHTML: () => [{
    tag: `span[${wordReflectionAttribute}]`,
    consuming: false,
    getAttrs: element => parseWordReflection((element as HTMLElement).getAttribute(wordReflectionAttribute)) ? {} : false,
  }],
  renderHTML: ({ HTMLAttributes }) => ['span', HTMLAttributes, 0],
});
export const SectionBreak = Node.create({
  name: 'sectionBreak', group: 'block', atom: true,
  addAttributes: () => ({ page: { default: '', parseHTML: el => el.getAttribute('data-section-page'), renderHTML: a => ({ 'data-section-page': a.page }) } }),
  parseHTML: () => [{ tag: 'div[data-section-page]' }],
  renderHTML: ({ HTMLAttributes }) => ['div', { ...HTMLAttributes, class: 'section-break' }, '구역 나누기'],
});
const ParagraphFormat = Extension.create({
  name: "paragraphFormat",
  addGlobalAttributes() {
    return [
      {
        types: ["paragraph", "heading"],
        attributes: {
          indent: {
            default: 0,
            parseHTML: (e) => parseFloat(e.style.marginLeft) || 0,
            renderHTML: (a) =>
              a.indent ? { style: `margin-left:${a.indent}px` } : {},
          },
          lineSpacing: {
            default: null,
            parseHTML: (e) => e.style.lineHeight || null,
            renderHTML: (a) => a.lineSpacing ? {
              style: `line-height:${a.lineSpacing}` + (/(?:px|pt)$/.test(a.lineSpacing)
                ? `;--exact-line-height:${a.lineSpacing}` : ''),
              // CSS gives each differently sized run its own half-leading.
              // An exact Word/Hancom line instead has one paragraph strut.
              ...(/(?:px|pt)$/.test(a.lineSpacing) && !['auto', 'atLeast'].includes(a.wordLineRule)
                ? { 'data-exact-line-height': 'true' } : {}),
            } : {},
          },
        },
      },
    ];
  },
});
export const editorExtensions = [
  StarterKit.configure({
    link: { openOnClick: false, protocols: ["https", "http", "mailto"] },
    underline: false, strike: false,
    // These containers already have an editable end. A synthetic paragraph
    // after a full final table/column otherwise creates an extra blank page.
    // The gap cursor can insert a following paragraph when the user types.
    trailingNode: { notAfter: ['columnGroup', 'sectionContent', 'table'] },
  }),
  Underline.extend({ parseHTML() { return [...(this.parent?.() || []), { style: 'text-decoration-line', consuming: false, getAttrs: value => String(value).includes('underline') ? {} : false }]; } }),
  Strike.extend({ parseHTML() { return [...(this.parent?.() || []), { style: 'text-decoration-line', consuming: false, getAttrs: value => String(value).includes('line-through') ? {} : false }]; } }),
  TextStyleKit,
  TextAlign.configure({ types: ["heading", "paragraph"] }),
  TableKit.configure({
    // Imported merged tables can contain 1–2 px grid subdivisions. Clamping
    // those invisible grid columns to 25 px expands the whole table.
    table: { resizable: true, cellMinWidth: 1, allowTableNodeSelection: true, View: StudioTableView },
  }),
  CellAppearance,
  TableRowResize,
  TableNavigation, SelectionNavigation,
  DocumentTextBox, TextBoxLayout,
  TableAnchorLayout,
  BlockImage.configure({ allowBase64: true }),
  InlineImage, DocumentDrawingAnchor,
  DocumentTab,
  DocumentBookmark,
  NoteReference, DocumentNote, DocumentField, ColumnGroup, SectionContent,
  DocumentSymbol,
  WordOverlap,
  HwpBoxedDigit,
  DocumentTabs,
  Highlight.configure({ multicolor: true }),
  Subscript,
  Superscript,
  Placeholder.configure({ placeholder: "여기에 문서를 작성하세요…" }),
  PageBreak,
  ColumnBreak,
  SectionBreak,
  CommentAnchor, WordReflection,
  ParagraphFormat,
  ImportedFormat,
  WordLineMetrics,
  WordCellSpacing,
  WordContextualSpacing,
  WordParagraphBottomBorderLayout,
  ParagraphBorderBoxLayout,
  WordListLabels,
  WordHeadingDefaults,
  CellCharacterScale,
  HwpCellLineMetrics,
  HwpSourceAdvanceEditor,
  HwpBodyLineMetrics,
  ColumnLayout,
  FloatingLayout,
  Pagination,
];

export function parseDocumentHtml(html: string) {
  const root = document.createElement('div'); root.innerHTML = html;
  return ProseMirrorDOMParser.fromSchema(getSchema(editorExtensions)).parse(root, { preserveWhitespace: true }).toJSON();
}
