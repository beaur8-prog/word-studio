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
import { ImportedFormat, StudioTableView } from './imported-format';
import { DocumentTab, DocumentTabs } from './document-tabs';
import { WordLineMetrics } from './word-line-metrics';
import { WordOverlap } from './word-overlap';
import { DocumentBookmark } from './document-anchors';
import { NoteReference, DocumentNote, DocumentField, ColumnGroup, SectionContent } from './document-stories';
import { DocumentSymbol } from './document-symbols';
import { ColumnLayout } from './column-layout';
import { FloatingLayout } from './floating-layout';
import { TableAnchorLayout } from './table-anchor';
import { DOMParser as ProseMirrorDOMParser } from '@tiptap/pm/model';
import { getSchema } from '@tiptap/core';

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
            renderHTML: (a) =>
              a.lineSpacing ? { style: `line-height:${a.lineSpacing}` } : {},
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
    // outside a full final column row otherwise creates an extra blank page.
    trailingNode: { notAfter: ['columnGroup', 'sectionContent'] },
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
  TableAnchorLayout,
  BlockImage.configure({ allowBase64: true }),
  InlineImage,
  DocumentTab,
  DocumentBookmark,
  NoteReference, DocumentNote, DocumentField, ColumnGroup, SectionContent,
  DocumentSymbol,
  WordOverlap,
  DocumentTabs,
  Highlight.configure({ multicolor: true }),
  Subscript,
  Superscript,
  Placeholder.configure({ placeholder: "여기에 문서를 작성하세요…" }),
  PageBreak,
  ColumnBreak,
  SectionBreak,
  CommentAnchor,
  ParagraphFormat,
  ImportedFormat,
  WordLineMetrics,
  ColumnLayout,
  FloatingLayout,
  Pagination,
];

export function parseDocumentHtml(html: string) {
  const root = document.createElement('div'); root.innerHTML = html;
  return ProseMirrorDOMParser.fromSchema(getSchema(editorExtensions)).parse(root, { preserveWhitespace: true }).toJSON();
}
