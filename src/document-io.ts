import { activeHwpLatinFont, hwpLatinFontNamespace } from './hwp-latin-font';
import { parseHwpNativeGrid } from './hwp-native-grid';
import { parseHwpRelatedTableLayout } from './hwp-related-table-layout';
import { parseHwpNestedTableLayout } from './hwp-nested-table-layout';
import { parseHwpCenteredCellLabels } from './hwp-centered-cell-labels';
import { parseHwpTableContinuation } from './hwp-table-continuation';
import { cssPoints } from './css-length';
import { HwpBoxedDigitRun } from './docx-hwp-boxed-digit';
import { parseHwpBoxedDigit } from './hwp-boxed-digit';
import { countDocxCheckboxes } from './docx-checkboxes';
import { readWordBackground, validWordBackground, writeWordBackground } from './docx-background';
import { W } from './docx-format';
import { parseCharacterScale } from './character-scale';
import { wordDefaultHangingSpacesProperty } from './word-trailing-spaces';
import { sameSectionFrameSettings } from './section-frame-settings';
import { encodeDocxXmlParts } from './docx-package-text';
import { paragraphTextAlignment } from './paragraph-text-alignment';
import { exportedRowCantSplit } from './row-split-policy';
import {parseDrawingGroup} from './drawing-group';
import {writeDrawingGroups} from './docx-drawing-groups';
import { wordListFontName } from './word-list-marker';
import { prepareTextBoxes, textBoxAnchorStyle, writeTextBoxes, type TextBoxExport } from './docx-textboxes';
import { textBoxFrame } from './textbox-frame';
import { wordHeadingMarks } from './word-heading-defaults';
import { parseWordParagraphBottomBorder } from './word-paragraph-border';
import { paragraphBorderBoxOOXML, parseParagraphBorderBox } from './paragraph-border-box';
import { parseWordReflection, wordReflectionXmlAttributes } from './word-reflection';
import { sanitizeImportedHtml, validImage, safeLink, RASTER_DATA } from "./document-html";
import { activeHwpCellLine, cssPixels, encodeHwpCellLine, hwpCellFont, parseHwpCellLine } from './hwp-cell-line';
import { HwpCellParagraph } from './docx-hwp-cell-line';
import { measureHwpExportLayout } from './hwp-cell-line-export';
import { parseHwpCellRasterFill } from './hwp-cell-raster-fill';
export { sanitizeImportedHtml } from "./document-html";
import type { JSONContent } from "@tiptap/react";
import JSZip from "jszip";
import mammoth from "mammoth";
import {
  AlignmentType,
  BorderStyle,
  Document as DocxDocument,
  ExternalHyperlink,
  InternalHyperlink,
  NumberValueElement, ImportedXmlComponent,
  Bookmark,
  Footer,
  Header,
  HeadingLevel,
  ImageRun,
  LevelFormat,
  LevelSuffix,
  LineRuleType,
  Packer,
  PageBreak,
  ColumnBreak,
  PageNumber,
  PageOrientation,
  Paragraph,
  Table,
  TableLayoutType,
  TableAnchorType,
  RelativeHorizontalPosition,
  TableCell,
  TableRow,
  TextRun,
  CommentRangeStart, CommentRangeEnd, CommentReference,
  UnderlineType,
  WidthType,
  VerticalAlignTable,
  HeightRule,
  Tab,
  TabStopType,
  SimpleField,
  FootnoteReferenceRun, EndnoteReferenceRun, SectionType, SymbolRun,
} from "docx";
import type {
  IParagraphOptions,
  IRunOptions,
  NumberingDefinition,
  ParagraphChild,
  ITableCellBorders,
} from "docx";
import { defaultPage, pageMargins, basePageMargins, paperSize, activeDocxPrintPaper, wordLanguageTag,
  activeHwpHorizontalMargin, parseHwpHorizontalMargins } from "./types";
import { prepareDocxMetadata, TABLE_GAP_STYLE, TABLE_ANCHOR_STYLE, TABLE_OFFSET_STYLE, SECTION_BOUNDARY_STYLE, TABLE_END_STYLE } from './docx-metadata';
import { emptyLineDocxProperty, emptyLineCompatibilityWarning } from './page-empty-lines';
import { normalizeDocxTables } from './docx-table-grid';
import { parseTabStops, parseWordNumberingTabs } from './document-tabs';
import { overlapParts } from './word-overlap';
import { rowHeight, MAX_ROW_HEIGHT } from './table-row-resize';
import { hwpRowReserve } from './hwp-row-reserve';
import { parseHwpIntactRow } from './hwp-intact-row';
import { storyHtml, documentNotes, attachDocumentNotes, structureDocxSections, writeHwpFootnoteLayouts } from './docx-stories';
import { hwpFootnoteMark, parseHwpFootnoteLayout } from './hwp-footnote-layout';
import { parseDocumentHtml } from './extensions';
import { imageFloating, restoreImageGeometry, type ImageGeometryRecord } from './image-floating';
import type { PageSettings, StudioDocument } from "./types";
import { PreservedTable, PreservedTableCell, PreservedTableRow } from './docx-table-cell';
import { cellTablePosition } from './cell-table-position';
import { tableAnchor, tableWrapFits } from './table-anchor';
import { hwpInlineTableBottom, hwpInlineTableTop, hwpKeptTablePage, hwpTableTopGap } from './table-anchor-data';
import { encodeHwpBodyLine, hwpBodyEmptyLine, hwpBodyTextLine, hwpBodyFallback, type HwpBodyLine } from './hwp-body-line';
import documentLimits from '../electron/document-limits.json';
const bookmarkNamesProperty = 'WordStudio.BookmarkNames';
const wordStudioCreator = 'Word Studio';
const wordStudioDescription = 'Created with Word Studio';
const exactPaperNamespace = 'https://word-studio.app/ooxml/paper-size';
const markupCompatibility = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const xmlNamespaces = 'http://www.w3.org/2000/xmlns/';
const twipToMm = 25.4 / 1440;
const mmToTwip = (mm: number) => Math.round((mm * 1440) / 25.4);

function exactPaperDimensions(size: Element | undefined, smaller: number, larger: number): [number, number] | undefined {
  const recorded = size?.getAttributeNS(exactPaperNamespace, 'sizeMm');
  if (!recorded || recorded.length > 64 || !/^\d+(?:\.\d+)? \d+(?:\.\d+)?$/.test(recorded)) return;
  const [width, height] = recorded.split(' ').map(Number);
  if (width >= 25 && width <= height && height <= 1320 &&
      mmToTwip(width) === smaller && mmToTwip(height) === larger) return [width, height];
}


export interface ImportResult {
  document: StudioDocument;
  html?: string;
  warnings: string[];
}
export const DOCX_EXPORT_NOTICE =
  "DOCX에는 본문 서식, 표, 그림, 구역 설정과 검토 메모를 저장합니다. 편집 상태의 백업에는 Word Studio 형식도 함께 사용하세요.";
const MAX_FILE_BYTES = documentLimits.documentFileBytes;
const MAX_INFLATED_BYTES = documentLimits.documentInflatedBytes;
export function documentFileLimit(extension: string | undefined): number {
  return extension === 'wordstudio' ? documentLimits.studioDocumentBytes : extension === 'html' || extension === 'htm' ? documentLimits.hwpEditableHtmlBytes : MAX_FILE_BYTES;
}
const INLINE = new Set(["text", "hardBreak", "inlineImage", "documentTab", "wordOverlap", "hwpBoxedDigit", "documentBookmark", "noteReference", "documentField", "documentSymbol", "documentDrawingAnchor"]);
const BLOCKS = new Set([
  "paragraph",
  "heading",
  "blockquote",
  "bulletList",
  "orderedList",
  "codeBlock",
  "horizontalRule",
  "pageBreak",
  "columnBreak",
  "sectionBreak",
  "image",
  "table",
  "documentNote", "columnGroup", "sectionContent", "documentTextBox",
]);
const MARKS = new Set([
  "bold",
  "italic",
  "underline",
  "strike",
  "subscript",
  "superscript",
  "textStyle",
  "highlight",
  "link",
  "code",
  "commentAnchor", "wordReflection",
]);

function fail(message: string): never {
  throw new Error(message);
}
function object(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}
function boundedString(value: unknown, max: number): value is string {
  return typeof value === "string" && value.length <= max;
}
function finite(value: unknown, min: number, max: number): value is number {
  return (
    typeof value === "number" &&
    Number.isFinite(value) &&
    value >= min &&
    value <= max
  );
}
function validDate(value: unknown): value is string {
  return (
    boundedString(value, 50) &&
    /^\d{4}-\d{2}-\d{2}T/.test(value) &&
    Number.isFinite(Date.parse(value))
  );
}
function keysOnly(
  value: Record<string, unknown>,
  keys: string[],
  label: string,
) {
  if (Object.keys(value).some((key) => !keys.includes(key)))
    fail(`${label}에 지원하지 않는 속성이 있습니다.`);
}
/** Validate before loading JSON into ProseMirror; never silently discard unknown document data. */
export function validateDocument(input: unknown): StudioDocument {
  if (!object(input)) fail("Word Studio 문서가 올바른 JSON 객체가 아닙니다.");
  keysOnly(
    input,
    ["id", "format", "version", "title", "content", "page", "wordDefaultLanguage", "wordAutoHyphenation", "wordWrapTrailSpaces", "wordDefaultHangingSpaces", "comments", "updatedAt"],
    "문서",
  );
  if (input.format !== "word-studio" || input.version !== 1)
    fail("지원하지 않는 Word Studio 문서 버전입니다.");
  if (input.id !== undefined && (typeof input.id !== 'string' || !/^[a-zA-Z0-9_-]{1,128}$/.test(input.id))) fail('문서 ID가 올바르지 않습니다.');
  if (input.wordDefaultLanguage !== undefined) {
    const language = input.wordDefaultLanguage;
    if (!object(language)) fail('Word 기본 언어 설정이 올바르지 않습니다.');
    keysOnly(language, ['value', 'eastAsia', 'bidirectional'], 'Word 기본 언어 설정');
    const tags = Object.values(language).filter(tag => tag !== undefined);
    if (!tags.length || !tags.every(wordLanguageTag)) fail('Word 기본 언어 설정이 올바르지 않습니다.');
  }
  if (input.wordAutoHyphenation !== undefined && typeof input.wordAutoHyphenation !== 'boolean')
    fail('Word 자동 하이픈 설정이 올바르지 않습니다.');
  if (input.wordWrapTrailSpaces !== undefined && typeof input.wordWrapTrailSpaces !== 'boolean')
    fail('Word 줄 끝 공백 설정이 올바르지 않습니다.');
  if (input.wordDefaultHangingSpaces !== undefined && input.wordDefaultHangingSpaces !== true)
    fail('Word 기본 공백 조판 설정이 올바르지 않습니다.');
  if (!boundedString(input.title, 512) || !validDate(input.updatedAt))
    fail("문서 제목 또는 저장 날짜가 올바르지 않습니다.");
  const page = input.page;
  if (!object(page)) fail("용지 설정이 없습니다.");
  keysOnly(
    page,
    ["size", "width", "height", "docxPrintPaper", "hwpHorizontalMargins", "orientation", "margin", "margins", "header", "footer", "pageNumbers",
      'differentFirstPage', 'differentOddEven', 'firstHeader', 'firstFooter', 'evenHeader', 'evenFooter',
      'firstPageNumbers', 'evenPageNumbers', 'headerDistance', 'footerDistance', 'gutter', 'gutterTop', 'breakType', 'columns', 'hideEmptyLines', 'suppressBottomSpacing',
      'headerHtml','footerHtml','firstHeaderHtml','firstFooterHtml','evenHeaderHtml','evenFooterHtml','wordBackground'],
    "용지 설정",
  );
  if (
    !["A4", "Letter", "Custom"].includes(String(page.size)) ||
    !["portrait", "landscape"].includes(String(page.orientation)) ||
    !finite(page.margin, 0, 75) ||
    !boundedString(page.header, 10000) ||
    !boundedString(page.footer, 10000) ||
    typeof page.pageNumbers !== "boolean"
  )
    fail("용지 설정이 올바르지 않습니다.");
  if ((page.size === 'Custom' || page.width !== undefined || page.height !== undefined) &&
    (![page.width, page.height].every(n => finite(n, 25, 1320)))) fail('사용자 용지 크기가 올바르지 않습니다.');
  if (page.hwpHorizontalMargins !== undefined && (!Array.isArray(page.hwpHorizontalMargins) ||
      !parseHwpHorizontalMargins(page.hwpHorizontalMargins))) fail('한글 원본의 좌우 여백 단위가 올바르지 않습니다.');
  if (page.docxPrintPaper !== undefined && (!Array.isArray(page.docxPrintPaper) || page.docxPrintPaper.length !== 2 ||
      !page.docxPrintPaper.every(mm => finite(mm, 25, 1320)) || page.docxPrintPaper[0] > page.docxPrintPaper[1]))
    fail('DOCX 인쇄 용지 크기가 올바르지 않습니다.');
  if (page.margins !== undefined) {
    if (!object(page.margins)) fail('방향별 여백이 올바르지 않습니다.');
    keysOnly(page.margins, ['top', 'right', 'bottom', 'left'], '방향별 여백');
    if (['top', 'right', 'bottom', 'left'].some(k => !finite((page.margins as Record<string, unknown>)[k], 0, 75))) fail('방향별 여백이 올바르지 않습니다.');
  }
  for (const key of ['differentFirstPage', 'differentOddEven', 'firstPageNumbers', 'evenPageNumbers', 'gutterTop', 'hideEmptyLines', 'suppressBottomSpacing'])
    if (page[key] !== undefined && typeof page[key] !== 'boolean') fail('용지 옵션이 올바르지 않습니다.');
  for (const key of ['firstHeader', 'firstFooter', 'evenHeader', 'evenFooter'])
    if (page[key] !== undefined && !boundedString(page[key], 10000)) fail('머리글·바닥글이 올바르지 않습니다.');
  for (const key of ['headerDistance', 'footerDistance', 'gutter'])
    if (page[key] !== undefined && !finite(page[key], 0, 75)) fail('용지 여백이 올바르지 않습니다.');
  if (page.breakType !== undefined && !['nextPage','continuous','oddPage','evenPage','nextColumn'].includes(String(page.breakType))) fail('구역 나누기 종류가 올바르지 않습니다.');
  if (page.columns !== undefined) {
    if (!object(page.columns) || !finite(page.columns.count, 1, 12) || !Number.isInteger(page.columns.count) || !finite(page.columns.space, 0, 1000)) fail('다단 설정이 올바르지 않습니다.');
    keysOnly(page.columns, ['count','space','widths','balance'], '다단');
    if (page.columns.balance !== undefined && typeof page.columns.balance !== 'boolean') fail('단 균형 설정이 올바르지 않습니다.');
    if (page.columns.widths !== undefined && (!Array.isArray(page.columns.widths) || page.columns.widths.length !== page.columns.count || !page.columns.widths.every(w => finite(w, 1, 10000)))) fail('단 너비가 올바르지 않습니다.');
  }
  for (const key of ['headerHtml','footerHtml','firstHeaderHtml','firstFooterHtml','evenHeaderHtml','evenFooterHtml'])
    if (page[key] !== undefined && !boundedString(page[key], MAX_FILE_BYTES)) fail('머리글·바닥글 본문이 너무 큽니다.');
  if (page.wordBackground !== undefined && !validWordBackground(page.wordBackground)) fail('페이지 배경 정보가 올바르지 않습니다.');
  if (!Array.isArray(input.comments) || input.comments.length > 10000)
    fail("검토 메모 목록이 올바르지 않습니다.");
  const ids = new Set<string>();
  for (const comment of input.comments) {
    if (!object(comment)) fail("검토 메모가 올바르지 않습니다.");
    keysOnly(
      comment,
      ["id", "quote", "text", "createdAt", "resolved", "author", "initials", "html"],
      "검토 메모",
    );
    if (
      !boundedString(comment.id, 128) ||
      !comment.id ||
      ids.has(comment.id) ||
      !boundedString(comment.quote, 100000) ||
      !boundedString(comment.text, 100000) ||
      !validDate(comment.createdAt) ||
      typeof comment.resolved !== "boolean"
    )
      fail("검토 메모가 올바르지 않습니다.");
    ids.add(comment.id);
    for (const key of ['author','initials','html']) if (comment[key] !== undefined && !boundedString(comment[key], key === 'html' ? 1000000 : 512)) fail('검토 메모 정보가 올바르지 않습니다.');
  }
  let nodes = 0;
  const walk = (value: unknown, parent: string | null, depth: number, framePage?: PageSettings): void => {
    if (++nodes > 100000 || depth > 64) fail("문서 구조가 너무 큽니다.");
    if (!object(value) || typeof value.type !== "string")
      fail("문서 본문 구조가 올바르지 않습니다.");
    keysOnly(value, ["type", "attrs", "content", "marks", "text"], "본문 요소");
    const type = value.type;
    let childFrame: PageSettings | undefined;
    if (parent === null ? type !== "doc" : type === "doc")
      fail("문서 최상위 구조가 올바르지 않습니다.");
    if (
      type !== "doc" &&
      !BLOCKS.has(type) &&
      !INLINE.has(type) &&
      !["listItem", "tableRow", "tableCell", "tableHeader"].includes(type)
    )
      fail(`지원하지 않는 본문 요소: ${type}`);
    if (
      parent === "paragraph" ||
      parent === "heading" ||
      parent === "codeBlock"
    ) {
      if (!INLINE.has(type)) fail("문단 안에 잘못된 요소가 있습니다.");
    } else if (parent === "bulletList" || parent === "orderedList") {
      if (type !== "listItem") fail("목록 구조가 올바르지 않습니다.");
    } else if (parent === "table") {
      if (type !== "tableRow") fail("표 행 구조가 올바르지 않습니다.");
    } else if (parent === "tableRow") {
      if (!["tableCell", "tableHeader"].includes(type))
        fail("표 셀 구조가 올바르지 않습니다.");
    } else if (parent !== null && !BLOCKS.has(type))
      fail("잘못된 본문 요소 위치입니다.");
    if(type==='documentDrawingAnchor' && (!object(value.attrs) || parseDrawingGroup(value.attrs.frame)?.position!=='inline'))fail('그림 그룹의 인라인 크기가 올바르지 않습니다.');
    if (type === 'documentTextBox' && (!object(value.attrs) || !textBoxFrame(value.attrs.frame))) fail('텍스트상자 설정이 올바르지 않습니다.');
    if (value.attrs !== undefined) {
      if (!object(value.attrs)) fail("본문 속성이 올바르지 않습니다.");
      for (const [key, attr] of Object.entries(value.attrs)) {
        if (
          ["__proto__", "constructor", "prototype"].includes(key) ||
          key.length > 100
        )
          fail("잘못된 본문 속성입니다.");
        if (
          attr !== null &&
          typeof attr !== "string" &&
          typeof attr !== "number" &&
          typeof attr !== "boolean" &&
          !(
            key === "colwidth" &&
            Array.isArray(attr) &&
            attr.length <= 1000 &&
            attr.every((n) => finite(n, 0, 10000))
          )
        )
          fail("본문 속성 형식이 올바르지 않습니다.");
        if (typeof attr === "string" && key !== "src" && key !== "hwpCellRasterFill" && attr.length > ((type === 'sectionBreak' || type === 'sectionContent') && key === 'page' ? MAX_FILE_BYTES : 10000))
          fail("본문 속성이 너무 깁니다.");
        if (typeof attr === "number" && !Number.isFinite(attr))
          fail("본문 수치 속성이 올바르지 않습니다.");
      }
    }
    const attrs = value.attrs as Record<string, unknown> | undefined;
    for (const key of ['wordParagraphStyle', 'wordParagraphStyleName'])
      if (attrs?.[key] != null && (!boundedString(attrs[key], 1000) || !attrs[key] || /[\u0000-\u001f]/.test(String(attrs[key])))) fail('문단 스타일 정보가 올바르지 않습니다.');
    if (attrs?.wordOutlineLevel != null && (!Number.isInteger(attrs.wordOutlineLevel) || Number(attrs.wordOutlineLevel) < 0 || Number(attrs.wordOutlineLevel) > 9)) fail('문단 개요 수준이 올바르지 않습니다.');
    if (attrs?.wordParagraphStyleKind != null && !/^(?:paragraph|heading[1-6])$/.test(String(attrs.wordParagraphStyleKind))) fail('문단 스타일 종류가 올바르지 않습니다.');
    for (const key of ['wordKeepNext', 'wordKeepLines']) if (attrs?.[key] != null && typeof attrs[key] !== 'boolean') fail('문단 묶음 정책이 올바르지 않습니다.');
    if (attrs?.wordContextualSpacing != null && typeof attrs.wordContextualSpacing !== 'boolean') fail('문단 간격 정책이 올바르지 않습니다.');
    if (attrs?.wordParagraphBottomBorder != null &&
        (!['paragraph', 'heading'].includes(type) || !parseWordParagraphBottomBorder(attrs.wordParagraphBottomBorder)))
      fail('문단 하단 테두리 정보가 올바르지 않습니다.');
    if (attrs?.paragraphBorderBox != null &&
        (!['paragraph', 'heading'].includes(type) || !parseParagraphBorderBox(attrs.paragraphBorderBox)))
      fail('문단 테두리 정보가 올바르지 않습니다.');
    if (attrs?.wordParagraphReflection != null &&
        (!['paragraph', 'heading'].includes(type) || !parseWordReflection(attrs.wordParagraphReflection)))
      fail('문단 반사 효과 정보가 올바르지 않습니다.');
    if (attrs?.wordTextAlignment != null && !paragraphTextAlignment(attrs.wordTextAlignment)) fail('문단 글줄 정렬이 올바르지 않습니다.');
    if (attrs?.hwpCenteredCellLabels != null && (type !== 'tableCell' || !parseHwpCenteredCellLabels(attrs.hwpCenteredCellLabels)))
      fail('한글 표 셀의 배치 정보가 올바르지 않습니다.');
    if (attrs?.hwpCellRasterFill != null && (type !== 'tableCell' || !parseHwpCellRasterFill(attrs.hwpCellRasterFill)))
      fail('한글 표 셀의 그림 채우기 정보가 올바르지 않습니다.');
    if (attrs?.hwpCellLine != null && !parseHwpCellLine(attrs.hwpCellLine)) fail('한글 셀의 줄 간격 정보가 올바르지 않습니다.');
    if (type === 'documentSymbol' && (!boundedString(attrs?.font, 100) || !attrs?.font || !/^\d{1,5}$/.test(String(attrs?.code)) || !finite(Number(attrs?.code), 0, 65535) || !boundedString(attrs?.glyph, 100))) fail('기호 글자가 올바르지 않습니다.');
    if (type === 'documentBookmark' && (!boundedString(attrs?.name, 1000) || !attrs?.name)) fail('책갈피 이름이 올바르지 않습니다.');
    if (type === 'documentField' && !['PAGE','NUMPAGES'].includes(String(attrs?.kind))) fail('페이지 필드가 올바르지 않습니다.');
    if ((type === 'documentNote' || type === 'noteReference') && (!boundedString(attrs?.id, 128) || !attrs?.id || !['footnote','endnote'].includes(String(attrs?.kind)) || !boundedString(attrs?.label, 100))) fail('각주·미주 정보가 올바르지 않습니다.');
    if (attrs?.hwpFootnoteLayout != null && (type !== 'documentNote' || attrs.kind !== 'footnote' ||
        !parseHwpFootnoteLayout(attrs.hwpFootnoteLayout))) fail('한글 각주 조판 정보가 올바르지 않습니다.');
    if (attrs?.hwpFootnoteMark != null && (type !== 'noteReference' || attrs.kind !== 'footnote' ||
        !hwpFootnoteMark(attrs.hwpFootnoteMark))) fail('한글 각주 표시가 올바르지 않습니다.');
    if (type === 'columnGroup' && (!finite(attrs?.count, 1, 12) || !Number.isInteger(attrs?.count) || !finite(attrs?.space, 0, 1000) || (attrs?.balance !== undefined && typeof attrs.balance !== 'boolean'))) fail('다단 본문 설정이 올바르지 않습니다.');
    if (type === 'wordOverlap' && !overlapParts(attrs?.parts).length) fail('겹친 글자 형식이 올바르지 않습니다.');
    if (type === 'hwpBoxedDigit' && !parseHwpBoxedDigit(attrs?.box)) fail('네모 숫자 정보가 올바르지 않습니다.');
    if (attrs?.cellTablePosition != null && (type !== 'table' || !cellTablePosition(attrs.cellTablePosition))) fail('셀 안의 표 위치가 올바르지 않습니다.');
    if (attrs?.hwpKeptTablePage != null && (type !== 'table' || typeof attrs.hwpKeptTablePage !== 'number' || !hwpKeptTablePage(attrs.hwpKeptTablePage)))
      fail('한글 인라인 표의 쪽 연속 정보가 올바르지 않습니다.');
    if (attrs?.hwpNativeGrid != null && (type !== 'table' || !parseHwpNativeGrid(attrs.hwpNativeGrid)))
      fail('한글 표의 원본 열 너비가 올바르지 않습니다.');
    if (attrs?.hwpInlineTableTop != null && (type !== 'table' || typeof attrs.hwpInlineTableTop !== 'number' ||
        hwpInlineTableTop(attrs.hwpInlineTableTop) === undefined)) fail('한글 인라인 표의 위쪽 바깥 여백이 올바르지 않습니다.');
    if (attrs?.hwpInlineTableBottom != null && (type !== 'table' || typeof attrs.hwpInlineTableBottom !== 'number' ||
        hwpInlineTableBottom(attrs.hwpInlineTableBottom) === undefined)) fail('한글 인라인 표의 아래쪽 바깥 여백이 올바르지 않습니다.');
    if (attrs?.hwpNestedTableLayout != null && (type !== 'table' || !parseHwpNestedTableLayout(attrs.hwpNestedTableLayout)))
      fail('한글 중첩 표의 조판 정보가 올바르지 않습니다.');
    if (attrs?.hwpRelatedTableLayout != null && (type !== 'table' || !parseHwpRelatedTableLayout(attrs.hwpRelatedTableLayout)))
      fail('한글 표와 인라인 표의 연속 조판 정보가 올바르지 않습니다.');
    if (attrs?.hwpTableContinuation != null && (type !== 'table' || !parseHwpTableContinuation(attrs.hwpTableContinuation)))
      fail('한글 표의 이어지는 조각 정보가 올바르지 않습니다.');
    if (type === 'table' && attrs?.hwpCenteredBorders != null && typeof attrs.hwpCenteredBorders !== 'boolean' &&
      attrs.hwpCenteredBorders !== 'inline' && attrs.hwpCenteredBorders !== 'continuation')
      fail('한글 표 테두리의 배치 출처가 올바르지 않습니다.');
    if (type === 'tableRow' && attrs?.rowHeight != null && !rowHeight(attrs.rowHeight))
      fail(`표 행 높이는 0보다 크고 ${MAX_ROW_HEIGHT}px 이하여야 합니다.`);
    if (type === 'tableRow' && attrs?.hwpGridRow != null && typeof attrs.hwpGridRow !== 'boolean')
      fail('한글 표의 행 높이 출처가 올바르지 않습니다.');
    if (type === 'tableRow' && attrs?.hwpRowContinued != null && typeof attrs.hwpRowContinued !== 'boolean')
      fail('한글 표 행의 쪽 나누기 출처가 올바르지 않습니다.');
    if (type === 'tableRow' && attrs?.hwpRowIntact != null && typeof attrs.hwpRowIntact !== 'boolean')
      fail('한글 표 행의 배치 출처가 올바르지 않습니다.');
    if (type === 'tableRow' && attrs?.hwpIntactTopRow != null && !parseHwpIntactRow(attrs.hwpIntactTopRow))
      fail('한글 표 행의 배치 정보가 올바르지 않습니다.');
    if (type === 'tableRow' && attrs?.hwpRowReserve != null && !hwpRowReserve(attrs.hwpRowReserve))
      fail('한글 표 행의 분할 높이가 올바르지 않습니다.');
    if (type === 'tableRow' && [attrs?.rowSplit, attrs?.wordRowSplit].some(value => value != null && !['allow', 'avoid'].includes(String(value))))
      fail('표 행의 쪽 나누기 설정이 올바르지 않습니다.');
    if (type === 'sectionBreak' || type === 'sectionContent') {
      if (parent !== 'doc' && !(type === 'sectionBreak' && parent === 'sectionContent' && framePage)) fail('구역 나누기는 문서의 최상위에 있어야 합니다.');
      if (typeof attrs?.page !== 'string' || attrs.page.length > MAX_FILE_BYTES) fail('구역 설정이 올바르지 않습니다.');
      const sectionPage = JSON.parse(attrs.page);
      if (parent === 'sectionContent' && (sectionPage.breakType !== 'continuous' || !sameSectionFrameSettings(framePage!, sectionPage))) fail('공유 구역에는 설정이 같은 이어지는 구역만 넣을 수 있습니다.');
      if (type === 'sectionContent') childFrame = sectionPage;
      validateDocument({ format: 'word-studio', version: 1, title: '', updatedAt: new Date().toISOString(), page: sectionPage,
        comments: [], content: { type: 'doc', content: [{ type: 'paragraph' }] } });
    }
    if ((type === "image" || type === "inlineImage") && !validImage(attrs?.src))
      fail("문서 그림은 PNG, JPG, GIF, BMP 또는 WebP 데이터여야 합니다.");
    if ((type === 'image' || type === 'inlineImage') && attrs?.floating != null && !imageFloating(attrs.floating)) fail('그림 둘러싸기 설정이 올바르지 않습니다.');
    if (
      type === "heading" &&
      (!finite(attrs?.level, 1, 6) || !Number.isInteger(attrs?.level))
    )
      fail("제목 수준이 올바르지 않습니다.");
    if (["tableCell", "tableHeader"].includes(type)) {
      if (attrs?.cellSplit != null && attrs.cellSplit !== "allow") fail("병합 셀의 쪽 나누기 설정이 올바르지 않습니다.");
      for (const key of ["colspan", "rowspan"])
        if (
          attrs?.[key] != null &&
          (!finite(attrs[key], 1, 1000) || !Number.isInteger(attrs[key]))
        )
          fail("표 셀 크기가 올바르지 않습니다.");
    }
    if (type === "text") {
      if (!boundedString(value.text, 4000000) || !value.text)
        fail("텍스트 요소가 올바르지 않습니다.");
    } else if (value.text !== undefined)
      fail("텍스트가 아닌 요소에 텍스트 속성이 있습니다.");
    if (value.marks !== undefined) {
      if (
        !Array.isArray(value.marks) ||
        value.marks.length > 20 ||
        !INLINE.has(type)
      )
        fail("텍스트 서식이 올바르지 않습니다.");
      for (const mark of value.marks) {
        if (!object(mark) || !MARKS.has(String(mark.type)))
          fail("지원하지 않는 텍스트 서식입니다.");
        keysOnly(mark, ["type", "attrs"], "텍스트 서식");
        if (
          mark.attrs !== undefined &&
          (!object(mark.attrs) ||
            Object.entries(mark.attrs).some(
              ([key, val]) =>
                ["__proto__", "constructor", "prototype"].includes(key) ||
                (val !== null &&
                  typeof val !== "string" &&
                  typeof val !== "number" &&
                  typeof val !== "boolean") ||
                (typeof val === "string" && val.length > 10000),
            ))
        )
          fail("텍스트 서식 속성이 올바르지 않습니다.");
        if (mark.type === 'wordReflection' && (type !== 'text' || !object(mark.attrs) || !parseWordReflection(mark.attrs.effect)))
          fail('글자 반사 효과 정보가 올바르지 않습니다.');
        if (
          mark.type === "link" &&
          (!object(mark.attrs) || !safeLink(mark.attrs.href))
        )
          fail("지원하지 않는 링크 주소입니다.");
      }
    }
    const leaf = [
      "text",
      "hardBreak",
      "horizontalRule",
      "pageBreak",
      "columnBreak",
      "sectionBreak",
      "image",
      "inlineImage",
      "documentTab",
      "wordOverlap",
      "hwpBoxedDigit",
      "documentBookmark",
      "noteReference", "documentField",
      "documentSymbol",
    ].includes(type);
    if (value.content !== undefined && (!Array.isArray(value.content) || leaf))
      fail("본문 자식 요소가 올바르지 않습니다.");
    const children = (value.content ?? []) as unknown[];
    if (
      [
        "doc",
        "bulletList",
        "orderedList",
        "listItem",
        "table",
        "tableCell",
        "tableHeader",
        "blockquote",
        "documentNote", "columnGroup", "sectionContent", "documentTextBox",
      ].includes(type) &&
      children.length === 0
    )
      fail("필수 본문 요소가 비어 있습니다.");
    if (
      type === "listItem" &&
      (!object(children[0]) || children[0].type !== "paragraph")
    )
      fail("목록 항목에는 첫 문단이 필요합니다.");
    children.forEach((child) => walk(child, type, depth + 1, childFrame));
    if (type === 'table') {
      // Empty rows are valid only when covered completely by earlier rowspans.
      // Validate the grid rather than rejecting all empty rows after a merge.
      const remaining: number[] = [];
      let width = 0;
      for (const row of children as JSONContent[]) {
        let column = 0;
        for (const cell of row.content ?? []) {
          while ((remaining[column] ?? 0) > 0) column++;
          const span = Number(cell.attrs?.colspan) || 1;
          if (column + span > 1000) fail('표의 열 수가 너무 많습니다.');
          for (let i = 0; i < span; i++) {
            if ((remaining[column + i] ?? 0) > 0) fail('표의 병합 셀이 겹칩니다.');
            remaining[column + i] = Number(cell.attrs?.rowspan) || 1;
          }
          column += span;
        }
        width ||= remaining.length;
        if (!width || remaining.length !== width) fail('표 행의 열 수가 일치하지 않습니다.');
        for (let i = 0; i < width; i++) {
          if (!(remaining[i] > 0)) fail('표 행에 채워지지 않은 셀이 있습니다.');
          remaining[i]--;
        }
      }
      if (remaining.some(value => value > 0)) fail('병합 셀이 표의 마지막 행을 벗어납니다.');
    }
  };
  walk(input.content, null, 0);
  return input as unknown as StudioDocument;
}

/** Removes active markup, external images and CSS capable of loading external resources. */
function makeDocument(name: string, content?: JSONContent): StudioDocument {
  return {
    id: crypto.randomUUID(),
    format: "word-studio",
    version: 1,
    title: name.replace(/\.[^.]+$/, "").slice(0, 512) || "제목 없는 문서",
    content: content ?? { type: "doc", content: [{ type: "paragraph" }] },
    page: { ...defaultPage },
    comments: [],
    updatedAt: new Date().toISOString(),
  };
}
function xml(source: string): XMLDocument {
  if (/<!DOCTYPE|<!ENTITY/i.test(source))
    fail("지원하지 않는 XML 선언이 있습니다.");
  const parsed = new DOMParser().parseFromString(source, "application/xml");
  if (parsed.getElementsByTagName("parsererror").length)
    fail("DOCX 내부 XML을 읽을 수 없습니다.");
  return parsed;
}
const elements = (node: Element | XMLDocument, name: string) =>
  Array.from(node.getElementsByTagNameNS("*", name));
const attr = (node: Element | undefined, name: string) =>
  node?.getAttributeNS(
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    name,
  ) ??
  node?.getAttribute(`w:${name}`) ??
  null;
function textOf(node: Element | XMLDocument): string {
  return elements(node, "p")
    .map((p) => {
      const copy = p.cloneNode(true) as Element;
      for (const field of elements(copy, "fldSimple"))
        if (/\b(?:PAGE|NUMPAGES)\b/i.test(attr(field, "instr") || ""))
          field.remove();
      let skipPageField = false;
      let result = "";
      for (const run of elements(copy, "r")) {
        if (
          elements(run, "instrText").some((instruction) =>
            /\b(?:PAGE|NUMPAGES)\b/i.test(instruction.textContent || ""),
          )
        )
          skipPageField = true;
        if (!skipPageField)
          result += elements(run, "*")
            .map(n => n.localName === 't' ? n.textContent || '' : n.localName === 'tab' ? '\t' : n.localName === 'br' ? '\n' : '')
            .join("");
        if (
          elements(run, "fldChar").some(
            (field) => attr(field, "fldCharType") === "end",
          )
        )
          skipPageField = false;
      }
      return result;
    })
    .filter(Boolean)
    .join("\n");
}

async function readDocxPage(
  zip: JSZip,
  body: XMLDocument,
  warnings: string[],
  selectedSection?: Element,
): Promise<PageSettings> {
  const page: PageSettings = { ...defaultPage, pageNumbers: false };
  const sections = elements(body, "sectPr");
  const section = selectedSection ?? sections[0];
  if (!section) return page;
  const breakType = attr(elements(section, 'type')[0], 'val') || 'nextPage';
  if (['nextPage','continuous','oddPage','evenPage','nextColumn'].includes(breakType)) page.breakType = breakType as PageSettings['breakType'];
  const columns = elements(section, 'cols')[0];
  if (columns) page.columns = { count: Math.max(1, Math.min(12, Number(attr(columns, 'num')) || elements(columns, 'col').length || 1)),
    space: Math.max(0, Math.min(1000, Number(attr(columns, 'space') ?? 720) / 15)) };
  const size = elements(section, "pgSz")[0];
  const width = Number(attr(size, "w")),
    height = Number(attr(size, "h"));
  page.orientation =
    attr(size, "orient") === "landscape" || (width > height && height > 0)
      ? "landscape"
      : "portrait";
  const smaller = Math.min(width, height),
    larger = Math.max(width, height);
  if (Math.abs(smaller - 12240) < 150 && Math.abs(larger - 15840) < 150)
    page.size = "Letter";
  else if (
    width &&
    (Math.abs(smaller - 11906) > 150 || Math.abs(larger - 16838) > 150)
  )
  {
    page.size = 'Custom';
    page.width = smaller * 25.4 / 1440;
    page.height = larger * 25.4 / 1440;
    // Preserve the physical output size separately. Sub-twip differences
    // can alter a line break even when the document uses the same paper.
    const precise = exactPaperDimensions(size, smaller, larger);
    if (precise) page.docxPrintPaper = precise;
  }
  const margin = elements(section, "pgMar")[0];
  const enabled = (el: Element | undefined) => !!el && !['0', 'false', 'off'].includes(attr(el, 'val') || '');
  page.differentFirstPage = enabled(elements(section, 'titlePg')[0]);
  const settingsFile = zip.file('word/settings.xml');
  const settings = settingsFile ? xml(await settingsFile.async('string')) : undefined;
  if (page.columns && settings) page.columns.balance = !enabled(elements(settings, 'noColumnBalance')[0]);
  // This compatibility flag is off when absent. Do not give an imported
  // Word document the native HWP final-line fit policy by default.
  page.suppressBottomSpacing = !!settings && enabled(elements(settings, 'suppressBottomSpacing')[0]);
  page.differentOddEven = !!settings && enabled(elements(settings, 'evenAndOddHeaders')[0]);
  page.gutterTop = !!settings && enabled(elements(settings, 'gutterAtTop')[0]);
  for (const [source, destination] of [['header', 'headerDistance'], ['footer', 'footerDistance'], ['gutter', 'gutter']] as const) {
    const value = attr(margin, source);
    if (value !== null && Number(value) >= 0 && Number(value) <= 4252) page[destination] = Number(value) * 25.4 / 1440;
  }
  const margins = ["top", "right", "bottom", "left"].map((key) =>
    Number(attr(margin, key)),
  );
  if (
    margin &&
    margins.every((n) => Number.isFinite(n) && n >= 0 && n <= 4252)
  ) {
    page.margin = Math.round((margins[0] / (1440 / 25.4)) * 10) / 10;
    if (margins.some((n) => n !== margins[0]))
      page.margins = { top: margins[0] * 25.4 / 1440, right: margins[1] * 25.4 / 1440, bottom: margins[2] * 25.4 / 1440, left: margins[3] * 25.4 / 1440 };
  }
  const nativeMargins = parseHwpHorizontalMargins(margin?.getAttributeNS(exactPaperNamespace, 'hwpHorizontal'));
  if (nativeMargins) {
    const precise = parseHwpHorizontalMargins(nativeMargins.map((original, index) => {
      const twips = attr(margin, index === 0 ? 'left' : 'right');
      return original !== null && twips !== null && Math.round(original / 5) === Number(twips) ? original : null;
    }));
    if (precise) page.hwpHorizontalMargins = precise;
  }
  const relFile = zip.file("word/_rels/document.xml.rels");
  if (!relFile) return page;
  const rels = xml(await relFile.async("string"));
  for (const kind of ["header", "footer"] as const) {
    // Each header/footer type is linked independently. Missing references
    // inherit; an explicit reference to an empty part breaks that link.
    const inherited = new Map<string, Element>();
    for (const prior of sections.slice(0, sections.indexOf(section) + 1)) {
      for (const ref of elements(prior, `${kind}Reference`))
        inherited.set(attr(ref, 'type') || 'default', ref);
    }
    for (const type of ['default', 'first', 'even'] as const) {
    const ref = inherited.get(type);
    if (!ref) continue;
    const id =
      ref.getAttributeNS(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "id",
      ) ?? ref.getAttribute("r:id");
    const rel = elements(rels, "Relationship").find(
      (r) =>
        r.getAttribute("Id") === id &&
        r.getAttribute("TargetMode") !== "External",
    );
    const target = rel?.getAttribute("Target");
    if (!target || target.includes("..") || /[\\:]/.test(target)) continue;
    const part = zip.file(
      target.startsWith("/") ? target.slice(1) : `word/${target}`,
    );
    if (!part) continue;
    const parsed = xml(await part.async("string"));
    const key = type === 'default' ? kind : `${type}${kind === 'header' ? 'Header' : 'Footer'}` as 'firstHeader' | 'firstFooter' | 'evenHeader' | 'evenFooter';
    page[key] = textOf(parsed).slice(0, 10000);
    page[`${key}Html` as 'headerHtml'] = sanitizeImportedHtml(await storyHtml(zip, part.name, parsed.documentElement, page)).html;
    if (
      elements(parsed, "instrText").some((n) =>
        /\bPAGE\b/i.test(n.textContent || ""),
      ) ||
      elements(parsed, "fldSimple").some((n) =>
        /\bPAGE\b/i.test(attr(n, "instr") || ""),
      )
    )
      page[type === 'default' ? 'pageNumbers' : type === 'first' ? 'firstPageNumbers' : 'evenPageNumbers'] = true;
    }
  }
  return page;
}

export async function importFile(file: File): Promise<ImportResult> {
  if (!file.size) fail("비어 있는 파일입니다.");
  const extension = file.name.split(".").pop()?.toLowerCase();
  // Base64 pictures expand when a compressed HWP/DOCX becomes editable JSON.
  const limit = documentFileLimit(extension);
  if (file.size > limit) fail(`${extension === 'wordstudio' ? 'Word Studio 편집 파일은' : '문서는'} ${limit / 1024 / 1024} MiB 이하로 열어 주세요.`);
  if (extension === "wordstudio") {
    let parsed: unknown;
    try {
      parsed = JSON.parse(await file.text());
    } catch {
      fail("Word Studio 파일의 JSON을 읽을 수 없습니다.");
    }
    return { document: validateDocument(parsed), warnings: [] };
  }
  if (extension === "txt") {
    const text = (await file.text()).replace(/^\uFEFF/, "");
    return {
      document: makeDocument(file.name, {
        type: "doc",
        content: text
          .split(/\r\n|\n|\r/)
          .map((line) => ({
            type: "paragraph",
            ...(line ? { content: [{ type: "text", text: line }] } : {}),
          })),
      }),
      warnings: [],
    };
  }
  if (extension === "html" || extension === "htm") {
    const sanitized = sanitizeImportedHtml(await file.text());
    return { document: makeDocument(file.name), ...sanitized };
  }
  if (extension !== "docx")
    fail(
      "DOCX, Word Studio, HTML 또는 TXT 파일을 선택해 주세요. 구형 DOC 파일은 DOCX로 변환해야 합니다.",
    );
  const buffer = await file.arrayBuffer();
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(buffer);
  } catch {
    fail(
      "DOCX 파일을 읽을 수 없습니다. 암호화 또는 손상된 파일인지 확인해 주세요.",
    );
  }
  const entries = Object.values(zip.files);
  if (entries.length > 10000) fail("DOCX에 내부 파일이 너무 많습니다.");
  // JSZip exposes central-directory size metadata before inflating file contents.
  const inflatedSize = entries.reduce(
    (sum, entry) =>
      sum +
      ((entry as unknown as { _data?: { uncompressedSize?: number } })._data
        ?.uncompressedSize ?? 0),
    0,
  );
  if (inflatedSize > MAX_INFLATED_BYTES)
    fail(`압축을 푼 DOCX 크기가 ${MAX_INFLATED_BYTES / 1024 / 1024} MiB를 초과합니다.`);
  const documentXml = zip.file("word/document.xml");
  if (!documentXml || !zip.file("[Content_Types].xml"))
    fail("유효한 DOCX 문서가 아닙니다.");
  const body = xml(await documentXml.async("string"));
  const warnings = [
    "DOCX는 편집 가능한 본문으로 변환됩니다. 원본의 세밀한 글꼴·배치·페이지 나눔은 달라질 수 있습니다.",
  ];
  if (elements(body, "drawing").some((n) => elements(n, "anchor").length))
    warnings.push("떠 있는 그림의 위치·양쪽 텍스트 흐름·원본 다각형 윤곽을 가져옵니다.");
  if (elements(body, "ins").length || elements(body, "del").length)
    warnings.push(
      "변경 내용 추적은 최종 본문으로 가져옵니다. 검토 이력은 보존하지 않습니다.",
    );
  if (
    elements(body, "oMath").length ||
    elements(body, "object").length ||
    elements(body, "pict").length
  )
    warnings.push(
      "수식, OLE 개체 또는 도형은 완전히 변환되지 않을 수 있습니다.",
    );
  const doc = makeDocument(file.name);
  const repairedTables = normalizeDocxTables(body);
  if (repairedTables) warnings.push(`생략된 표 격자 또는 가로 병합 ${repairedTables}개를 편집 가능한 셀로 보완했습니다.`);
  const sections = elements(body, 'sectPr');
  const pages = await Promise.all(sections.map(s => readDocxPage(zip, body, warnings, s)));
  // DOCX has no equivalent of Hancom's conditional two-line section rule.
  // Keep just that policy in a standard custom property, without replacing
  // real paragraphs or freezing their heights to the current pagination.
  let defaultHangingSpaces = false;
  const customPart = zip.file('docProps/custom.xml');
  if (customPart) {
    const raw = await customPart.async('string');
    if (raw.length <= 65536) {
      const properties = elements(xml(raw), 'property');
      const hangingSpaces = properties.find(p => p.getAttribute('name') === wordDefaultHangingSpacesProperty &&
        p.namespaceURI === 'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties')?.firstElementChild;
      defaultHangingSpaces = hangingSpaces?.namespaceURI === 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes' &&
        hangingSpaces.localName === 'lpwstr' && hangingSpaces.textContent === 'hang';
      const names = properties.find(p => p.getAttribute('name') === bookmarkNamesProperty);
      if (names) try {
        const pairs = JSON.parse(names.textContent || 'null');
        if (Array.isArray(pairs) && pairs.length <= 2048 && pairs.every(p => Array.isArray(p) && p.length === 2 &&
          typeof p[0] === 'string' && p[0].length > 0 && p[0].length <= 40 && typeof p[1] === 'string' && p[1].length > 0 && p[1].length <= 1000)) {
          const restored = new Map<string, string>(pairs);
          for (const element of [...elements(body, 'bookmarkStart'), ...elements(body, 'hyperlink')]) {
            const attribute = element.localName === 'bookmarkStart' ? 'name' : 'anchor';
            const current = attr(element, attribute) || '', original = restored.get(current);
            if (original) element.setAttributeNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', `w:${attribute}`, original);
          }
        }
      } catch { /* Optional original names must not prevent opening the document. */ }
      const property = properties.find(p => p.getAttribute('name') === emptyLineDocxProperty);
      if (property) try {
        const indexes = JSON.parse(property.textContent || 'null');
        if (Array.isArray(indexes) && indexes.length <= pages.length && indexes.every(i => Number.isInteger(i) && i >= 0 && i < pages.length)) {
          for (const index of indexes) pages[index].hideEmptyLines = true;
          if (indexes.length) warnings.push(emptyLineCompatibilityWarning);
        }
      } catch { /* A malformed optional policy must not damage the body. */ }
    }
  }
  doc.page = pages[0] ?? { ...defaultPage };
  const settingsPart = zip.file('word/settings.xml');
  const settings = settingsPart ? xml(await settingsPart.async('string')) : undefined;
  const settingsRoot = settings?.documentElement;
  if (settingsRoot?.namespaceURI === W && settingsRoot.localName === 'settings') {
    const wordChild = (parent: Element | undefined, name: string) => parent && Array.from(parent.children)
      .find(el => el.namespaceURI === W && el.localName === name);
    const readOption = (element?: Element) => {
      if (!element) return;
      const value = element.getAttributeNS(W, 'val');
      if (value === null || value === 'true' || value === '1' || value === 'on') return true;
      if (value === 'false' || value === '0' || value === 'off') return false;
    };
    const autoHyphenation = readOption(wordChild(settingsRoot, 'autoHyphenation'));
    if (autoHyphenation !== undefined) doc.wordAutoHyphenation = autoHyphenation;
    const wrapTrailSpaces = readOption(wordChild(wordChild(settingsRoot, 'compat'), 'wrapTrailSpaces'));
    if (wrapTrailSpaces !== undefined) doc.wordWrapTrailSpaces = wrapTrailSpaces;
  }
  if (doc.wordWrapTrailSpaces === undefined) {
    // Older Studio exports of HWP and Word share the same package signature.
    // Its existing whitespace policy stays in effect unless the file carries
    // the Word default, with explicit Word settings taking precedence.
    const corePart = zip.file('docProps/core.xml');
    const coreText = await corePart?.async('string');
    const core = coreText && coreText.length <= 65536 ? xml(coreText).documentElement : undefined;
    const dc = 'http://purl.org/dc/elements/1.1/';
    const legacyStudio = core?.getElementsByTagNameNS(dc, 'creator')[0]?.textContent === wordStudioCreator &&
      core.getElementsByTagNameNS(dc, 'description')[0]?.textContent === wordStudioDescription;
    if (defaultHangingSpaces || !legacyStudio) doc.wordDefaultHangingSpaces = true;
  }
  const background = readWordBackground(body, settings, warnings);
  if (background) for (const page of new Set([doc.page, ...pages])) page.wordBackground = background;
  const checkboxCount = await countDocxCheckboxes(zip, body);
  if (checkboxCount) warnings.push(`체크 상자 ${checkboxCount}개를 선택 상태가 유지되는 기호로 가져왔습니다. Word 양식 컨트롤은 기호로 변환됩니다.`);
  const textboxes = await prepareTextBoxes(zip, body, pages);
  if (textboxes.count) warnings.push(`텍스트상자 ${textboxes.count}개의 내용을 배경·크기와 함께 편집 가능한 별도 영역으로 가져왔습니다.`);
  const metadata = await prepareDocxMetadata(zip, body, pages, { connectedBodyBorders: true });
  doc.comments = metadata.comments;
  if (metadata.wordDefaultLanguage) doc.wordDefaultLanguage = metadata.wordDefaultLanguage;
  const commentsPart = zip.file('word/comments.xml');
  if (commentsPart) {
    const raw = elements(xml(await commentsPart.async('string')), 'comment');
    for (let i = 0; i < doc.comments.length; i++) if (raw[i])
      doc.comments[i].html = sanitizeImportedHtml(await storyHtml(zip, 'word/comments.xml', raw[i], doc.page)).html;
  }
  const result = await mammoth.convertToHtml(
    { arrayBuffer: metadata.buffer },
    {
      externalFileAccess: false,
      includeEmbeddedStyleMap: false,
      ignoreEmptyParagraphs: false,
      styleMap: [
        "u => u",
        "strike => s",
        "p[style-name='Title'] => h1:fresh",
        "p[style-name='Subtitle'] => h2:fresh",
      ],
    },
  );
  const notes = await documentNotes(zip, doc.page);
  const sanitized = sanitizeImportedHtml(structureDocxSections(attachDocumentNotes(textboxes.decorate(metadata.decorate(result.value)), notes), pages));
  return {
    document: doc,
    html: sanitized.html,
    warnings: [
      ...new Set([
        ...warnings,
        ...sanitized.warnings,
        ...result.messages.map((message) => message.message),
      ]),
    ],
  };
}

function rgb(value: unknown): string | undefined {
  if (typeof value !== "string") return undefined;
  const hex = value.replace("#", "");
  if (/^[0-9a-f]{6}$/i.test(hex)) return hex.toUpperCase();
  if (/^[0-9a-f]{3}$/i.test(hex))
    return hex
      .split("")
      .map((c) => c + c)
      .join("")
      .toUpperCase();
  const match = value.match(/^rgba?\(\s*(\d+)\s*[, ]\s*(\d+)\s*[, ]\s*(\d+)/i);
  if (match)
    return match
      .slice(1, 4)
      .map((n) => Math.min(255, Number(n)).toString(16).padStart(2, "0"))
      .join("")
      .toUpperCase();
  const names: Record<string, string> = {
    black: "000000",
    white: "FFFFFF",
    yellow: "FFFF00",
    red: "FF0000",
    blue: "0000FF",
    green: "008000",
    gray: "808080",
  };
  return names[value.toLowerCase()];
}
function cellBorders(attrs: Record<string, unknown> = {}): ITableCellBorders {
  return Object.fromEntries(['top', 'right', 'bottom', 'left'].map(side => {
    const value = attrs[`border${side[0].toUpperCase()}${side.slice(1)}`];
    const css = document.createElement('span').style;
    css.border = typeof value === 'string' ? value : '1px solid #c6ced8';
    const styles = { solid: BorderStyle.SINGLE, dashed: BorderStyle.DASHED, dotted: BorderStyle.DOTTED, double: BorderStyle.DOUBLE };
    return [side, {
      style: ['none', 'hidden'].includes(css.borderStyle) ? BorderStyle.NIL : styles[css.borderStyle as keyof typeof styles] ?? BorderStyle.SINGLE,
      size: Math.max(0, Math.round((cssPoints(css.borderWidth) ?? .75) * 8 / (css.borderStyle === 'double' ? 3 : 1))), color: rgb(css.borderColor) ?? 'C6CED8',
    }];
  }));
}

function inheritCellText(nodes: JSONContent[], attrs: Record<string, unknown> = {}): JSONContent[] {
  const defaults = { fontFamily: attrs.fontFamily, fontSize: attrs.fontSize, color: attrs.textColor };
  if (!Object.values(defaults).some(Boolean)) return nodes;
  return nodes.map(node => {
    if (node.type === 'text') {
      const existing = node.marks?.find(mark => mark.type === 'textStyle')?.attrs ?? {};
      return { ...node, marks: [...(node.marks ?? []).filter(mark => mark.type !== 'textStyle'),
        { type: 'textStyle', attrs: { ...defaults, ...Object.fromEntries(Object.entries(existing).filter(([, value]) => value != null)) } }] };
    }
    const inherited = ['paragraph', 'heading'].includes(node.type || '') ? { ...attrs,
      fontFamily: node.attrs?.paragraphFont || attrs.fontFamily,
      fontSize: node.attrs?.paragraphSize || attrs.fontSize,
      textColor: node.attrs?.paragraphColor || attrs.textColor } : attrs;
    return node.content ? { ...node, content: inheritCellText(node.content, inherited) } : node;
  });
}
const propertiesBeforeCharacterSpacing = new Set([
  'rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike',
  'outline', 'shadow', 'emboss', 'imprint', 'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color',
].map(name => `w:${name}`));

class DocumentTextRun extends TextRun {
  private readonly explicitZeroSpacing: boolean;

  constructor(options: IRunOptions, reflection?: unknown, private readonly latinFont?: string) {
    super(options);
    // docx 9.x omits numeric zero, which explicitly cancels inherited tracking.
    this.explicitZeroSpacing = options.characterSpacing === 0;
    if (this.explicitZeroSpacing) this.properties.push(new NumberValueElement('w:spacing', 0));
    const attributes = wordReflectionXmlAttributes(reflection);
    if (attributes) this.properties.push(new ImportedXmlComponent('w14:reflection', attributes));
  }

  prepForXml(context: Parameters<TextRun['prepForXml']>[0]) {
    const result = super.prepForXml(context);
    const children = result?.['w:r'];
    if (!Array.isArray(children)) return result;
    if (this.latinFont) children.unshift({ _attr: {
      'xmlns:wslf': hwpLatinFontNamespace, 'xmlns:mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
      'mc:Ignorable': 'wslf', 'mc:PreserveAttributes': 'wslf:face', 'wslf:face': this.latinFont,
    } });
    if (!this.explicitZeroSpacing) return result;
    const properties = children.find(child => child && typeof child === 'object' && 'w:rPr' in child)?.['w:rPr'];
    if (!Array.isArray(properties)) return result;
    const spacing = properties.findIndex(property => property && Object.hasOwn(property, 'w:spacing'));
    const position = properties.findIndex(property =>
      property && Object.keys(property).some(name => !propertiesBeforeCharacterSpacing.has(name)));
    // Match the library's nonzero spacing order, ahead of size and later properties.
    if (position >= 0 && position < spacing) properties.splice(position, 0, properties.splice(spacing, 1)[0]);
    return result;
  }
}

/** Word may discard a whitespace-only literal tab inside w:t on save.
 * Emit each authored tab as the standard run control, retaining run marks. */
function documentTextRun(options: IRunOptions, reflection?: unknown, latinFont?: string): TextRun {
  if (!options.text?.includes('\t')) return new DocumentTextRun(options, reflection, latinFont);
  const { text, ...properties } = options;
  return new DocumentTextRun({ ...properties, children: text.split('\t').flatMap((part, i) => [
    ...(i ? [new Tab()] : []), ...(part ? [part] : []),
  ]) }, reflection, latinFont);
}
function runOptions(node: JSONContent): IRunOptions {
  const options: { -readonly [K in keyof IRunOptions]: IRunOptions[K] } = {
    text: node.text ?? "",
  };
  for (const mark of node.marks ?? []) {
    switch (mark.type) {
      case "bold":
        options.bold = true;
        break;
      case "italic":
        options.italics = true;
        break;
      case "underline":
        options.underline = { type: UnderlineType.SINGLE };
        break;
      case "strike":
        options.strike = true;
        break;
      case "subscript":
        options.subScript = true;
        break;
      case "superscript":
        options.superScript = true;
        break;
      case "code":
        options.font = "Consolas";
        options.shading = { fill: "F1F3F5" };
        break;
      case "highlight":
        options.shading = { fill: rgb(mark.attrs?.color) ?? "FFF59D" };
        break;
      case "textStyle": {
        if (mark.attrs?.smallCaps) options.smallCaps = true;
        if (mark.attrs?.allCaps) options.allCaps = true;
        if (mark.attrs?.decorationStyle === 'double') {
          if (node.marks?.some(m => m.type === 'underline')) options.underline = { type: UnderlineType.DOUBLE };
          if (node.marks?.some(m => m.type === 'strike')) { options.strike = false; options.doubleStrike = true; }
        }
        if (mark.attrs?.fontFamily)
          options.font = String(mark.attrs.fontFamily)
            .split(",")[0]
            .replace(/["']/g, "");
        const size = cssPoints(mark.attrs?.fontSize);
        if (size) options.size = Math.round(size * 2);
        const color = rgb(mark.attrs?.color);
        if (color) options.color = color;
        const background = rgb(mark.attrs?.backgroundColor);
        if (background) options.shading = { fill: background };
        const tracking = Number.parseFloat(mark.attrs?.letterSpacing);
        if (Number.isFinite(tracking)) options.characterSpacing = Math.round(tracking * (String(mark.attrs?.letterSpacing).endsWith('pt') ? 20 : 15));
        const scale = parseCharacterScale(mark.attrs?.characterScale);
        if (scale && scale !== 100) options.scale = scale;
        break;
      }
    }
  }
  if (node.marks?.some(mark => mark.type === 'textStyle' && mark.attrs?.decorationStyle === 'double')) {
    if (node.marks.some(mark => mark.type === 'underline')) options.underline = { type: UnderlineType.DOUBLE };
    if (node.marks.some(mark => mark.type === 'strike')) { options.strike = false; options.doubleStrike = true; }
  }
  return options;
}
function alignment(value: unknown) {
  return (
    (
      {
        left: AlignmentType.LEFT,
        center: AlignmentType.CENTER,
        right: AlignmentType.RIGHT,
        justify: AlignmentType.JUSTIFIED,
      } as const
    )[String(value) as "left"] ?? AlignmentType.LEFT
  );
}
function paragraphOptions(node: JSONContent, defaultAfter = 150, measuredCellLine?: number, bodyLine?: HwpBodyLine | null): IParagraphOptions {
  const cellLine = activeHwpCellLine(node.attrs);
  // Word's AUTO percentage scales its font line metrics, whereas native HWP
  // scales the em size. Keep the established physical fallback for other
  // readers; our ignorable policy restores live, run-dependent leading here.
  const lineHeight = bodyLine && !bodyLine.precisePitch ? `${hwpBodyFallback(node, bodyLine)}px` : cellLine ? `${measuredCellLine ?? cellLine.base}px` :
    node.attrs?.lineSpacing ??
    node.attrs?.lineHeight ??
    node.content
      ?.flatMap((n) => n.marks ?? [])
      .find((m) => m.type === "textStyle")?.attrs?.lineHeight;
  const parsedLine = Number.parseFloat(lineHeight);
  // An explicit normal line is single spacing. Omitting it inherits the
  // export document's 1.15 default and stretches imported tables on reopen.
  const line = lineHeight === 'normal' ? 1 : typeof lineHeight === 'string' && lineHeight.endsWith('%') ? parsedLine / 100 : parsedLine;
  const indent = Number(node.attrs?.indent);
  const right = Number.parseFloat(node.attrs?.rightIndent), first = Number.parseFloat(node.attrs?.firstLineIndent);
  const before = Number.parseFloat(node.attrs?.spaceBefore), after = Number.parseFloat(node.attrs?.spaceAfter);
  const exact = typeof lineHeight === 'string' && /(?:pt|px)$/.test(lineHeight);
  // Fixed text leading must not crop taller inline equations/pictures. Keep
  // explicit OOXML exact rules, but let imported CSS image lines grow.
  const inlinePicture = !node.attrs?.wordLineRule && node.content?.some(n => n.type === 'inlineImage' && !n.attrs?.floating);
  // Plain authored paragraphs inherit 11pt / 1.65em in the editor. Falling
  // through to docx's 1.15-line / 6pt defaults shrinks long documents on reopen.
  // A minimum physical line height preserves that CSS baseline while allowing
  // large glyphs and inline pictures to grow; imported Word rules stay intact.
  const authoredDefault = node.type === 'paragraph' && !lineHeight;
  const headingDefaults = wordHeadingMarks(node.attrs?.wordHeadingMarks);
  const bottomBorder = parseWordParagraphBottomBorder(node.attrs?.wordParagraphBottomBorder);
  const borderBox = parseParagraphBorderBox(node.attrs?.paragraphBorderBox);
  const defaultLine = Math.round(Math.max(cssPoints(node.attrs?.paragraphSize) ?? 11,
    ...(node.content ?? []).flatMap(n => (n.marks ?? []).filter(m => m.type === 'textStyle')
      .map(m => cssPoints(m.attrs?.fontSize) ?? 0))) * 1.65 * 20);
  return {
    alignment: alignment(node.attrs?.textAlign),
    // docx 9.x's true option emits w:wordWrap w:val="0" (disable breaking
    // inside words); passing false omits the XML element entirely.
    ...(node.attrs?.wordBreak === 'keep-all' ? { wordWrap: true } : {}),
    tabStops: parseTabStops(node.attrs?.tabStops).map(stop => ({ position: Math.round(stop.position * 15), type: stop.type,
      ...(['dot', 'hyphen', 'underscore'].includes(stop.leader || '') ? { leader: stop.leader } : {}) })),
    indent: {
      ...(Number.isFinite(indent) ? { left: Math.round(indent * 15) } : {}),
      ...(Number.isFinite(right) ? { right: Math.round(right * 15) } : {}),
      ...(Number.isFinite(first) ? first < 0 ? { hanging: Math.round(-first * 15) } : { firstLine: Math.round(first * 15) } : {}),
    },
    run: { ...runOptions({ type: 'text', marks: [{ type: 'textStyle', attrs: {
      fontFamily: node.attrs?.paragraphFont, fontSize: node.attrs?.paragraphSize, color: node.attrs?.paragraphColor,
    } }] }), ...(headingDefaults ? { bold: headingDefaults.bold, italics: headingDefaults.italic, strike: headingDefaults.strike,
      underline: { type: headingDefaults.underline ? UnderlineType.SINGLE : UnderlineType.NONE } } : {}), text: undefined },
    spacing: {
      after: Number.isFinite(after) ? Math.round(after * 15) : node.type === 'paragraph' ? defaultAfter : 120,
      ...(Number.isFinite(before) ? { before: Math.round(before * 15) } : {}),
      ...(authoredDefault ? { line: defaultLine, lineRule: LineRuleType.AT_LEAST } : {}),
      ...(exact && Number.isFinite(line) && line > 0 ? { line: Math.round(line * (lineHeight.endsWith('pt') ? 20 : 15)), lineRule: node.attrs?.wordLineRule === 'atLeast' || inlinePicture ? LineRuleType.AT_LEAST : LineRuleType.EXACT } : {}),
      ...(!exact && Number.isFinite(line) && line > 0 && line < 10
        ? { line: Math.round(line * 240), lineRule: LineRuleType.AUTO }
        : {}),
    },
    ...(typeof node.attrs?.wordContextualSpacing === 'boolean' ? { contextualSpacing: node.attrs.wordContextualSpacing } : {}),
    ...(borderBox ? { border: paragraphBorderBoxOOXML(borderBox) } : bottomBorder ? { border: { bottom: bottomBorder } } : {}),
    // Flatten the visible policy. Native/HWP headings have no Word keep
    // attributes; do not acquire docx's built-in heading keeps on reopen.
    keepNext: node.attrs?.wordKeepNext === true,
    keepLines: node.attrs?.wordKeepLines === true,
    // Word imports resolve their enabled default. Unmarked Studio paragraphs
    // must keep the splits they had before the first DOCX save.
    widowControl: node.attrs?.wordWidowControl === true,
  };
}

async function imageRun(
  node: JSONContent,
  availablePixels: number,
  geometries: Map<string, ImageGeometryRecord>,
): Promise<ImageRun> {
  const src = node.attrs?.src;
  if (!validImage(src))
    fail("그림을 DOCX로 저장할 수 없습니다. 로컬 그림을 다시 삽입해 주세요.");
  const decoded = await new Promise<HTMLImageElement>((resolve, reject) => {
    const img = new Image();
    const timeout = window.setTimeout(
      () => reject(new Error("그림을 읽는 시간이 초과되었습니다.")),
      10000,
    );
    img.onload = () => {
      window.clearTimeout(timeout);
      resolve(img);
    };
    img.onerror = () => {
      window.clearTimeout(timeout);
      reject(new Error("손상된 그림이 있어 DOCX를 저장하지 못했습니다."));
    };
    img.src = src;
  });
  let data = src;
  let type = src.match(RASTER_DATA)![1].toLowerCase();
  if (type === "jpeg") type = "jpg";
  if (type === "webp") {
    const canvas = document.createElement("canvas");
    canvas.width = decoded.naturalWidth;
    canvas.height = decoded.naturalHeight;
    const context = canvas.getContext("2d");
    if (!context) fail("그림 변환을 시작하지 못했습니다.");
    context.drawImage(decoded, 0, 0);
    data = canvas.toDataURL("image/png");
    type = "png";
  }
  let width = Number(node.attrs?.width) || decoded.naturalWidth;
  let height =
    Number(node.attrs?.height) ||
    (width * decoded.naturalHeight) / decoded.naturalWidth;
  const floating = imageFloating(node.attrs?.floating);
  // Floating figures have explicit anchor geometry and may span a whole
  // paper or cross cell edges. Only inline pictures fit the text container.
  const scale = floating ? 1 : Math.min(1, availablePixels / width, 950 / height);
  width = Math.max(1, floating ? width : Math.round(width * scale));
  height = Math.max(1, floating ? height : Math.round(height * scale));
  const name = String(node.attrs?.title || '그림');
  const geometryId = floating?.polygon || floating?.drawingGroup || floating?.hwpBodyFlow || floating?.hwpPaperDrawing ||
    floating?.hwpCellRasterFill || floating?.hwpRasterPictureBrightness ? `wsFloat_${crypto.randomUUID()}` : undefined;
  if (geometryId) geometries.set(geometryId, { name, geometry: floating! });
  return new ImageRun({
    type: type as "png" | "jpg" | "gif" | "bmp",
    data,
    transformation: { width, height },
    ...(floating ? { floating } : {}),
    altText: {
      name: geometryId || name,
      title: String(node.attrs?.title || ""),
      description: String(node.attrs?.alt || ""),
    },
  });
}

export async function exportDocx(document: StudioDocument): Promise<Blob> {
  validateDocument(document);
  const { heights: cellLineHeights, positionedPairs, rasterCells } = await measureHwpExportLayout(document);
  const page = document.page;
  const geometries = new Map<string, ImageGeometryRecord>();
  const [portraitWidth, portraitHeight] = paperSize(page);
  let currentLeftMargin = pageMargins(page).left * 96 / 25.4;
  let currentPaperBottom = (page.orientation === 'landscape' ? portraitWidth : portraitHeight) * 96 / 25.4;
  let currentBodyBottom = currentPaperBottom - pageMargins(page).bottom * 96 / 25.4;
  let contentWidth = mmToTwip(
    (page.orientation === "landscape" ? portraitHeight : portraitWidth) -
      pageMargins(page).left - pageMargins(page).right,
  );
  const numbering: NumberingDefinition[] = [];
  const configuredLists = new Set<string>();
  const listLevels = new Map<string, Map<number, Record<string, unknown>>>();
  const notes = new Map<string, JSONContent>();
  // Materialize formatting as before, but retain the equivalence classes needed
  // by standard w:contextualSpacing. Never collide with built-in/spacer styles.
  const paragraphStyleNames = new Map<string, string>();
  const scannedFurniture = new Set<string>();
  const scanFurniture = (page: PageSettings) => {
    for (const key of ['headerHtml','footerHtml','firstHeaderHtml','firstFooterHtml','evenHeaderHtml','evenFooterHtml'] as const) {
      const html = page[key];
      if (html && !scannedFurniture.has(html)) { scannedFurniture.add(html); scan(parseDocumentHtml(sanitizeImportedHtml(html).html)); }
    }
  };
  const listStarts = new Map<string, Map<number, number>>();
  const scan = (node: JSONContent) => {
    if (node.type === 'orderedList') {
      const item = node.content?.[0]?.attrs;
      if (item?.wordListId) {
        const id = String(item.wordListId), level = Number(item.wordListLevel) || 0;
        const starts = listStarts.get(id) || new Map<number, number>();
        // numId is shared by nested levels and separate continuation shells.
        // Each level keeps its own first start, not the outer list's start.
        if (!starts.has(level)) starts.set(level, Math.max(1, Number(node.attrs?.start) || 1));
        listStarts.set(id, starts);
      }
    }
    const style = node.attrs?.wordParagraphStyle;
    if (typeof style === 'string' && ['paragraph', 'heading'].includes(node.type || '')) {
      if (!paragraphStyleNames.has(style)) paragraphStyleNames.set(style, node.attrs?.wordParagraphStyleName || style.replace(/^Heading([1-6])$/, 'Heading $1'));
    }
    if ((node.type === 'sectionBreak' || node.type === 'sectionContent') && node.attrs?.page) scanFurniture(JSON.parse(node.attrs.page));
    if (node.type === 'documentNote') notes.set(String(node.attrs?.id), node);
    if (node.type === 'listItem' && node.attrs?.wordListId) {
      const id = String(node.attrs.wordListId), levels = listLevels.get(id) || new Map();
      levels.set(Number(node.attrs.wordListLevel) || 0, node.attrs); listLevels.set(id, levels);
    }
    node.content?.forEach(scan);
  }; scan(document.content); scanFurniture(page);
  // Only reuse definitions actually emitted by docx. Subtitle is not one.
  const builtInStyles = new Set(['Normal', 'Title', 'ListParagraph', ...Array.from({ length: 6 }, (_, i) => `Heading${i + 1}`)]);
  // Unflagged classes must remain distinct too: falling back to Normal next
  // to a contextual Normal paragraph would suppress a different style's gap.
  const paragraphStyles = new Map([...paragraphStyleNames].map(([source, name], index) => {
    const builtIn = builtInStyles.has(source) && name.replace(/\s/g, '').toLowerCase() === source.toLowerCase();
    // Reuse matching standard IDs rather than introducing duplicate built-in
    // names. Custom style equivalence classes still receive collision-free IDs.
    return [source, { id: builtIn ? source : `StudioContext${index + 1}`, name, basedOn: 'Normal', builtIn }];
  }));
  const noteIds = new Map([...notes.keys()].map((id, index) => [id, index + 1]));
  let listId = 0;
  const commentIds = new Map(document.comments.map((comment, i) => [comment.id, i]));
  const remaining = new Map<string, number>();
  const bookmarkNames = new Set<string>(), bookmarkAliases = new Map<string, string>();
  const anchored = (node: JSONContent): string[] => (node.marks?.find(m => m.type === 'commentAnchor')?.attrs?.ids || '').split(' ').filter((id: string) => commentIds.has(id));
  const count = (node: JSONContent) => {
    if (node.type === 'documentBookmark') bookmarkNames.add(String(node.attrs?.name || ''));
    if (node.type === 'text' || node.type === 'documentTab') for (const id of anchored(node)) remaining.set(id, (remaining.get(id) ?? 0) + 1);
    node.content?.forEach(count);
  };
  count(document.content);
  const bookmarkName = (name: string) => {
    bookmarkNames.add(name);
    // Word rejects names containing spaces or punctuation, even when shorter than 40 characters.
    // Keep original Hancom names in our bounded custom-property alias map.
    if (name.length <= 40 && /^[\p{L}_][\p{L}\p{N}_]*$/u.test(name)) return name;
    if (bookmarkAliases.has(name)) return bookmarkAliases.get(name)!;
    let index = bookmarkAliases.size, alias = `ws_bookmark_${index}`;
    while (bookmarkNames.has(alias)) alias = `ws_bookmark_${++index}`;
    bookmarkNames.add(alias); bookmarkAliases.set(name, alias); return alias;
  };
  const opened = new Set<string>();
  let wroteUnanchored = false;
  const inline = async (nodes: JSONContent[], width: number, defaults: IRunOptions = {}): Promise<ParagraphChild[]> => {
    const runs: ParagraphChild[] = [];
    if (!wroteUnanchored) {
      wroteUnanchored = true;
      for (const [id, numeric] of commentIds) if (!remaining.has(id))
        runs.push(new CommentRangeStart(numeric), new CommentRangeEnd(numeric), new TextRun({ children: [new CommentReference(numeric)] }));
    }
    for (const node of nodes) {
      if (node.type === "text" || node.type === "documentTab") {
        const ids = anchored(node);
        for (const id of ids) if (!opened.has(id)) { runs.push(new CommentRangeStart(commentIds.get(id)!)); opened.add(id); }
        // Paragraph rPr formats the paragraph mark in OOXML, not every text
        // run. Materialize inherited CSS font/size/color on otherwise plain
        // text, while retaining each run's explicit overrides.
        const reflection = node.type === 'text' ? node.marks?.find(mark => mark.type === 'wordReflection')?.attrs?.effect : undefined;
        const textStyle = node.marks?.find(mark => mark.type === 'textStyle')?.attrs;
        const latinFont = activeHwpLatinFont(textStyle?.hwpLatinFont, textStyle?.fontFamily) ? textStyle?.hwpLatinFont : undefined;
        const run = documentTextRun({ ...defaults, ...runOptions(node), ...(node.type === 'documentTab' ? { text: '\t' } : {}) }, reflection, latinFont);
        const link = node.marks?.find((m) => m.type === "link");
        if (
          link &&
          safeLink(link.attrs?.href)
        )
          runs.push(
            link.attrs.href.startsWith('#')
              ? new InternalHyperlink({ children: [run], anchor: bookmarkName(link.attrs.href.slice(1)) })
              : new ExternalHyperlink({ children: [run], link: link.attrs.href }),
          );
        else runs.push(run);
        for (const id of ids) {
          remaining.set(id, remaining.get(id)! - 1);
          if (!remaining.get(id)) runs.push(new CommentRangeEnd(commentIds.get(id)!), new TextRun({ children: [new CommentReference(commentIds.get(id)!)] }));
        }
      } else if (node.type === "hardBreak")
        runs.push(new TextRun({ ...defaults, break: 1 }));
      // The actual picture run below becomes the one standard inline group.
      else if(node.type==='documentDrawingAnchor')continue;
      else if (node.type === 'documentBookmark') runs.push(new Bookmark({ id: bookmarkName(String(node.attrs?.name || '')), children: [] }));
      else if (node.type === 'documentField') runs.push(new TextRun({ ...defaults, children: [node.attrs?.kind === 'NUMPAGES' ? PageNumber.TOTAL_PAGES : PageNumber.CURRENT] }));
      else if (node.type === 'documentSymbol') runs.push(new SymbolRun({ ...defaults, ...runOptions(node), symbolfont: String(node.attrs?.font || 'Symbol'), char: Number(node.attrs?.code).toString(16).padStart(4, '0') }));
      else if (node.type === 'noteReference') {
        const id = noteIds.get(String(node.attrs?.id));
        if (id) runs.push(node.attrs?.kind === 'endnote' ? new EndnoteReferenceRun(id) : new FootnoteReferenceRun(id));
      }
      else if (node.type === 'hwpBoxedDigit') {
        const box = parseHwpBoxedDigit(node.attrs?.box);
        if (box) runs.push(new HwpBoxedDigitRun(box, { ...defaults, ...runOptions(node) }));
      }
      else if (node.type === 'wordOverlap') {
        const parts = overlapParts(node.attrs?.parts);
        const field = new SimpleField(`EQ \\o\\ac(${parts.join(',')})`);
        field.addChildElement(new TextRun({ ...defaults, ...runOptions(node), text: parts.join('') }));
        runs.push(field);
      }
      else if (node.type === "image" || node.type === "inlineImage")
        runs.push(await imageRun(node, width / 15, geometries));
    }
    return runs;
  };
  type Context = {
    depth?: number;
    indent?: number;
    cellRun?: IRunOptions;
    cellRaster?: { anchor: JSONContent; run: ImageRun };
    tableGap?: number;
    bodyTableOffsets?: boolean;
    width?: number;
    list?: { reference: string; level: number };
    importedList?: boolean;
  };
  // Consumers coalesce adjacent tables when no paragraph separates them. The
  // later narrow table can then resize every preceding row of the long table.
  const tableBoundary = (gap = 0, pageSpacing = false) => {
    const twips = Math.max(1, Math.round(gap * 15));
    // Between body tables this is a margin, not a real blank text line.
    // An exact-height spacer travels to the next page and pushes its title
    // down; ordinary before-spacing can be discarded at that page boundary.
    // Keep the full same-page gap and the existing cell-end line reserve.
    return new Paragraph({ style: TABLE_GAP_STYLE, spacing: {
      before: pageSpacing ? twips - 1 : 0, after: 0,
      line: pageSpacing ? 1 : twips, lineRule: LineRuleType.EXACT,
    }, run: { size: 2 } });
  };
  const textBoxes: TextBoxExport[] = [];
  const tableSpacing = new WeakMap<Table, number>();
  const blocks = async (
    nodes: JSONContent[],
    context: Context = {},
  ): Promise<(Paragraph | Table)[]> => {
    const width = context.width ?? contentWidth;
    const output: (Paragraph | Table)[] = [];
    for (const [nodeIndex, node] of nodes.entries()) {
      if (node.type === 'documentTextBox') {
        const frame = textBoxFrame(node.attrs?.frame);
        if (!frame) throw Error('텍스트상자 설정이 올바르지 않습니다.');
        const id = 'ws_tbox_' + crypto.randomUUID().replaceAll('-','').slice(0,16);
        textBoxes.push({ id, frame });
        output.push(new Paragraph({ children: [new Bookmark({ id:id+'_start', children:[] })] }),
          ...await blocks(node.content || [], { ...context, width:Math.max(1,(frame.width-frame.padding[1]-frame.padding[3])*15) }),
          new Paragraph({ children: [new Bookmark({ id:id+'_end', children:[] })] }));
      } else if (
        node.type === "paragraph" ||
        node.type === "heading" ||
        node.type === "codeBlock"
      ) {
        const headings = [
          HeadingLevel.HEADING_1,
          HeadingLevel.HEADING_2,
          HeadingLevel.HEADING_3,
          HeadingLevel.HEADING_4,
          HeadingLevel.HEADING_5,
          HeadingLevel.HEADING_6,
        ];
        const sourceBodyLine = hwpBodyTextLine(node) ?? (context.list ? null : hwpBodyEmptyLine(node));
        const bodyLine = sourceBodyLine && Boolean(sourceBodyLine.cell) === Boolean(context.cellRun) ? sourceBodyLine : null;
        let paragraph = paragraphOptions(node, context.list ? 75 : context.cellRun ? 0 : 150, cellLineHeights.get(node), bodyLine);
        if (context.importedList) {
          const textStops = paragraph.tabStops ?? [];
          const textPositions = new Set(textStops.map(stop => stop.position));
          // Numbering levels are regenerated on export. Materialize effective
          // list stops per paragraph so a direct override stays local.
          const numberingStops = parseWordNumberingTabs(node.attrs?.wordNumberingTabs)
            .filter(position => !textPositions.has(position))
            .map(position => ({ position, type: TabStopType.NUM }));
          if (numberingStops.length) paragraph = { ...paragraph, tabStops: [...textStops, ...numberingStops].sort((a, b) => a.position - b.position) };
        }
        const sourceExactLine = bodyLine && (bodyLine.align === 'column-end' || bodyLine.precisePitch);
        const retainsBodyLine = !sourceExactLine || paragraph.spacing?.line === Math.round(bodyLine.base * 15);
        const inlineTableSpaces = node.type === 'paragraph' && node.attrs?.hwpInlineTableSpaces === true &&
          paragraph.spacing?.lineRule === LineRuleType.EXACT && paragraph.spacing.line === 2 &&
          !!node.content?.length && node.content.every(part => part.type === 'text' && /^ +$/.test(part.text ?? ''));
        const sourceStyle = paragraphStyles.get(node.attrs?.wordParagraphStyle);
        if (sourceStyle) paragraph = { ...paragraph, style: sourceStyle.id };
        const cellLine = activeHwpCellLine(node.attrs);
        const previous = output.at(-1);
        const gap = previous instanceof Table ? tableSpacing.get(previous) || 0 : 0;
        // Inline OOXML tables have no bottom-margin property. Carry their CSS
        // gap as the next paragraph's spacing, which Word and other readers
        // render too. Do not silently drop it at a page boundary.
        if (gap) paragraph = { ...paragraph, spacing: { ...paragraph.spacing, before: Math.max(Number(paragraph.spacing?.before) || 0, Math.round(gap * 15)) } };
        const runs = await inline(node.content ?? [], width, { ...context.cellRun, ...paragraph.run,
          // An absent emphasis mark in the model means normal text.
          ...(wordHeadingMarks(node.attrs?.wordHeadingMarks) ? { bold: false, italics: false, strike: false,
            underline: { type: UnderlineType.NONE } } : {}),
        });
        if (context.cellRaster?.anchor === node) runs.unshift(context.cellRaster.run);
        output.push(
          new HwpCellParagraph({
            ...paragraph,
            children: runs,
            ...(context.cellRun ? { run: { ...context.cellRun, ...paragraph.run } } : {}),
            ...(node.type === "heading"
              ? sourceStyle ? { outlineLevel: (node.attrs?.level ?? 1) - 1 } : { heading: headings[(node.attrs?.level ?? 1) - 1] }
              : {}),
            // Explicit0–8 and body-text9 override a heading style's default
            // outline, while leaving the source keep and visual kind intact.
            ...(node.attrs?.wordOutlineLevel != null ? { outlineLevel: node.attrs.wordOutlineLevel }
              : node.type === 'heading' ? { outlineLevel: (node.attrs?.level ?? 1) - 1 } : {}),
            ...(node.type === "codeBlock"
              ? {
                  run: { font: "Consolas", size: 20 },
                  shading: { fill: "F3F4F6" },
                }
              : {}),
            ...(context.list ? { numbering: context.list } : {}),
            ...(context.indent ? { indent: { left: context.indent } } : {}),
          }, cellLine ? JSON.stringify({
            policy: encodeHwpCellLine(cellLine), line: paragraph.spacing?.line,
            // Run sizes are serialized as integral half-points too. Compare
            // against that representation, including source sizes like 9.6pt.
            font: Math.round(hwpCellFont(node, cellLine) * 1.5) / 1.5,
            ...(cssPixels(node.attrs?.paragraphSize) !== null ? {
              strutFont: cssPixels(node.attrs?.paragraphSize), defaultFont: Number(paragraph.run?.size),
            } : {}),
          }) : undefined, bodyLine && retainsBodyLine ? JSON.stringify({ policy: encodeHwpBodyLine(bodyLine),
            line: paragraph.spacing?.line, rule: paragraph.spacing?.lineRule }) : undefined,
            { de: node.attrs?.wordAutoSpaceDE, dn: node.attrs?.wordAutoSpaceDN }, node.attrs?.wordTextAlignment, node.attrs?.wordParagraphReflection,
            node.attrs?.hwpSourceSpaces === true, node.attrs?.hwpSourceWrap, inlineTableSpaces),
        );
      } else if (node.type === "image")
        output.push(
          new Paragraph({
            children: [await imageRun(node, width / 15, geometries)],
            alignment: alignment(node.attrs?.textAlign),
          }),
        );
      else if (node.type === "pageBreak")
        output.push(new Paragraph({ children: [new PageBreak()], spacing: { before: 0, after: 0, line: 1, lineRule: LineRuleType.EXACT } }));
      else if (node.type === 'columnBreak')
        output.push(new Paragraph({ children: [new ColumnBreak()], spacing: { before: 0, after: 0, line: 1, lineRule: LineRuleType.EXACT } }));
      else if (node.type === "horizontalRule")
        output.push(
          new Paragraph({
            border: {
              bottom: { color: "B8C2CC", style: BorderStyle.SINGLE, size: 6 },
            },
          }),
        );
      else if (node.type === "blockquote")
        output.push(
          ...(await blocks(node.content ?? [], {
            ...context,
            indent: (context.indent ?? 0) + 360,
          })),
        );
      else if (node.type === 'columnGroup' || node.type === 'sectionContent') output.push(...await blocks(node.content || [], context));
      else if (node.type === "bulletList" || node.type === "orderedList") {
        // An HTML shell may contain several Word numIds. Configure and use
        // each item's own definition, keeping one shared id for authored items.
        let authoredReference: string | undefined;
        for (const item of node.content ?? []) {
          const imported = item.attrs;
          const reference = imported?.wordListId ? `studio-imported-${imported.wordListId}` : (authoredReference ??= `studio-list-${++listId}`);
          const level = Math.min(Number(imported?.wordListLevel ?? context.depth ?? 0), 8);
          const definitions = listLevels.get(String(imported?.wordListId));
          if (!configuredLists.has(reference)) numbering.push({
            reference,
            levels: Array.from({ length: 9 }, (_, i) => {
              const source = definitions?.get(i), font = wordListFontName(source?.wordListFont);
              return {
                level: i,
                format: (source?.wordListFormat as typeof LevelFormat.DECIMAL) || (node.type === "orderedList" ? LevelFormat.DECIMAL : LevelFormat.BULLET),
                text: String(source?.wordListPattern || (node.type === "orderedList" ? `%${i + 1}.` : ["•", "◦", "▪"][i % 3])),
                start: imported?.wordListId ? listStarts.get(String(imported.wordListId))?.get(i) ?? Math.max(1, Number(source?.wordListStart) || 1)
                  : node.type === "orderedList" ? Math.max(1, Number(node.attrs?.start) || 1) : 1,
                alignment: AlignmentType.LEFT,
                suffix: source?.wordListSuffix === 'space' ? LevelSuffix.SPACE : source?.wordListSuffix === 'nothing' ? LevelSuffix.NOTHING : LevelSuffix.TAB,
                style: { ...(font ? { run: { font } } : {}), paragraph: { indent: { left: 720 * (i + 1), hanging: 360 } } },
              };
            }),
          });
          configuredLists.add(reference);
          let first = true;
          for (const child of item.content ?? []) {
            const isParagraph = ["paragraph", "heading"].includes(child.type ?? "");
            output.push(...await blocks([child], {
              cellRun: context.cellRun, width: context.width, tableGap: context.tableGap,
              depth: level + 1,
              ...(first && isParagraph ? { list: { reference, level }, importedList: !!imported?.wordListId } : {}),
              ...(!first && isParagraph ? { indent: 720 * (level + 1) } : {}),
            }));
            if (isParagraph) first = false;
          }
        }
      } else if (node.type === "table") {
        const rows: TableRow[] = [];
        const columnWidths: number[] = [];
        const firstRow = node.content?.[0]?.content ?? [];
        for (const cell of firstRow) {
          const span = Math.max(1, Number(cell.attrs?.colspan) || 1);
          for (let i = 0; i < span; i++)
            columnWidths.push(
              Number(cell.attrs?.colwidth?.[i])
                ? Math.round(cell.attrs!.colwidth[i] * 15)
                : 0,
            );
        }
        const fallback = Math.floor(
          width / Math.max(1, columnWidths.length),
        );
        for (let i = 0; i < columnWidths.length; i++)
          if (!columnWidths[i]) columnWidths[i] = fallback;
        const total = columnWidths.reduce((a, b) => a + b, 0);
        if (total > width)
          for (let i = 0; i < columnWidths.length; i++)
            columnWidths[i] = Math.max(
              1,
              Math.floor((columnWidths[i] * width) / total),
            );
        const continuingMergeRows = new Set<number>();
        (node.content ?? []).forEach((row, index) => {
          for (const cell of row.content ?? []) if (cell.attrs?.cellSplit === 'allow')
            for (let i = index; i < Math.min(node.content!.length, index + Number(cell.attrs?.rowspan || 1)); i++) continuingMergeRows.add(i);
        });
        let occupied: number[] = [];
        let continuations: { remaining: number; attrs: Record<string, unknown> }[] = [];
        const borderPixels = (attrs: Record<string, unknown>, side: string) => {
          const css = globalThis.document.createElement('span').style;
          css.border = String(attrs[`border${side}`] ?? '1px solid #c6ced8');
          return ['none', 'hidden'].includes(css.borderStyle) ? 0 : (cssPoints(css.borderWidth) ?? .75) * 4 / 3;
        };
        for (const [rowIndex, row] of (node.content ?? []).entries()) {
          const cells: TableCell[] = [];
          // Subtract the same border reserve Word adds to minimum heights,
          // including continuation cells materialized by the DOCX writer.
          const rowBorder = Math.max(0,
            ...continuations.flatMap(c => ['Top', 'Bottom'].map(side => borderPixels(c.attrs, side))),
            ...(row.content ?? []).filter(c => Number(c.attrs?.rowspan || 1) === 1).map(c => borderPixels(c.attrs ?? {}, 'Bottom')));
          const insideMerge = continuations.length > 0 && (row.content ?? []).every(c => Number(c.attrs?.rowspan) > 1);
          // A padded merge start spans several rows. Its padding can exceed
          // the native minimum of the first row next to an empty grid cell.
          const besideMergeStart = !continuations.length && (row.content ?? []).some(c => Number(c.attrs?.rowspan) > 1) &&
            (row.content ?? []).some(c => Number(c.attrs?.rowspan || 1) === 1) &&
            (row.content ?? []).every(c => Number(c.attrs?.rowspan || 1) > 1 || (c.content?.length === 1 &&
              c.content[0].type === 'paragraph' && !c.content[0].content?.length));
          continuations = continuations.map(c => ({ ...c, remaining: c.remaining - 1 })).filter(c => c.remaining > 0);
          for (const cell of row.content ?? []) if (Number(cell.attrs?.rowspan) > 1)
            continuations.push({ remaining: Number(cell.attrs!.rowspan) - 1, attrs: cell.attrs ?? {} });
          const rowPadding = Math.max(0, ...(row.content ?? []).map(cell => ['Top', 'Bottom'].reduce((sum, side) => {
            const padding = parseFloat(cell.attrs?.[`padding${side}`]);
            return sum + (Number.isFinite(padding) ? Math.max(0, padding) : 8);
          }, 0)));
          let column = 0;
          const nextOccupied = occupied.map(remaining => Math.max(0, remaining - 1));
          for (const cell of row.content ?? []) {
            while ((occupied[column] || 0) > 0) column++;
            const span = Math.max(1, Number(cell.attrs?.colspan) || 1);
            const cellWidth = columnWidths.slice(column, column + span).reduce((sum, n) => sum + n, 0);
            const inset = ['Left', 'Right'].reduce((sum, side) => {
              const padding = parseFloat(cell.attrs?.[`padding${side}`]);
              return sum + (Number.isFinite(padding) ? Math.max(0, padding) * 15 : 150);
            }, 0);
            if (Number(cell.attrs?.rowspan) > 1) for (let i = 0; i < span; i++) nextOccupied[column + i] = Number(cell.attrs!.rowspan) - 1;
            column += span;
            const content = inheritCellText(cell.content ?? [], cell.attrs);
            const fill = parseHwpCellRasterFill(cell.attrs?.hwpCellRasterFill);
            const placement = rasterCells.get(cell);
            const left = parseFloat(cell.attrs?.paddingLeft), top = parseFloat(cell.attrs?.paddingTop);
            const fillRun = fill && await imageRun({ type: 'inlineImage', attrs: {
              src: fill.source, width: placement?.width ?? fill.width, height: placement?.height ?? fill.height,
              title: '그림으로 채운 셀',
              floating: JSON.stringify({
                horizontalPosition: { relative: 'column', offset: -Math.round((Number.isFinite(left) ? left : 10) * 9525) },
                verticalPosition: { relative: 'paragraph', offset: -Math.round((placement?.paragraphOffset ?? (Number.isFinite(top) ? top : 8)) * 9525) },
                wrap: { type: 0 }, behindDocument: true, layoutInCell: true, hwpCellRasterFill: true,
                hwpCellRasterBrightness: fill.brightness,
              }),
            } }, Math.max(1, (cellWidth - inset) / 15), geometries);
            const fillAnchor = fillRun && content.find(node => node.type === 'paragraph' || node.type === 'heading');
            const children = await blocks(content, {
              width: Math.max(1, cellWidth - inset),
              tableGap: context.tableGap,
              ...(fillRun && fillAnchor && { cellRaster: { anchor: fillAnchor, run: fillRun } }),
              cellRun: runOptions({ type: 'text', marks: [{ type: 'textStyle', attrs: {
                fontFamily: cell.attrs?.fontFamily, fontSize: cell.attrs?.fontSize, color: cell.attrs?.textColor,
              } }] }),
            });
            if (fillRun && !fillAnchor)
              children.push(new Paragraph({ children: [fillRun], spacing: { before: 0, after: 0, line: 1, lineRule: LineRuleType.EXACT } }));
            if (
              !children.length ||
              children[children.length - 1] instanceof Table
            )
              children.push(tableBoundary(tableSpacing.get(children.at(-1) as Table) || 0));
            const color =
              cell.attrs?.backgroundColor === 'transparent' ? undefined : rgb(cell.attrs?.backgroundColor) ??
              (cell.type === "tableHeader" ? "E8EEF7" : undefined);
            cells.push(
              new PreservedTableCell({
                children,
                ...(cell.attrs?.wordCellWidthAuto === true ? { width: { size: 0, type: WidthType.AUTO } } : {}),
                columnSpan: Number(cell.attrs?.colspan) || 1,
                rowSpan: Number(cell.attrs?.rowspan) || 1,
                ...(color || cell.attrs?.backgroundColor === 'transparent' ? { shading: { fill: color ?? 'auto' } } : {}),
                borders: cellBorders(cell.attrs),
                verticalAlign: cell.attrs?.verticalAlign === 'middle' ? VerticalAlignTable.CENTER : cell.attrs?.verticalAlign === 'bottom' ? VerticalAlignTable.BOTTOM : VerticalAlignTable.TOP,
                margins: Object.fromEntries(['top', 'right', 'bottom', 'left'].map(side => {
                  const padding = Number.parseFloat(cell.attrs?.[`padding${side[0].toUpperCase()}${side.slice(1)}`]);
                  return [side, Number.isFinite(padding) ? Math.round(Math.max(0, padding) * 15) : ['left', 'right'].includes(side) ? 150 : 120];
                })),
              }, cell.attrs?.whiteSpace === 'nowrap',
              cell.type === 'tableCell' && cell.attrs?.verticalAlign === 'middle' && span === 1 && (Number(cell.attrs?.rowspan) || 1) === 1 &&
                parseHwpCenteredCellLabels(cell.attrs?.hwpCenteredCellLabels) ? cell.attrs?.hwpCenteredCellLabels : undefined),
            );
          }
          occupied = nextOccupied;
          const cantSplit = exportedRowCantSplit(row.attrs, continuingMergeRows.has(rowIndex));
          const minimum = rowHeight(row.attrs?.rowHeight);
          const exact = row.attrs?.wordRowRule === 'exact';
          const inset = exact ? 0 : rowPadding + rowBorder;
          const gridHeight = row.attrs?.hwpGridRow === true && (insideMerge || besideMergeStart) && minimum !== null && minimum < inset ? minimum : undefined;
          const reserve = hwpRowReserve(row.attrs?.hwpRowReserve);
          const sourceContinuation = row.attrs?.hwpRowContinued === true && continuingMergeRows.has(rowIndex) && cantSplit === false;
          const preserveReserve = !exact && minimum !== null && reserve && Math.abs(reserve.authoredHeight - minimum) <= .1 &&
            (reserve.kind === 1 || reserve.kind === 4 ? sourceContinuation : cantSplit === (reserve.kind === 3));
          const intactTopRow = parseHwpIntactRow(row.attrs?.hwpIntactTopRow);
          rows.push(new PreservedTableRow({
            children: cells,
            ...(cantSplit !== undefined ? { cantSplit } : {}),
            ...(minimum === null ? {} : { height: {
              value: Math.max(1, Math.round((minimum - inset) * 15)),
              rule: exact ? HeightRule.EXACT : HeightRule.ATLEAST,
            } }),
            ...((row.content?.length ?? 0) > 0 && (row.content ?? []).every(
              cell => cell.type === "tableHeader",
            ) ? { tableHeader: true } : {}),
          }, gridHeight, sourceContinuation,
          preserveReserve ? row.attrs?.hwpRowReserve : undefined, !exact && cantSplit === false && row.attrs?.hwpRowIntact === true,
          cantSplit === false && !exact && minimum !== null && intactTopRow && Math.abs(intactTopRow.height - minimum) <= .16
            ? row.attrs?.hwpIntactTopRow : undefined));
        }
        const requestedAnchor = tableAnchor(node.attrs?.tableAnchor), overlayHeight = cellLineHeights.get(node);
        const unsupportedWrap = requestedAnchor?.wrap && (!['paragraph', 'heading'].includes(nodes[nodeIndex + 1]?.type ?? '') ||
          !tableWrapFits(requestedAnchor, columnWidths.reduce((a, b) => a + b, 0) / 15, width / 15));
        const anchor = requestedAnchor?.overlay && !(overlayHeight && overlayHeight > 0) || unsupportedWrap ? undefined : requestedAnchor;
        const paragraphAnchor = anchor?.reference === 'paragraph' && !anchor.wrap ? anchor : undefined;
        const nestedPosition = positionedPairs.has(node) && cellTablePosition(node.attrs?.cellTablePosition);
        const overlayPosition = anchor?.overlay ? { x: Math.round(anchor.x * 15) / 15,
          y: Math.round(((anchor.overlay === 'page-bottom' ? currentBodyBottom : currentPaperBottom) - anchor.y - overlayHeight!) * 15) / 15,
          width: columnWidths.reduce((a, b) => a + b, 0) / 15 } : undefined;
        const inlineIndent = paragraphAnchor ? paragraphAnchor.x - (paragraphAnchor.horizontal === 'margin' ? 0 : currentLeftMargin) : undefined;
        // Word only honors noWrap under autofit. Keep the explicit policy of
        // an imported Word table; otherwise preserve nowrap before grid locking.
        const noWrap = (node.content ?? []).some(row => (row.content ?? []).some(cell => cell.attrs?.whiteSpace === 'nowrap'));
        const autofit = node.attrs?.wordTableLayout === 'autofit' || node.attrs?.wordTableLayout !== 'fixed' && noWrap;
        const firstRowCells = node.content?.[0]?.content;
        const repeated = firstRowCells?.length && firstRowCells.every(cell => cell.type === 'tableHeader');
        const table = new PreservedTable({
            rows,
            layout: autofit ? TableLayoutType.AUTOFIT : TableLayoutType.FIXED,
            width: node.attrs?.wordTableWidthAuto === true ? { size: 0, type: WidthType.AUTO } : { size: columnWidths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
            ...(paragraphAnchor || anchor?.wrap ? { alignment: AlignmentType.LEFT } : ['center', 'right'].includes(node.attrs?.tableAlignment) ? { alignment: node.attrs?.tableAlignment === 'center' ? AlignmentType.CENTER : AlignmentType.RIGHT } : {}),
            ...(anchor?.wrap ? { indent: { size: 0, type: WidthType.DXA } } : inlineIndent !== undefined || node.attrs?.tableIndent ? { indent: { size: Math.round((inlineIndent ?? Number.parseFloat(node.attrs!.tableIndent)) * 15) || 0, type: WidthType.DXA } } : {}),
            columnWidths,
            ...(nestedPosition ? { float: {
              horizontalAnchor: TableAnchorType.MARGIN, verticalAnchor: TableAnchorType.MARGIN,
              absoluteHorizontalPosition: Math.round(nestedPosition.x * 15), absoluteVerticalPosition: Math.round(nestedPosition.y * 15),
              leftFromText: 0, rightFromText: 0, topFromText: 0, bottomFromText: 0,
            } } : anchor && !paragraphAnchor ? { float: { horizontalAnchor: anchor.horizontal === 'margin' ? TableAnchorType.MARGIN : TableAnchorType.PAGE,
              verticalAnchor: anchor.reference === 'paragraph' ? TableAnchorType.TEXT : TableAnchorType.PAGE,
              ...(anchor.horizontalAlign === 'center' ? { relativeHorizontalPosition: RelativeHorizontalPosition.CENTER }
                : { absoluteHorizontalPosition: Math.round(anchor.x * 15) }),
              absoluteVerticalPosition: Math.round((overlayPosition?.y ?? anchor.y) * 15), bottomFromText: Math.round(anchor.bottom * 15),
              // Page floats wrap above/below. A left-side paragraph label
              // permits text only on its right and below the actual table.
              leftFromText: overlayPosition ? 0 : contentWidth, rightFromText: anchor.wrap ? Math.round(anchor.right! * 15) : overlayPosition ? 0 : contentWidth } } : {}),
          }, overlayPosition || anchor?.wrap ? anchor : paragraphAnchor, overlayPosition,
          repeated && !anchor ? hwpTableTopGap(node.attrs?.hwpTableTopGap) : undefined,
          !nestedPosition && (!anchor && node.attrs?.hwpCenteredBorders === 'inline' || repeated && (!anchor || !!paragraphAnchor))
            ? node.attrs?.hwpCenteredBorders === 'inline' || node.attrs?.hwpCenteredBorders === 'continuation'
              ? node.attrs.hwpCenteredBorders : node.attrs?.hwpCenteredBorders === true
            : false,
          !anchor && !nestedPosition ? hwpKeptTablePage(node.attrs?.hwpKeptTablePage) : undefined,
          repeated && !nestedPosition && (!anchor || !!paragraphAnchor) && typeof node.attrs?.hwpTableContinuation === 'string' &&
            parseHwpTableContinuation(node.attrs.hwpTableContinuation) ? node.attrs.hwpTableContinuation : undefined,
          !nestedPosition && parseHwpNestedTableLayout(node.attrs?.hwpNestedTableLayout)
            ? node.attrs?.hwpNestedTableLayout : undefined,
          !anchor && !nestedPosition && parseHwpRelatedTableLayout(node.attrs?.hwpRelatedTableLayout)
            ? node.attrs?.hwpRelatedTableLayout : undefined,
          !anchor && !nestedPosition ? hwpInlineTableTop(node.attrs?.hwpInlineTableTop) : undefined,
          !anchor && !nestedPosition && parseHwpNativeGrid(node.attrs?.hwpNativeGrid) ? node.attrs?.hwpNativeGrid : undefined,
          !anchor && !nestedPosition ? hwpInlineTableBottom(node.attrs?.hwpInlineTableBottom) : undefined);
        // Missing imported CSS uses the editor's 18px table gap. Furniture
        // and note tables have a zero default, passed through nested blocks.
        tableSpacing.set(table, anchor ? 0 : Math.max(0, cssPoints(node.attrs?.tableSpacing) ?? (context.tableGap ?? 18) * .75) * 4 / 3);
        const tableOffset = (pixels: number, before = 0) => new Paragraph({ style: TABLE_OFFSET_STYLE,
          spacing: { before, after: 0, line: Math.max(1, Math.round(pixels * 15)), lineRule: LineRuleType.EXACT }, run: { size: 2 } });
        if (paragraphAnchor?.y) {
          const previous = nodes[nodeIndex - 1], previousTable = output.at(-1);
          const gap = context.bodyTableOffsets && context.width === undefined && columnWidths.reduce((total, col) => total + col, 0) >= width * .8 &&
            previous?.type === 'table' && previousTable instanceof Table && cssPoints(previous.attrs?.tableSpacing) !== undefined ? tableSpacing.get(previousTable) || 0 : 0;
          // The inline table gap and the next anchor collapse to the larger
          // reserve in Studio. Keep the anchor's exact line and represent the
          // excess as discardable before-spacing on that same DOCX carrier.
          const offset = Math.max(1, Math.round(paragraphAnchor.y * 15));
          output.push(tableOffset(paragraphAnchor.y, Math.max(0, Math.round(gap * 15) - offset)));
        }
        output.push(table);
        if (paragraphAnchor?.bottom) output.push(tableOffset(paragraphAnchor.bottom));
      }
    }
    return output.flatMap((block, i) => i > 0 && block instanceof Table && output[i - 1] instanceof Table
      ? [tableBoundary(tableSpacing.get(output[i - 1] as Table), !context.cellRun), block] : [block]);
  };
  const groups: { page: PageSettings; nodes: JSONContent[] }[] = [{ page, nodes: [] }];
  let sectionPage = page;
  const columnSignature = (p: PageSettings) => JSON.stringify(p.columns ?? { count: 1, space: 0 });
  const appendGroup = (settings: PageSettings, nodes: JSONContent[]) => {
    const current = groups[groups.length - 1];
    if (!current.nodes.length) { current.page = settings; current.nodes.push(...nodes); }
    else if (columnSignature(current.page) === columnSignature(settings)) current.nodes.push(...nodes);
    else groups.push({ page: { ...settings, breakType: 'continuous' }, nodes: [...nodes] });
  };
  const sectionNodes = (document.content.content ?? []).flatMap(node => node.type === 'sectionContent' ? node.content || [] : [node]);
  for (const node of sectionNodes) {
    if (node.type === 'sectionBreak') {
      sectionPage = JSON.parse(node.attrs!.page); groups.push({ page: sectionPage, nodes: [] });
    } else if (node.type === 'columnGroup') {
      appendGroup({ ...sectionPage, columns: { count: Number(node.attrs?.count) || 1, space: Number(node.attrs?.space) || 0, balance: node.attrs?.balance !== false } }, node.content || []);
    } else appendGroup(sectionPage, [node]);
  }
  const sections = [];
  const differentOddEven = groups.some(group => group.page.differentOddEven);
  let hasTableEnd = false;
  const footnotes: Record<string, { children: Paragraph[] }> = {}, endnotes: Record<string, { children: Paragraph[] }> = {};
  const hwpFootnoteLayouts = new Map<number, string>();
  for (const [id, note] of notes) {
    // Both paragraph and table objects implement the block XML component
    // protocol, although docx's public note type lists Paragraph only.
    const children = await blocks(note.content || [], { tableGap: 0 });
    const number = noteIds.get(id)!;
    (note.attrs?.kind === 'endnote' ? endnotes : footnotes)[String(number)] = { children: children as Paragraph[] };
    const layout = note.attrs?.kind === 'footnote' && parseHwpFootnoteLayout(note.attrs?.hwpFootnoteLayout);
    if (layout && Number.parseInt(layout.mark) === number)
      hwpFootnoteLayouts.set(number, JSON.stringify(layout));
  }
  for (const group of groups) {
  const page = group.page;
  const [portraitWidth, portraitHeight] = paperSize(page);
  const margins = pageMargins(page);
  currentLeftMargin = margins.left * 96 / 25.4;
  currentPaperBottom = (page.orientation === 'landscape' ? portraitWidth : portraitHeight) * 96 / 25.4;
  currentBodyBottom = currentPaperBottom - margins.bottom * 96 / 25.4;
  const fullWidth = mmToTwip((page.orientation === 'landscape' ? portraitHeight : portraitWidth) - margins.left - margins.right);
  const count = page.columns?.count || 1;
  contentWidth = Math.max(1, Math.floor((fullWidth - Math.round((page.columns?.space || 0) * 15) * (count - 1)) / count));
  const children = await blocks(group.nodes, { bodyTableOffsets: count === 1 });
  // An omitted Word end-paragraph is created with Normal's 11pt leading and
  // can push a full-page form onto an otherwise blank second page.
  if (group === groups.at(-1) && children.at(-1) instanceof Table) {
    children.push(new Paragraph({ style: TABLE_END_STYLE, run: { size: 2 }, widowControl: false,
      spacing: { before: 0, after: 0, line: 20, lineRule: LineRuleType.EXACT } }));
    hasTableEnd = true;
  }

  const richBlocks = async (html: string) => blocks(parseDocumentHtml(sanitizeImportedHtml(html).html).content || [], { tableGap: 0 });
  const plainFurniture: IParagraphOptions = { run: { size: 18, color: '697386' }, spacing: { before: 0, after: 0, line: 336, lineRule: LineRuleType.AUTO } };
  // Keep explicit empty stories to prevent inheritance from earlier sections,
  // without reserving a normal text line at the header/footer distance.
  const emptyFurniture = () => new Paragraph({ run: { size: 2 }, spacing: { before: 0, after: 0, line: 1, lineRule: LineRuleType.EXACT } });
  const footer = async (text: string, numbers: boolean, html?: string, headerHtml?: string) => new Footer({ children: [
    ...(html ? await richBlocks(html) : text ? [new Paragraph({ ...plainFurniture, children: [documentTextRun({ text })], alignment: AlignmentType.CENTER })] : !numbers ? [emptyFurniture()] : []),
    ...(numbers && !/data-page-number/.test((html || '') + (headerHtml || '')) ? [new Paragraph({ ...plainFurniture, alignment: AlignmentType.CENTER,
      children: [new TextRun({ children: [PageNumber.CURRENT], size: 18, color: '697386' })] })] : []),
  ] });
  const header = async (text: string, html?: string) => new Header({ children: html ? await richBlocks(html) : [text ? new Paragraph({ ...plainFurniture, children: [documentTextRun({ text })] }) : emptyFurniture()] });
  const baseMargins = basePageMargins(page);
  sections.push(
      {
        children: children.length ? children : [new Paragraph("")],
        properties: {
          type: page.breakType || SectionType.NEXT_PAGE,
          ...(page.columns ? { column: { count: page.columns.count, space: Math.round(page.columns.space * 15), equalWidth: true } } : {}),
          titlePage: page.differentFirstPage,
          page: {
            size: {
              width: mmToTwip(portraitWidth),
              height: mmToTwip(portraitHeight),
              orientation:
                page.orientation === "landscape"
                  ? PageOrientation.LANDSCAPE
                  : PageOrientation.PORTRAIT,
            },
            margin: {
              top: mmToTwip(baseMargins.top + (page.gutterTop ? page.gutter ?? 0 : 0)),
              right: mmToTwip(baseMargins.right),
              bottom: mmToTwip(baseMargins.bottom),
              left: mmToTwip(baseMargins.left),
              gutter: mmToTwip(page.gutterTop ? 0 : page.gutter ?? 0),
              header: mmToTwip(page.headerDistance ?? Math.max(0, margins.top / 2 - 3)),
              footer: mmToTwip(page.footerDistance ?? Math.max(0, margins.bottom / 2 - 3)),
            },
          },
        },
        // Odd/even selection is global in DOCX. A mixed document needs a
        // default copy in sections that use one story for both page types.
        // With no global switch, keep explicit dormant stories and do not
        // invent even references that overwrite their values on a roundtrip.
        headers: { default: await header(page.header, page.headerHtml),
          first: await header(page.firstHeader ?? '', page.firstHeaderHtml),
          ...(differentOddEven || page.evenHeader !== undefined || page.evenHeaderHtml !== undefined
            ? { even: await header(differentOddEven && !page.differentOddEven ? page.header : page.evenHeader ?? '', differentOddEven && !page.differentOddEven ? page.headerHtml : page.evenHeaderHtml) } : {}) },
        footers: { default: await footer(page.footer, page.pageNumbers, page.footerHtml, page.headerHtml),
          first: await footer(page.firstFooter ?? '', page.firstPageNumbers ?? false, page.firstFooterHtml, page.firstHeaderHtml),
          ...(differentOddEven || page.evenFooter !== undefined || page.evenFooterHtml !== undefined || page.evenPageNumbers !== undefined
            ? { even: await footer(differentOddEven && !page.differentOddEven ? page.footer : page.evenFooter ?? '', differentOddEven && !page.differentOddEven ? page.pageNumbers : page.evenPageNumbers ?? false, differentOddEven && !page.differentOddEven ? page.footerHtml : page.evenFooterHtml, differentOddEven && !page.differentOddEven ? page.headerHtml : page.evenHeaderHtml) } : {}) },
      },
  );
  }
  const comments = await Promise.all(document.comments.map(async (comment, id) => ({ id, durableId: (id + 1).toString(16).padStart(8, '0').toUpperCase(), author: comment.author ?? 'Word Studio', initials: comment.initials, date: new Date(comment.createdAt),
    resolved: comment.resolved, children: comment.html ? await blocks(parseDocumentHtml(sanitizeImportedHtml(comment.html).html).content || []) : comment.text.split('\n').map(text => new Paragraph({ children: [documentTextRun({ text })] })) })));
  const docx = new DocxDocument({
    title: document.title,
    creator: wordStudioCreator,
    description: wordStudioDescription,
    customProperties: [
      ...(document.wordWrapTrailSpaces === undefined && document.wordDefaultHangingSpaces
        ? [{ name: wordDefaultHangingSpacesProperty, value: 'hang' }] : []),
      ...(groups.some(group => group.page.hideEmptyLines) ? [{ name: emptyLineDocxProperty, value: JSON.stringify(groups.flatMap((group, index) => group.page.hideEmptyLines ? [index] : [])) }] : []),
      ...(bookmarkAliases.size ? [{ name: bookmarkNamesProperty, value: JSON.stringify([...bookmarkAliases].map(([name, alias]) => [alias, name])) }] : []),
    ],
    evenAndOddHeaderAndFooters: differentOddEven,
    hyphenation: { autoHyphenation: document.wordAutoHyphenation },
    compatibility: { noColumnBalance: groups.some(group => group.page.columns?.balance === false), doNotBreakWrappedTables: false,
      suppressBottomSpacing: document.page.suppressBottomSpacing !== false, wrapTrailSpaces: document.wordWrapTrailSpaces },
    numbering: { config: numbering },
    footnotes, endnotes,
    comments: { children: comments },
    styles: { default: { document: { run: { font: '맑은 고딕', size: 22, language: document.wordDefaultLanguage }, paragraph: { spacing: { after: 120, line: 276 } } } },
      paragraphStyles: [
        ...[...paragraphStyles.values()].filter(style => !style.builtIn),
        ...[TABLE_GAP_STYLE, TABLE_ANCHOR_STYLE, TABLE_OFFSET_STYLE, SECTION_BOUNDARY_STYLE].map(id => ({ id, name: id === TABLE_GAP_STYLE ? 'Table spacing' : id === SECTION_BOUNDARY_STYLE ? 'Section boundary' : 'Table anchor', run: { size: 2 }, paragraph: { spacing: { before: 0, after: 0, line: 1, lineRule: LineRuleType.EXACT } } })),
        ...(textBoxes.some(({frame}) => !frame.drawingGroup)
          ? [textBoxAnchorStyle]
          : []),
        ...(hasTableEnd ? [{ id: TABLE_END_STYLE, name: 'Table end', next: 'Normal', semiHidden: true, run: { size: 2 },
          paragraph: { spacing: { before: 0, after: 0, line: 20, lineRule: LineRuleType.EXACT } } }] : []),
      ] },
    sections,
  });
  const blob = await Packer.toBlob(docx);
  const explicitNoWrapTrailSpaces = document.wordWrapTrailSpaces === false;
  const exactPaper = groups.map(({ page }) => {
    if (page.size !== 'Custom') return;
    const [width, height] = activeDocxPrintPaper(page) ?? paperSize(page);
    if (width < 25 || height < width || height > 1320 ||
      [width, height].every(mm => Math.abs(mm - mmToTwip(mm) * twipToMm) < 1e-9)) return;
    return [width, height];
  });
  const hasExactPaper = exactPaper.some(Boolean);
  const exactMargins = groups.map(({ page }) => parseHwpHorizontalMargins([
    activeHwpHorizontalMargin(page, 'left') ?? null,
    activeHwpHorizontalMargin(page, 'right') ?? null,
  ]));
  const hasExactMargins = exactMargins.some(Boolean);
  if (!bookmarkNames.size && !document.comments.length && !geometries.size && !textBoxes.length &&
      groups.length === 1 && !page.wordBackground && !explicitNoWrapTrailSpaces && !hasExactPaper && !hasExactMargins && !hwpFootnoteLayouts.size) return blob;
  // docx 9.12 emits resolution metadata only for replies. Standalone comments
  // need the same standard OOXML part; use its generated paragraph IDs.
  const zip = await JSZip.loadAsync(await blob.arrayBuffer());
  await writeHwpFootnoteLayouts(zip, hwpFootnoteLayouts);
  if (page.wordBackground || explicitNoWrapTrailSpaces) {
    const settings = xml(await zip.file('word/settings.xml')!.async('string'));
    if (page.wordBackground) {
      const main = xml(await zip.file('word/document.xml')!.async('string'));
      writeWordBackground(main, settings, page.wordBackground);
      zip.file('word/document.xml', new XMLSerializer().serializeToString(main));
    }
    if (explicitNoWrapTrailSpaces) {
      // The writer omits false compatibility flags. Restore the authored off
      // value after the generated legacy flags and before the mode setting.
      const compat = settings.getElementsByTagNameNS(W, 'compat')[0];
      const flag = settings.createElementNS(W, 'w:wrapTrailSpaces');
      flag.setAttributeNS(W, 'w:val', 'false');
      compat.insertBefore(flag, compat.getElementsByTagNameNS(W, 'compatSetting')[0] ?? null);
    }
    zip.file('word/settings.xml', new XMLSerializer().serializeToString(settings));
  }
  if (bookmarkNames.size) {
    // Web Word can replace a collapsed named destination with a copy of its
    // following text, duplicating a heading. Give authored point bookmarks a
    // real following run as their range; retain their exact start position,
    // text, formatting and link targets. End-of-paragraph destinations remain
    // collapsed when there is no following text. Generated textbox carriers
    // are intentionally excluded: writeTextBoxes needs those exact markers.
    for (const name of Object.keys(zip.files).filter(name => /^word\/(?:document|header\d+|footer\d+|footnotes|endnotes|comments)\.xml$/.test(name))) {
      const part = xml(await zip.file(name)!.async('string'));
      const ends = new Map(elements(part, 'bookmarkEnd').map(end => [attr(end, 'id'), end]));
      let changed = false;
      for (const start of elements(part, 'bookmarkStart')) {
        if (!bookmarkNames.has(attr(start, 'name') || '')) continue;
        const end = ends.get(attr(start, 'id'));
        if (!end || start.nextElementSibling !== end || start.parentElement?.localName !== 'p') continue;
        let next = end.nextElementSibling;
        while (next && ['bookmarkStart', 'bookmarkEnd', 'commentRangeStart', 'commentRangeEnd'].includes(next.localName)) next = next.nextElementSibling;
        if (!next || !['r', 'hyperlink'].includes(next.localName)) continue;
        const textRuns = next.localName === 'r' ? [next] : Array.from(next.children);
        // Never enlarge a destination over pictures, fields, tabs or breaks.
        // A later caption is not the point's following text run when a real
        // object/control stands between them. Mixed hyperlink runs count too.
        if (!textRuns.length || !textRuns.every(run => run.localName === 'r' &&
          Array.from(run.children).every(child => ['rPr', 't'].includes(child.localName))) ||
          !textRuns.some(run => elements(run, 't').some(text => text.textContent?.trim().length))) continue;
        next.parentNode!.insertBefore(end, next.nextSibling);
        changed = true;
      }
      if (changed) zip.file(name, new XMLSerializer().serializeToString(part));
    }
  }
  if (groups.length > 1 || hasExactPaper || hasExactMargins) {
    const body = xml(await zip.file('word/document.xml')!.async('string'));
    if (hasExactPaper) {
      const sectionProperties = elements(body, 'sectPr');
      if (sectionProperties.length === exactPaper.length) sectionProperties.forEach((section, index) => {
        const dimensions = exactPaper[index], size = elements(section, 'pgSz')[0];
        if (!dimensions || !size) return;
        const twips = ['w', 'h'].map(key => Number(attr(size, key))).sort((a, b) => a - b);
        if (!dimensions.every((mm, axis) => mmToTwip(mm) === twips[axis])) return;
        size.setAttributeNS(xmlNamespaces, 'xmlns:wspg', exactPaperNamespace);
        size.setAttributeNS(xmlNamespaces, 'xmlns:mc', markupCompatibility);
        size.setAttributeNS(markupCompatibility, 'mc:Ignorable', 'wspg');
        size.setAttributeNS(markupCompatibility, 'mc:PreserveAttributes', 'wspg:sizeMm');
        size.setAttributeNS(exactPaperNamespace, 'wspg:sizeMm', dimensions.join(' '));
      });
    }
    if (hasExactMargins) {
      const sectionProperties = elements(body, 'sectPr');
      if (sectionProperties.length === exactMargins.length) sectionProperties.forEach((section, index) => {
        const original = exactMargins[index], margins = elements(section, 'pgMar')[0];
        if (!original || !margins || !original.every((hu, side) => hu === null || Math.round(hu / 5) === Number(attr(margins, side === 0 ? 'left' : 'right')))) return;
        margins.setAttributeNS(xmlNamespaces, 'xmlns:wspg', exactPaperNamespace);
        margins.setAttributeNS(xmlNamespaces, 'xmlns:mc', markupCompatibility);
        margins.setAttributeNS(markupCompatibility, 'mc:Ignorable', 'wspg');
        margins.setAttributeNS(markupCompatibility, 'mc:PreserveAttributes', 'wspg:hwpHorizontal');
        margins.setAttributeNS(exactPaperNamespace, 'wspg:hwpHorizontal', original.map(hu => hu ?? '-').join(' '));
      });
    }
    // docx adds Paragraph({}) when a section ends with a table. Mark only
    // those generated carriers, never an actual paragraph from the model.
    for (const p of groups.length > 1 ? elements(body, 'p') : []) {
      if (p.parentElement?.localName !== 'body' || p.children.length !== 1) continue;
      const properties = p.firstElementChild!;
      if (properties.localName !== 'pPr' || properties.children.length !== 1 || properties.firstElementChild?.localName !== 'sectPr') continue;
      const style = body.createElementNS(W, 'w:pStyle'); style.setAttributeNS(W, 'w:val', SECTION_BOUNDARY_STYLE); properties.prepend(style);
      const spacing = body.createElementNS(W, 'w:spacing');
      for (const [name, value] of Object.entries({ before: '0', after: '0', line: '1', lineRule: 'exact' })) spacing.setAttributeNS(W, `w:${name}`, value);
      properties.insertBefore(spacing, properties.lastElementChild);
    }
    zip.file('word/document.xml', new XMLSerializer().serializeToString(body));
  }
  await writeTextBoxes(zip, textBoxes);
  await restoreImageGeometry(zip, geometries);
  await writeDrawingGroups(zip);
  if (document.comments.length) {
    const commentsXml = xml(await zip.file('word/comments.xml')!.async('string'));
    const extended = document.comments.map((comment, index) => {
      const el = elements(commentsXml, 'comment').find(c => attr(c, 'id') === String(index));
      const paragraph = el && elements(el, 'p').at(-1);
      const paraId = paragraph && Array.from(paragraph.attributes).find(a => a.localName === 'paraId')?.value;
      return paraId ? `<w15:commentEx w15:paraId="${paraId}" w15:done="${comment.resolved ? 1 : 0}"/>` : '';
  }).join('');
  zip.file('word/commentsExtended.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">${extended}</w15:commentsEx>`);
  const relsPath = 'word/_rels/document.xml.rels';
  zip.file(relsPath, (await zip.file(relsPath)!.async('string')).replace('</Relationships>', '<Relationship Id="rIdStudioCommentsExtended" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/></Relationships>'));
  zip.file('[Content_Types].xml', (await zip.file('[Content_Types].xml')!.async('string')).replace('</Types>', '<Override PartName="/word/commentsExtended.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"/></Types>'));
  }
  await encodeDocxXmlParts(zip);
  return zip.generateAsync({ type: 'blob', compression: 'DEFLATE', compressionOptions: { level: 6 }, mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
}
