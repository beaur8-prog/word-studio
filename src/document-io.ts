import { sanitizeImportedHtml, validImage, safeLink, RASTER_DATA } from "./document-html";
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
  Bookmark,
  Footer,
  Header,
  HeadingLevel,
  ImageRun,
  LevelFormat,
  LineRuleType,
  Packer,
  PageBreak,
  PageNumber,
  PageOrientation,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  CommentRangeStart, CommentRangeEnd, CommentReference,
  UnderlineType,
  WidthType,
  VerticalAlignTable,
  HeightRule,
  Tab,
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
import { defaultPage, pageMargins, basePageMargins } from "./types";
import { prepareDocxMetadata } from './docx-metadata';
import { normalizeDocxTables } from './docx-table-grid';
import { parseTabStops } from './document-tabs';
import { overlapParts } from './word-overlap';
import { rowHeight, MAX_ROW_HEIGHT } from './table-row-resize';
import { storyHtml, documentNotes, attachDocumentNotes, structureDocxSections } from './docx-stories';
import { parseDocumentHtml } from './extensions';
import { imageFloating, restoreImageGeometry, type ImageGeometryRecord } from './image-floating';
import type { PageSettings, StudioDocument } from "./types";

export interface ImportResult {
  document: StudioDocument;
  html?: string;
  warnings: string[];
}
export const DOCX_EXPORT_NOTICE =
  "DOCX에는 본문 서식, 표, 그림, 구역 설정과 검토 메모를 저장합니다. 편집 상태의 백업에는 Word Studio 형식도 함께 사용하세요.";
const MAX_FILE_BYTES = 40 * 1024 * 1024;
const MAX_INFLATED_BYTES = 100 * 1024 * 1024;
const INLINE = new Set(["text", "hardBreak", "inlineImage", "documentTab", "wordOverlap", "documentBookmark", "noteReference", "documentField", "documentSymbol"]);
const BLOCKS = new Set([
  "paragraph",
  "heading",
  "blockquote",
  "bulletList",
  "orderedList",
  "codeBlock",
  "horizontalRule",
  "pageBreak",
  "sectionBreak",
  "image",
  "table",
  "documentNote", "columnGroup", "sectionContent",
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
  "commentAnchor",
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
    ["id", "format", "version", "title", "content", "page", "comments", "updatedAt"],
    "문서",
  );
  if (input.format !== "word-studio" || input.version !== 1)
    fail("지원하지 않는 Word Studio 문서 버전입니다.");
  if (input.id !== undefined && (typeof input.id !== 'string' || !/^[a-zA-Z0-9_-]{1,128}$/.test(input.id))) fail('문서 ID가 올바르지 않습니다.');
  if (!boundedString(input.title, 512) || !validDate(input.updatedAt))
    fail("문서 제목 또는 저장 날짜가 올바르지 않습니다.");
  const page = input.page;
  if (!object(page)) fail("용지 설정이 없습니다.");
  keysOnly(
    page,
    ["size", "orientation", "margin", "margins", "header", "footer", "pageNumbers",
      'differentFirstPage', 'differentOddEven', 'firstHeader', 'firstFooter', 'evenHeader', 'evenFooter',
      'firstPageNumbers', 'evenPageNumbers', 'headerDistance', 'footerDistance', 'gutter', 'gutterTop', 'breakType', 'columns',
      'headerHtml','footerHtml','firstHeaderHtml','firstFooterHtml','evenHeaderHtml','evenFooterHtml'],
    "용지 설정",
  );
  if (
    !["A4", "Letter"].includes(String(page.size)) ||
    !["portrait", "landscape"].includes(String(page.orientation)) ||
    !finite(page.margin, 0, 75) ||
    !boundedString(page.header, 10000) ||
    !boundedString(page.footer, 10000) ||
    typeof page.pageNumbers !== "boolean"
  )
    fail("용지 설정이 올바르지 않습니다.");
  if (page.margins !== undefined) {
    if (!object(page.margins)) fail('방향별 여백이 올바르지 않습니다.');
    keysOnly(page.margins, ['top', 'right', 'bottom', 'left'], '방향별 여백');
    if (['top', 'right', 'bottom', 'left'].some(k => !finite((page.margins as Record<string, unknown>)[k], 0, 75))) fail('방향별 여백이 올바르지 않습니다.');
  }
  for (const key of ['differentFirstPage', 'differentOddEven', 'firstPageNumbers', 'evenPageNumbers', 'gutterTop'])
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
  const walk = (value: unknown, parent: string | null, depth: number): void => {
    if (++nodes > 100000 || depth > 64) fail("문서 구조가 너무 큽니다.");
    if (!object(value) || typeof value.type !== "string")
      fail("문서 본문 구조가 올바르지 않습니다.");
    keysOnly(value, ["type", "attrs", "content", "marks", "text"], "본문 요소");
    const type = value.type;
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
        if (typeof attr === "string" && key !== "src" && attr.length > ((type === 'sectionBreak' || type === 'sectionContent') && key === 'page' ? MAX_FILE_BYTES : 10000))
          fail("본문 속성이 너무 깁니다.");
        if (typeof attr === "number" && !Number.isFinite(attr))
          fail("본문 수치 속성이 올바르지 않습니다.");
      }
    }
    const attrs = value.attrs as Record<string, unknown> | undefined;
    if (type === 'documentSymbol' && (!boundedString(attrs?.font, 100) || !attrs?.font || !/^\d{1,5}$/.test(String(attrs?.code)) || !finite(Number(attrs?.code), 0, 65535) || !boundedString(attrs?.glyph, 100))) fail('기호 글자가 올바르지 않습니다.');
    if (type === 'documentBookmark' && (!boundedString(attrs?.name, 1000) || !attrs?.name)) fail('책갈피 이름이 올바르지 않습니다.');
    if (type === 'documentField' && !['PAGE','NUMPAGES'].includes(String(attrs?.kind))) fail('페이지 필드가 올바르지 않습니다.');
    if ((type === 'documentNote' || type === 'noteReference') && (!boundedString(attrs?.id, 128) || !attrs?.id || !['footnote','endnote'].includes(String(attrs?.kind)) || !boundedString(attrs?.label, 100))) fail('각주·미주 정보가 올바르지 않습니다.');
    if (type === 'columnGroup' && (!finite(attrs?.count, 1, 12) || !Number.isInteger(attrs?.count) || !finite(attrs?.space, 0, 1000) || (attrs?.balance !== undefined && typeof attrs.balance !== 'boolean'))) fail('다단 본문 설정이 올바르지 않습니다.');
    if (type === 'wordOverlap' && !overlapParts(attrs?.parts).length) fail('겹친 글자 형식이 올바르지 않습니다.');
    if (type === 'tableRow' && attrs?.rowHeight != null && !rowHeight(attrs.rowHeight))
      fail(`표 행 높이는 0보다 크고 ${MAX_ROW_HEIGHT}px 이하여야 합니다.`);
    if (type === 'sectionBreak' || type === 'sectionContent') {
      if (parent !== 'doc') fail('구역 나누기는 문서의 최상위에 있어야 합니다.');
      if (typeof attrs?.page !== 'string' || attrs.page.length > MAX_FILE_BYTES) fail('구역 설정이 올바르지 않습니다.');
      const sectionPage = JSON.parse(attrs.page);
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
      "sectionBreak",
      "image",
      "inlineImage",
      "documentTab",
      "wordOverlap",
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
        "documentNote", "columnGroup", "sectionContent",
      ].includes(type) &&
      children.length === 0
    )
      fail("필수 본문 요소가 비어 있습니다.");
    if (
      type === "listItem" &&
      (!object(children[0]) || children[0].type !== "paragraph")
    )
      fail("목록 항목에는 첫 문단이 필요합니다.");
    children.forEach((child) => walk(child, type, depth + 1));
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
          result += elements(run, "t")
            .map((t) => t.textContent || "")
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
    warnings.push("지원하지 않는 용지 크기를 A4로 표시합니다.");
  const margin = elements(section, "pgMar")[0];
  const enabled = (el: Element | undefined) => !!el && !['0', 'false', 'off'].includes(attr(el, 'val') || '');
  page.differentFirstPage = enabled(elements(section, 'titlePg')[0]);
  const settingsFile = zip.file('word/settings.xml');
  const settings = settingsFile ? xml(await settingsFile.async('string')) : undefined;
  if (page.columns && settings) page.columns.balance = !enabled(elements(settings, 'noColumnBalance')[0]);
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
  if (file.size > MAX_FILE_BYTES) fail("40MB 이하의 문서를 열어 주세요.");
  const extension = file.name.split(".").pop()?.toLowerCase();
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
    fail("압축을 푼 DOCX 크기가 100MB를 초과합니다.");
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
  doc.page = pages[0] ?? { ...defaultPage };
  const metadata = await prepareDocxMetadata(zip, body, pages);
  doc.comments = metadata.comments;
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
  const sanitized = sanitizeImportedHtml(structureDocxSections(attachDocumentNotes(metadata.decorate(result.value), notes), pages));
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
function fontPoints(value: unknown): number | undefined {
  const parsed = Number.parseFloat(String(value));
  if (!Number.isFinite(parsed) || parsed <= 0) return undefined;
  return Math.min(400, String(value).endsWith("px") ? parsed * 0.75 : parsed);
}

function cellBorders(attrs: Record<string, unknown> = {}): ITableCellBorders {
  return Object.fromEntries(['top', 'right', 'bottom', 'left'].map(side => {
    const value = attrs[`border${side[0].toUpperCase()}${side.slice(1)}`];
    const css = document.createElement('span').style;
    css.border = typeof value === 'string' ? value : '1px solid #c6ced8';
    const styles = { solid: BorderStyle.SINGLE, dashed: BorderStyle.DASHED, dotted: BorderStyle.DOTTED, double: BorderStyle.DOUBLE };
    return [side, {
      style: ['none', 'hidden'].includes(css.borderStyle) ? BorderStyle.NIL : styles[css.borderStyle as keyof typeof styles] ?? BorderStyle.SINGLE,
      size: Math.max(0, Math.round((fontPoints(css.borderWidth) ?? .75) * 8 / (css.borderStyle === 'double' ? 3 : 1))), color: rgb(css.borderColor) ?? 'C6CED8',
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
    return node.content ? { ...node, content: inheritCellText(node.content, attrs) } : node;
  });
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
        const size = fontPoints(mark.attrs?.fontSize);
        if (size) options.size = Math.round(size * 2);
        const color = rgb(mark.attrs?.color);
        if (color) options.color = color;
        const background = rgb(mark.attrs?.backgroundColor);
        if (background) options.shading = { fill: background };
        const tracking = Number.parseFloat(mark.attrs?.letterSpacing);
        if (Number.isFinite(tracking)) options.characterSpacing = Math.round(tracking * (String(mark.attrs?.letterSpacing).endsWith('pt') ? 20 : 15));
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
function paragraphOptions(node: JSONContent): IParagraphOptions {
  const lineHeight =
    node.attrs?.lineSpacing ??
    node.attrs?.lineHeight ??
    node.content
      ?.flatMap((n) => n.marks ?? [])
      .find((m) => m.type === "textStyle")?.attrs?.lineHeight;
  const line = Number.parseFloat(lineHeight);
  const indent = Number(node.attrs?.indent);
  const right = Number.parseFloat(node.attrs?.rightIndent), first = Number.parseFloat(node.attrs?.firstLineIndent);
  const before = Number.parseFloat(node.attrs?.spaceBefore), after = Number.parseFloat(node.attrs?.spaceAfter);
  const exact = typeof lineHeight === 'string' && /(?:pt|px)$/.test(lineHeight);
  return {
    alignment: alignment(node.attrs?.textAlign),
    tabStops: parseTabStops(node.attrs?.tabStops).map(stop => ({ position: Math.round(stop.position * 15), type: stop.type,
      ...(['dot', 'hyphen', 'underscore'].includes(stop.leader || '') ? { leader: stop.leader } : {}) })),
    indent: {
      ...(Number.isFinite(indent) ? { left: Math.round(indent * 15) } : {}),
      ...(Number.isFinite(right) ? { right: Math.round(right * 15) } : {}),
      ...(Number.isFinite(first) ? first < 0 ? { hanging: Math.round(-first * 15) } : { firstLine: Math.round(first * 15) } : {}),
    },
    run: { ...runOptions({ type: 'text', marks: [{ type: 'textStyle', attrs: {
      fontFamily: node.attrs?.paragraphFont, fontSize: node.attrs?.paragraphSize, color: node.attrs?.paragraphColor,
    } }] }), text: undefined },
    spacing: {
      after: Number.isFinite(after) ? Math.round(after * 15) : 120,
      ...(Number.isFinite(before) ? { before: Math.round(before * 15) } : {}),
      ...(exact && Number.isFinite(line) && line > 0 ? { line: Math.round(line * (lineHeight.endsWith('pt') ? 20 : 15)), lineRule: node.attrs?.wordLineRule === 'atLeast' ? LineRuleType.AT_LEAST : LineRuleType.EXACT } : {}),
      ...(!exact && Number.isFinite(line) && line > 0 && line < 10
        ? { line: Math.round(line * 240), lineRule: LineRuleType.AUTO }
        : {}),
    },
    widowControl: true,
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
  const scale = Math.min(1, availablePixels / width, 950 / height);
  width = Math.max(1, Math.round(width * scale));
  height = Math.max(1, Math.round(height * scale));
  const floating = imageFloating(node.attrs?.floating), name = String(node.attrs?.title || '그림');
  const geometryId = floating?.polygon ? `wsFloat_${crypto.randomUUID()}` : undefined;
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
  const page = document.page;
  const geometries = new Map<string, ImageGeometryRecord>();
  const mmToTwip = (mm: number) => Math.round((mm * 1440) / 25.4);
  const portraitWidth = page.size === "Letter" ? 215.9 : 210;
  const portraitHeight = page.size === "Letter" ? 279.4 : 297;
  let contentWidth = mmToTwip(
    (page.orientation === "landscape" ? portraitHeight : portraitWidth) -
      pageMargins(page).left - pageMargins(page).right,
  );
  const numbering: NumberingDefinition[] = [];
  const configuredLists = new Set<string>();
  const listLevels = new Map<string, Map<number, Record<string, unknown>>>();
  const notes = new Map<string, JSONContent>();
  const scan = (node: JSONContent) => {
    if (node.type === 'documentNote') notes.set(String(node.attrs?.id), node);
    if (node.type === 'listItem' && node.attrs?.wordListId) {
      const id = String(node.attrs.wordListId), levels = listLevels.get(id) || new Map();
      levels.set(Number(node.attrs.wordListLevel) || 0, node.attrs); listLevels.set(id, levels);
    }
    node.content?.forEach(scan);
  }; scan(document.content);
  const noteIds = new Map([...notes.keys()].map((id, index) => [id, index + 1]));
  let listId = 0;
  const commentIds = new Map(document.comments.map((comment, i) => [comment.id, i]));
  const remaining = new Map<string, number>();
  const anchored = (node: JSONContent): string[] => (node.marks?.find(m => m.type === 'commentAnchor')?.attrs?.ids || '').split(' ').filter((id: string) => commentIds.has(id));
  const count = (node: JSONContent) => {
    if (node.type === 'text') for (const id of anchored(node)) remaining.set(id, (remaining.get(id) ?? 0) + 1);
    node.content?.forEach(count);
  };
  count(document.content);
  const opened = new Set<string>();
  let wroteUnanchored = false;
  const inline = async (nodes: JSONContent[]): Promise<ParagraphChild[]> => {
    const runs: ParagraphChild[] = [];
    if (!wroteUnanchored) {
      wroteUnanchored = true;
      for (const [id, numeric] of commentIds) if (!remaining.has(id))
        runs.push(new CommentRangeStart(numeric), new CommentRangeEnd(numeric), new TextRun({ children: [new CommentReference(numeric)] }));
    }
    for (const node of nodes) {
      if (node.type === "text") {
        const ids = anchored(node);
        for (const id of ids) if (!opened.has(id)) { runs.push(new CommentRangeStart(commentIds.get(id)!)); opened.add(id); }
        const run = new TextRun(runOptions(node));
        const link = node.marks?.find((m) => m.type === "link");
        if (
          link &&
          safeLink(link.attrs?.href)
        )
          runs.push(
            link.attrs.href.startsWith('#')
              ? new InternalHyperlink({ children: [run], anchor: link.attrs.href.slice(1) })
              : new ExternalHyperlink({ children: [run], link: link.attrs.href }),
          );
        else runs.push(run);
        for (const id of ids) {
          remaining.set(id, remaining.get(id)! - 1);
          if (!remaining.get(id)) runs.push(new CommentRangeEnd(commentIds.get(id)!), new TextRun({ children: [new CommentReference(commentIds.get(id)!)] }));
        }
      } else if (node.type === "hardBreak")
        runs.push(new TextRun({ break: 1 }));
      else if (node.type === 'documentTab') runs.push(new TextRun({ children: [new Tab()] }));
      else if (node.type === 'documentBookmark') runs.push(new Bookmark({ id: String(node.attrs?.name || ''), children: [] }));
      else if (node.type === 'documentField') runs.push(new TextRun({ children: [node.attrs?.kind === 'NUMPAGES' ? PageNumber.TOTAL_PAGES : PageNumber.CURRENT] }));
      else if (node.type === 'documentSymbol') runs.push(new SymbolRun({ ...runOptions(node), symbolfont: String(node.attrs?.font || 'Symbol'), char: Number(node.attrs?.code).toString(16).padStart(4, '0') }));
      else if (node.type === 'noteReference') {
        const id = noteIds.get(String(node.attrs?.id));
        if (id) runs.push(node.attrs?.kind === 'endnote' ? new EndnoteReferenceRun(id) : new FootnoteReferenceRun(id));
      }
      else if (node.type === 'wordOverlap') {
        const parts = overlapParts(node.attrs?.parts);
        const field = new SimpleField(`EQ \\o\\ac(${parts.join(',')})`);
        field.addChildElement(new TextRun({ ...runOptions(node), text: parts.join('') }));
        runs.push(field);
      }
      else if (node.type === "image" || node.type === "inlineImage")
        runs.push(await imageRun(node, contentWidth / 15, geometries));
    }
    return runs;
  };
  type Context = {
    depth?: number;
    indent?: number;
    cellRun?: IRunOptions;
    list?: { reference: string; level: number };
  };
  const blocks = async (
    nodes: JSONContent[],
    context: Context = {},
  ): Promise<(Paragraph | Table)[]> => {
    const output: (Paragraph | Table)[] = [];
    for (const node of nodes) {
      if (
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
        output.push(
          new Paragraph({
            ...paragraphOptions(node),
            children: await inline(node.content ?? []),
            ...(context.cellRun ? { run: { ...context.cellRun, ...paragraphOptions(node).run } } : {}),
            ...(node.type === "heading"
              ? { heading: headings[(node.attrs?.level ?? 1) - 1] }
              : {}),
            ...(node.type === "codeBlock"
              ? {
                  run: { font: "Consolas", size: 20 },
                  shading: { fill: "F3F4F6" },
                }
              : {}),
            ...(context.list ? { numbering: context.list } : {}),
            ...(context.indent ? { indent: { left: context.indent } } : {}),
          }),
        );
      } else if (node.type === "image")
        output.push(
          new Paragraph({
            children: [await imageRun(node, contentWidth / 15, geometries)],
            alignment: alignment(node.attrs?.textAlign),
          }),
        );
      else if (node.type === "pageBreak")
        output.push(new Paragraph({ children: [new PageBreak()] }));
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
        const imported = node.content?.[0]?.attrs;
        const reference = imported?.wordListId ? `studio-imported-${imported.wordListId}` : `studio-list-${++listId}`;
        const level = Math.min(Number(imported?.wordListLevel ?? context.depth ?? 0), 8);
        if (!configuredLists.has(reference)) numbering.push({
          reference,
          levels: Array.from({ length: 9 }, (_, i) => ({
            level: i,
            format: (listLevels.get(String(imported?.wordListId))?.get(i)?.wordListFormat as typeof LevelFormat.DECIMAL) || (
              node.type === "orderedList"
                ? LevelFormat.DECIMAL
                : LevelFormat.BULLET),
            text: String(listLevels.get(String(imported?.wordListId))?.get(i)?.wordListPattern || (
              node.type === "orderedList"
                ? `%${i + 1}.`
                : ["•", "◦", "▪"][i % 3])),
            start:
              node.type === "orderedList"
                ? Math.max(1, Number(node.attrs?.start) || 1)
                : 1,
            alignment: AlignmentType.LEFT,
            style: {
              paragraph: { indent: { left: 720 * (i + 1), hanging: 360 } },
            },
          })),
        });
        configuredLists.add(reference);
        for (const item of node.content ?? []) {
          let first = true;
          for (const child of item.content ?? []) {
            const isParagraph = ["paragraph", "heading"].includes(
              child.type ?? "",
            );
            output.push(
              ...(await blocks([child], {
                depth: level + 1,
                ...(first && isParagraph ? { list: { reference, level } } : {}),
                ...(!first && isParagraph ? { indent: 720 * (level + 1) } : {}),
              })),
            );
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
          contentWidth / Math.max(1, columnWidths.length),
        );
        for (let i = 0; i < columnWidths.length; i++)
          if (!columnWidths[i]) columnWidths[i] = fallback;
        const total = columnWidths.reduce((a, b) => a + b, 0);
        if (total > contentWidth)
          for (let i = 0; i < columnWidths.length; i++)
            columnWidths[i] = Math.max(
              1,
              Math.floor((columnWidths[i] * contentWidth) / total),
            );
        let continuations: { remaining: number; attrs: Record<string, unknown> }[] = [];
        const borderPixels = (attrs: Record<string, unknown>, side: string) => {
          const css = globalThis.document.createElement('span').style;
          css.border = String(attrs[`border${side}`] ?? '1px solid #c6ced8');
          return ['none', 'hidden'].includes(css.borderStyle) ? 0 : (fontPoints(css.borderWidth) ?? .75) * 4 / 3;
        };
        for (const row of node.content ?? []) {
          const cells: TableCell[] = [];
          // Subtract the same border reserve Word adds to minimum heights,
          // including continuation cells materialized by the DOCX writer.
          const rowBorder = Math.max(0,
            ...continuations.flatMap(c => ['Top', 'Bottom'].map(side => borderPixels(c.attrs, side))),
            ...(row.content ?? []).filter(c => Number(c.attrs?.rowspan || 1) === 1).map(c => borderPixels(c.attrs ?? {}, 'Bottom')));
          continuations = continuations.map(c => ({ ...c, remaining: c.remaining - 1 })).filter(c => c.remaining > 0);
          for (const cell of row.content ?? []) if (Number(cell.attrs?.rowspan) > 1)
            continuations.push({ remaining: Number(cell.attrs!.rowspan) - 1, attrs: cell.attrs ?? {} });
          const rowPadding = Math.max(0, ...(row.content ?? []).map(cell => ['Top', 'Bottom'].reduce((sum, side) => {
            const padding = parseFloat(cell.attrs?.[`padding${side}`]);
            return sum + (Number.isFinite(padding) ? Math.max(0, padding) : 8);
          }, 0)));
          for (const cell of row.content ?? []) {
            const children = await blocks(inheritCellText(cell.content ?? [], cell.attrs), {
              cellRun: runOptions({ type: 'text', marks: [{ type: 'textStyle', attrs: {
                fontFamily: cell.attrs?.fontFamily, fontSize: cell.attrs?.fontSize, color: cell.attrs?.textColor,
              } }] }),
            });
            if (
              !children.length ||
              children[children.length - 1] instanceof Table
            )
              children.push(new Paragraph(""));
            const color =
              cell.attrs?.backgroundColor === 'transparent' ? undefined : rgb(cell.attrs?.backgroundColor) ??
              (cell.type === "tableHeader" ? "E8EEF7" : undefined);
            cells.push(
              new TableCell({
                children,
                columnSpan: Number(cell.attrs?.colspan) || 1,
                rowSpan: Number(cell.attrs?.rowspan) || 1,
                ...(color || cell.attrs?.backgroundColor === 'transparent' ? { shading: { fill: color ?? 'auto' } } : {}),
                borders: cellBorders(cell.attrs),
                verticalAlign: cell.attrs?.verticalAlign === 'middle' ? VerticalAlignTable.CENTER : cell.attrs?.verticalAlign === 'bottom' ? VerticalAlignTable.BOTTOM : VerticalAlignTable.TOP,
                margins: Object.fromEntries(['top', 'right', 'bottom', 'left'].map(side => {
                  const padding = Number.parseFloat(cell.attrs?.[`padding${side[0].toUpperCase()}${side.slice(1)}`]);
                  return [side, Number.isFinite(padding) ? Math.round(Math.max(0, padding) * 15) : ['left', 'right'].includes(side) ? 150 : 120];
                })),
              }),
            );
          }
          rows.push(
            new TableRow({
              children: cells,
              ...(rowHeight(row.attrs?.rowHeight) ? { height: { value: Math.max(1, Math.round((row.attrs!.rowHeight - (row.attrs?.wordRowRule === 'exact' ? 0 : rowPadding + rowBorder)) * 15)), rule: row.attrs?.wordRowRule === 'exact' ? HeightRule.EXACT : HeightRule.ATLEAST } } : {}),
              ...((row.content?.length ?? 0) > 0 && (row.content ?? []).every(
                (cell) => cell.type === "tableHeader",
              ) ? { tableHeader: true } : {}),
            }),
          );
        }
        output.push(
          new Table({
            rows,
            width: { size: columnWidths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
            ...(node.attrs?.tableIndent ? { indent: { size: Math.round(Number.parseFloat(node.attrs.tableIndent) * 15) || 0, type: WidthType.DXA } } : {}),
            columnWidths,
          }),
        );
      }
    }
    return output;
  };
  const groups: { page: PageSettings; nodes: JSONContent[] }[] = [{ page, nodes: [] }];
  for (const node of document.content.content ?? []) {
    if (node.type === 'sectionBreak') groups.push({ page: JSON.parse(node.attrs!.page), nodes: [] });
    else groups[groups.length - 1].nodes.push(node);
  }
  const sections = [];
  const footnotes: Record<string, { children: Paragraph[] }> = {}, endnotes: Record<string, { children: Paragraph[] }> = {};
  for (const [id, note] of notes) {
    // Both paragraph and table objects implement the block XML component
    // protocol, although docx's public note type lists Paragraph only.
    const children = await blocks(note.content || []);
    (note.attrs?.kind === 'endnote' ? endnotes : footnotes)[String(noteIds.get(id))] = { children: children as Paragraph[] };
  }
  for (const group of groups) {
  const page = group.page;
  const portraitWidth = page.size === 'Letter' ? 215.9 : 210;
  const portraitHeight = page.size === 'Letter' ? 279.4 : 297;
  const margins = pageMargins(page);
  contentWidth = mmToTwip((page.orientation === 'landscape' ? portraitHeight : portraitWidth) - margins.left - margins.right);
  const children = await blocks(group.nodes);
  const richBlocks = async (html: string) => blocks(parseDocumentHtml(sanitizeImportedHtml(html).html).content || []);
  const footer = async (text: string, numbers: boolean, html?: string, headerHtml?: string) => new Footer({ children: [
    ...(html ? await richBlocks(html) : [new Paragraph({ text, alignment: AlignmentType.CENTER, run: { size: 18, color: '697386' } })]),
    ...(numbers && !/data-page-number/.test((html || '') + (headerHtml || '')) ? [new Paragraph({ alignment: AlignmentType.CENTER,
      children: [new TextRun({ children: [PageNumber.CURRENT], size: 18, color: '697386' })] })] : []),
  ] });
  const header = async (text: string, html?: string) => new Header({ children: html ? await richBlocks(html) : [new Paragraph({ text, run: { size: 18, color: '697386' } })] });
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
              header: mmToTwip(page.headerDistance ?? 6.35),
              footer: mmToTwip(page.footerDistance ?? 6.35),
            },
          },
        },
        headers: { default: await header(page.header, page.headerHtml),
          first: await header(page.firstHeader ?? '', page.firstHeaderHtml), even: await header(page.differentOddEven ? page.evenHeader ?? '' : page.header, page.differentOddEven ? page.evenHeaderHtml : page.headerHtml) },
        footers: { default: await footer(page.footer, page.pageNumbers, page.footerHtml, page.headerHtml),
          first: await footer(page.firstFooter ?? '', page.firstPageNumbers ?? false, page.firstFooterHtml, page.firstHeaderHtml),
          even: await footer(page.differentOddEven ? page.evenFooter ?? '' : page.footer, page.differentOddEven ? page.evenPageNumbers ?? false : page.pageNumbers, page.differentOddEven ? page.evenFooterHtml : page.footerHtml, page.differentOddEven ? page.evenHeaderHtml : page.headerHtml) },
      },
  );
  }
  const docx = new DocxDocument({
    title: document.title,
    creator: 'Word Studio',
    description: 'Created with Word Studio',
    evenAndOddHeaderAndFooters: groups.some(group => group.page.differentOddEven),
    compatibility: { noColumnBalance: groups.some(group => group.page.columns?.balance === false) },
    numbering: { config: numbering },
    footnotes, endnotes,
    comments: { children: await Promise.all(document.comments.map(async (comment, id) => ({ id, durableId: (id + 1).toString(16).padStart(8, '0').toUpperCase(), author: comment.author ?? 'Word Studio', initials: comment.initials, date: new Date(comment.createdAt),
      resolved: comment.resolved, children: comment.html ? await blocks(parseDocumentHtml(sanitizeImportedHtml(comment.html).html).content || []) : comment.text.split('\n').map(text => new Paragraph(text)) }))) },
    styles: { default: { document: { run: { font: '맑은 고딕', size: 22 }, paragraph: { spacing: { after: 120, line: 276 } } } } },
    sections,
  });
  const blob = await Packer.toBlob(docx);
  if (!document.comments.length && !geometries.size) return blob;
  // docx 9.12 emits resolution metadata only for replies. Standalone comments
  // need the same standard OOXML part; use its generated paragraph IDs.
  const zip = await JSZip.loadAsync(await blob.arrayBuffer());
  await restoreImageGeometry(zip, geometries);
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
  zip.file('[Content_Types].xml', (await zip.file('[Content_Types].xml')!.async('string')).replace('</Types>', '<Override PartName="/word/commentsExtended.xml" ContentType="application/vnd.ms-word.commentsExtended+xml"/></Types>'));
  }
  return zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
}
