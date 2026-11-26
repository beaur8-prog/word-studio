import { paperInset } from './paper-inset';
import { cellRelativePicture, reserveCellPicture, reservePictureMinimum, restoreCellPictureReservations, cellPictureReservation } from './floating-cell-relative';
import { documentScale } from './document-scale';
import { editorLayoutResume, editorLayoutSuspended } from './editor-layout-task';
import {inlineDrawingAnchor,prepareDrawingGroupAnchors} from './drawing-group';
import { layoutTextBoxes, textBoxFrame } from './textbox-frame';
import { documentMetadata } from './document-history';
import { wordFontMetricsChanged } from './word-font-fallbacks';
import { applyFloatingBottomBorder, clearFloatingBottomBorders, floatingBottomBorderDimensions, floatingBottomBorderPaint, wordBottomBorderEnd } from './word-paragraph-border-layout';
import { measuredBodySpaceAdjustment } from './word-line-metrics';
import { isOrdinaryWordBodyParagraph } from './word-trailing-spaces';
import { tableAnchor } from './table-anchor-data';
import { Extension } from "@tiptap/core";
import { Plugin, PluginKey } from "@tiptap/pm/state";
import type { Node as ProseMirrorNode } from "@tiptap/pm/model";
import { Decoration, DecorationSet, type EditorView } from "@tiptap/pm/view";
import { imageOutsidePaper, recoverableCellImage, imageFloating, type ImageFloating } from "./image-floating";

type Boundary = [globalThis.Node, number];
type Unit = {
  from: number;
  to: number;
  start: Boundary;
  end: Boundary;
  text: string;
  hard?: boolean;
  floating?: boolean;
};
type Widget = {
  offset: number;
  width?: number;
  break?: boolean;
  height?: number;
};
type Picture = {
  wrapCell?: HTMLElement | null;
  cellRelative?: boolean;
  minimumHeight?: number;
  hwpFullBodyWidth?: boolean;
  unconstrainedWidth?: boolean;
  hwpPaperCover?: boolean;
  offPage?: boolean;
  index: number;
  left: number;
  top: number;
  width: number;
  height: number;
  cssLeft: number;
  cssTop: number;
  cssTranslation?: string;
  geometry: ImageFloating;
};
type ParagraphPlan = { index: number; widgets: Widget[]; pictures: Picture[]; minimumHeight?: number; originalMinimum?: string; startPage?: boolean; wordSpaceAdjustment?: number; bottomBorderInset?: number };
type PageFlow = { starts: Set<number>; backwards: Set<number>; candidates: Set<number> };
type WordFloatingSpacePass = { skip: ReadonlySet<number>; rejected: Set<number> };

/** Read the anchor's paper in document pixels, including CSS zoom. */
function ownerPaper(element: HTMLElement, root: HTMLElement) {
  const paragraph = element.closest<HTMLElement>('p,h1,h2,h3,h4,h5,h6');
  if (!paragraph) return;
  const sheet = paragraph.closest<HTMLElement>('[data-page-height],.print-page') || root;
  const scale = documentScale(paragraph);
  const y = (paragraph.getBoundingClientRect().top - sheet.getBoundingClientRect().top) / scale;
  const height = Number(sheet.dataset.pageHeight) || 1122.52, top = paperInset(sheet.dataset.margin, 96);
  const page = Math.max(0, Math.floor((y - top + .5) / (height + 24)));
  return { page, y, height, top, sheet, scale, bottom: paperInset(sheet.dataset.bottomMargin, top) };
}

function reserveHwpBodyPicture(paragraph: HTMLElement, picture: Picture, root: HTMLElement) {
  const { geometry, height, top } = picture;
  const proof = geometry.hwpBodyFlow, next = paragraph.nextElementSibling;
  if (!proof || !picture.hwpFullBodyWidth || !(next instanceof HTMLParagraphElement) || next.textContent?.trim() ||
      next.querySelector('img,svg,br:not(.ProseMirror-trailingBreak),[data-note-reference]')) return;
  const paper = ownerPaper(paragraph, root);
  if (!paper || Math.abs(paper.scale - 1) > 1 / 64 || Math.abs(paper.y - paper.page * (paper.height + 24) - proof.anchorY) > .2) return;
  const sheet = paper.sheet.getBoundingClientRect();
  const nextStyle = getComputedStyle(next);
  if (Math.abs(parseFloat(nextStyle.fontSize) - proof.followingEm) > .06 ||
      Math.abs(parseFloat(nextStyle.marginTop) || 0) > .05) return;
  const target = top + height - proof.followingEm - paragraph.getBoundingClientRect().top;
  if (target > 0 && top + height <= sheet.top + (paper.page + 1) * paper.height + paper.page * 24 - paper.bottom + .16)
    reservePictureMinimum(paragraph, target);
}

type OwnerPageStep = { page: number; layout: number; pass: number };
type RepeatedOwner = {
  index: number;
  image: HTMLImageElement;
  paper: NonNullable<ReturnType<typeof ownerPaper>>;
  history: OwnerPageStep[];
};

function pageSquareBounds({ image, paper }: Pick<RepeatedOwner, "image" | "paper">) {
  const geometry = imageFloating(image.dataset.imageFloating);
  if (!geometry || geometry.drawingGroup || geometry.behindDocument || geometry.wrap?.type !== 1 ||
      geometry.wrap.side !== 'bothSides' || geometry.horizontalPosition.relative !== 'page' ||
      geometry.verticalPosition.relative !== 'page' ||
      geometry.horizontalPosition.align || geometry.verticalPosition.align ||
      (geometry.margins && Object.values(geometry.margins).some(value => value !== 0))) return;
  const horizontal = geometry.horizontalPosition.offset, vertical = geometry.verticalPosition.offset;
  if (horizontal === undefined || vertical === undefined) return;
  const x = horizontal / 9525, y = vertical / 9525;
  const width = Number(image.getAttribute('width')), height = Number(image.getAttribute('height'));
  if (!(width > 0 && height > 0) || x < 0 || x + width > paper.sheet.clientWidth ||
      y < paper.top || y + height > paper.height - paper.bottom) return;
  return { left: x, right: x + width, top: y, bottom: y + height };
}

function hwpPaperCoverEditorTop(picture: Picture, image: HTMLImageElement, root: HTMLElement) {
  if (!picture.hwpPaperCover) return;
  const paper = ownerPaper(image, root);
  if (!paper) return;
  // This picture wraps no text. Use the paper containing its live anchor;
  // the detached flow can place the same paragraph before a page seam.
  const stride = paper.height + 24;
  const start = Math.max(0, Math.floor((paper.y + .5) / stride)) * stride;
  if (paper.y >= start - .5 && paper.y < start + paper.height) return start - paper.y;
}

/** A source page containing only one drawing has no text to wrap. If the
 * detached flow places it on the preceding paper, paint on the already laid
 * out editor paper without changing its anchor, exclusion, or pagination. */
function isolatedPagePictureTop(picture: Picture, image: HTMLImageElement, view: EditorView, position: number) {
  const root = view.dom;
  if (picture.cssTop >= 0 || picture.offPage || picture.geometry.verticalPosition.relative !== 'page') return picture.cssTop;
  const paragraph = image.parentElement;
  if (paragraph?.tagName !== 'P' || !(paragraph.parentElement === root || paragraph.parentElement?.classList.contains('section-editable')))
    return picture.cssTop;
  const paper = ownerPaper(image, root);
  if (!paper?.page) return picture.cssTop;
  const stride = paper.height + 24, pageTop = paper.page * stride;
  const paintedTop = paper.y + picture.cssTop;
  if (paintedTop >= pageTop - .5 || paintedTop < pageTop - stride - .5 ||
      paintedTop + picture.height > pageTop - 24 + .5 || paper.y > pageTop + paper.height - paper.bottom + .5)
    return picture.cssTop;
  const bounds = pageSquareBounds({ image, paper });
  if (!bounds || paragraph.getClientRects().length !== 1 || parseFloat(getComputedStyle(paragraph).borderTopWidth)) return picture.cssTop;
  const anchor = view.state.doc.resolve(position), parent = anchor.parent, block = anchor.nodeAfter;
  if (block?.type.name !== 'paragraph' || block.childCount !== 1 || block.firstChild?.type.name !== 'inlineImage') return picture.cssTop;
  const boundedByPageBreak = (step: -1 | 1) => {
    let offset = step === -1 ? position : position + block.nodeSize;
    for (let index = anchor.index() + step; index >= 0 && index < parent.childCount; index += step) {
      const sibling = parent.child(index);
      if (step === -1) offset -= sibling.nodeSize;
      if (sibling.type.name === 'pageBreak') {
        const pageBreak = view.nodeDOM(offset);
        if (!(pageBreak instanceof HTMLElement)) return false;
        const box = pageBreak.getBoundingClientRect();
        // A break coalesced with a section start does not isolate a paper.
        if (box.height / paper.scale <= .5) return false;
        return step === 1 || Math.abs(
          (box.bottom - paper.sheet.getBoundingClientRect().top) / paper.scale - pageTop - paper.top
        ) <= .5;
      }
      if (sibling.type.name !== 'paragraph' || sibling.childCount) return false;
      if (step === 1) offset += sibling.nodeSize;
    }
    return false;
  };
  return boundedByPageBreak(-1) && boundedByPageBreak(1) ? pageTop + bounds.top - paper.y : picture.cssTop;
}

/** Two image-only anchors can share one repeated wrap state. When they
 * collide on the earlier paper, try moving the later source anchor first. */
function coupledOwnerPair(earlier: RepeatedOwner, later: RepeatedOwner, positions: Map<HTMLElement, number>, doc: ProseMirrorNode) {
  if (later.index !== earlier.index + 1 || earlier.paper.sheet !== later.paper.sheet) return false;
  const a = earlier.image.closest<HTMLElement>('p'), b = later.image.closest<HTMLElement>('p');
  if (!a || !b || a === b || a.parentElement !== b.parentElement) return false;
  const left = positions.get(a), right = positions.get(b);
  if (left === undefined || right === undefined || left >= right) return false;
  const from = doc.resolve(left), to = doc.resolve(right);
  if (from.parent !== to.parent) return false;
  for (const block of [from.nodeAfter, to.nodeAfter])
    if (!block || block.type.name !== 'paragraph' || block.childCount !== 1 || block.firstChild?.type.name !== 'inlineImage') return false;
  for (let index = from.index() + 1; index < to.index(); index++) {
    const block = from.parent.child(index);
    if (block.type.name !== 'paragraph' || block.childCount) return false;
  }

  // Require the same observations in the same layout passes. A coincident
  // report from a different cycle must retain its independent start decision.
  const first = earlier.history, second = later.history;
  if (!first.every((step, index) => step.pass === second[index].pass && step.layout === second[index].layout)) return false;
  const common = first[0].page < first[1].page ? 0 : 1, advanced = 1 - common;
  if (first[common].page !== second[common].page || first[advanced].page <= first[common].page ||
      first[advanced].page >= second[advanced].page) return false;
  const bounds = pageSquareBounds(earlier), other = pageSquareBounds(later);
  if (!bounds || !other) return false;
  // Edges that only meet can round in opposite directions when scaled from EMUs.
  const horizontal = Math.min(bounds.right, other.right) - Math.max(bounds.left, other.left);
  const vertical = Math.min(bounds.bottom, other.bottom) - Math.max(bounds.top, other.top);
  return horizontal > 1 / 64 && vertical > 1 / 64;
}

function layoutHash(value: string) {
  let hash = 2166136261;
  for (let index = 0; index < value.length; index++) hash = Math.imul(hash ^ value.charCodeAt(index), 16777619);
  return hash >>> 0;
}

const atomic =
  "img,[data-drawing-group-anchor],[data-symbol-font],[data-word-overlap],[data-docx-tab],[data-docx-bookmark],[data-note-reference],[data-page-number]";
const floatingImages = (root: HTMLElement) =>
  [...root.querySelectorAll<HTMLImageElement>("img[data-image-floating]")].filter(
    (img) => !img.closest(".document-furniture,.paper-backgrounds,.section-papers"),
  );
const paragraphs = (root: HTMLElement) =>
  [...root.querySelectorAll<HTMLElement>("p,h1,h2,h3,h4,h5,h6")].filter(
    (p) => !p.closest(".document-furniture,.paper-backgrounds,.section-papers"),
  );

function units(p: HTMLElement) {
  const result: Unit[] = [];
  let offset = 0;
  const visit = (node: globalThis.Node) => {
    if (node.nodeType === Node.TEXT_NODE) {
      for (const part of new Intl.Segmenter(undefined, {
        granularity: "grapheme",
      }).segment(node.textContent || "")) {
        result.push({
          from: offset,
          to: offset + part.segment.length,
          start: [node, part.index],
          end: [node, part.index + part.segment.length],
          text: part.segment,
        });
        offset += part.segment.length;
      }
    } else if (node instanceof HTMLElement) {
      if (
        node.hasAttribute("data-flow-widget") ||
        node.classList.contains("ProseMirror-separator") ||
        node.classList.contains("ProseMirror-trailingBreak")
      )
        return;
      if (node.matches(atomic) || node.tagName === "BR") {
        const index = [...node.parentNode!.childNodes].indexOf(node);
        result.push({
          from: offset,
          to: offset + 1,
          start: [node.parentNode!, index],
          end: [node.parentNode!, index + 1],
          text: node.tagName === "BR" ? "\n" : "\ufffc",
          hard: node.tagName === "BR",
          floating: node.matches("img[data-image-floating]"),
        });
        offset++;
      } else node.childNodes.forEach(visit);
    }
  };
  p.childNodes.forEach(visit);
  return result;
}
function widgetDOM(widget: Widget) {
  const el = document.createElement(widget.break ? "br" : "span");
  el.dataset.flowWidget = "true";
  el.contentEditable = "false";
  el.setAttribute("aria-hidden", "true");
  if (!widget.break)
    Object.assign(
      el.style,
      widget.height !== undefined
        ? { display: "block", height: `${widget.height}px`, lineHeight: "0" }
        : {
            display: "inline-block",
            width: `${Math.max(0, widget.width || 0)}px`,
            height: "0",
            verticalAlign: "baseline",
            pointerEvents: "none",
          },
    );
  return el;
}
function insertWidgets(p: HTMLElement, widgets: Widget[]) {
  const index = units(p);
  for (const { widget } of widgets
    .map((widget, order) => ({ widget, order }))
    .sort((a, b) => b.widget.offset - a.widget.offset || b.order - a.order)) {
    const unit = index.find((u) => u.from === widget.offset);
    const boundary = unit?.start || index.at(-1)?.end;
    if (!boundary) continue;
    const range = document.createRange();
    range.setStart(...boundary);
    range.collapse(true);
    range.insertNode(widgetDOM(widget));
  }
}

/** Return the part of a polygon intersected by an entire text line's ink band.
 * Sampling its vertices as well as both band edges prevents sloped contours
 * from cutting through letters between the sampled baselines. */
function exclusion(picture: Picture, y: number, height: number): [number, number][] {
  const { geometry: g, left, top, width, height: ph } = picture,
    m = g.margins;
  const y1 = y - (m?.bottom || 0) / 9525,
    y2 = y + height + (m?.top || 0) / 9525;
  if (g.wrap?.type === 0 || y2 <= top + 0.1 || y1 >= top + ph - 0.1) return [];
  if (g.wrap?.type === 3) return [[-1e7, 1e7]];
  let a = left,
    b = left + width;
  if ((g.wrap?.type === 2 || g.wrap?.type === 4) && g.polygon) {
    const low = Math.max(0, (y1 - top) / ph),
      high = Math.min(1, (y2 - top) / ph),
      values: number[] = [];
    const ys = [low, high, ...g.polygon.map((p) => p[1]).filter((v) => v > low && v < high)];
    for (const scan of ys)
      for (let i = 0; i < g.polygon.length; i++) {
        const p = g.polygon[i],
          q = g.polygon[(i + 1) % g.polygon.length];
        if (Math.abs(p[1] - q[1]) < 1e-8) {
          if (Math.abs(scan - p[1]) < 1e-6) values.push(p[0], q[0]);
        } else if (scan >= Math.min(p[1], q[1]) - 1e-7 && scan <= Math.max(p[1], q[1]) + 1e-7)
          values.push(p[0] + ((q[0] - p[0]) * (scan - p[1])) / (q[1] - p[1]));
      }
    if (!values.length) return [];
    a = left + Math.min(...values) * width;
    b = left + Math.max(...values) * width;
  }
  return [[a - (m?.left || 0) / 9525, b + (m?.right || 0) / 9525]];
}
/** A Word group excludes text as one drawing, including gaps between children. */
function groupExclusion(picture: Picture):Picture|undefined {
  const g=picture.geometry.drawingGroup;if(!g||picture.offPage)return;
  return {wrapCell:picture.wrapCell,index:-100000-picture.index,left:picture.left-((picture.geometry.horizontalPosition.offset||0)/9525-g.x),
    top:picture.top-((picture.geometry.verticalPosition.offset||0)/9525-g.y),width:g.width,height:g.height,cssLeft:0,cssTop:0,
    geometry:{horizontalPosition:{relative:'page',offset:g.x*9525},verticalPosition:{relative:g.vertical,offset:g.y*9525},wrap:g.wrap,
      ...(g.polygon?{polygon:g.polygon}:{}),drawingGroup:g,margins:{top:g.margins[0]*9525,right:g.margins[1]*9525,bottom:g.margins[2]*9525,left:g.margins[3]*9525}}};
}
function slots(pictures: Picture[], x: number, y: number, width: number, height: number) {
  let spans: [number, number][] = [[x, x + width]];
  for (const picture of pictures)
    for (let [a, b] of exclusion(picture, y, height)) {
      const side = picture.geometry.wrap?.side;
      if (side === "left") b = 1e7;
      if (side === "right") a = -1e7;
      if (side === "largest") {
        if (a - x >= x + width - b) b = 1e7;
        else a = -1e7;
      }
      spans = spans.flatMap(([l, r]) =>
        b <= l || a >= r
          ? [[l, r] as [number, number]]
          : ([
              [l, Math.min(r, a)],
              [Math.max(l, b), r],
            ].filter(([l, r]) => r - l >= 4) as [number, number][]),
      );
    }
  return spans;
}
function imagePosition(
  img: HTMLImageElement,
  index: number,
  root: HTMLElement,
): Picture | undefined {
  const offPage=imageOutsidePaper(img);
  const g = imageFloating(img.dataset.imageFloating);
  if (!g) return;
  const p = img.closest<HTMLElement>("p,h1,h2,h3,h4,h5,h6");
  if (!p) return;
  const recovery = recoverableCellImage(img);
  const cellFrame = cellRelativePicture(img, root);
  const cellTop = cellFrame ? reserveCellPicture(cellFrame) : undefined;
  // A cell-contained drawing reserves its authored extent without changing the
  // document's coordinates or forcing the picture into the old row height.
  const minimumHeight = recovery ? Math.max(
    parseFloat(p.style.minHeight) || 0,
    (g.verticalPosition.offset || 0) / 9525 + (Number(img.getAttribute("height")) || 0),
  ) : undefined;
  if (minimumHeight) p.style.minHeight = `${minimumHeight}px`;
  p.style.position = "relative";
  const rect = p.getClientRects()[0] || p.getBoundingClientRect(),
    group = p.closest<HTMLElement>("[data-column-count]");
  const sheet = p.closest<HTMLElement>("[data-page-height],.print-page") || root;
  const sr = sheet.getBoundingClientRect(),
    width = Number(img.getAttribute("width")) || img.width,
    height = Number(img.getAttribute("height")) || img.height;
  const ph = Number(sheet.dataset.pageHeight) || parseFloat(sheet.style.height) || 1122.52;
  const textRoot = p.closest<HTMLElement>(".section-editable") || root;
  const rootStyle = getComputedStyle(textRoot),
    top = paperInset(sheet.dataset.margin,
      parseFloat(rootStyle.paddingTop) || parseFloat(root.style.top) || 96);
  const bottom = paperInset(sheet.dataset.bottomMargin, parseFloat(rootStyle.paddingBottom) || 96);
  const page = Math.max(0, Math.floor((rect.top - sr.top - top + 0.5) / (ph + 24))),
    py = sr.top + page * (ph + 24);
  const leftMargin = paperInset(sheet.dataset.leftMargin, parseFloat(rootStyle.paddingLeft) || parseFloat(root.style.left) || 0);
  const rightMargin = paperInset(sheet.dataset.rightMargin, parseFloat(rootStyle.paddingRight) || leftMargin);
  const cell = p.closest<HTMLElement>("td,th");
  const cellStyle = cell ? getComputedStyle(cell) : undefined;
  const column = cell ? cell.getBoundingClientRect().left + (parseFloat(cellStyle!.paddingLeft) || 0) : group
    ? group.getBoundingClientRect().left +
      Math.floor(
        (rect.left - group.getBoundingClientRect().left + 0.5) /
          ((group.clientWidth + Number(group.dataset.columnSpace)) /
            Number(group.dataset.columnCount)),
      ) *
        ((group.clientWidth + Number(group.dataset.columnSpace)) /
          Number(group.dataset.columnCount))
    : sr.left + leftMargin;
  const horizontal = g.horizontalPosition,
    vertical = g.verticalPosition;
  const hbase =
    horizontal.relative === "page"
      ? sr.left
      : horizontal.relative === "column"
        ? column
        : sr.left + leftMargin;
  const hwidth =
    horizontal.relative === "page"
      ? sheet.clientWidth
      : horizontal.relative === "column" && cell
        ? cell.getBoundingClientRect().width - (parseFloat(cellStyle!.paddingLeft) || 0) - (parseFloat(cellStyle!.paddingRight) || 0)
      : horizontal.relative === "column" && group
        ? (group.clientWidth -
            (Number(group.dataset.columnCount) - 1) * Number(group.dataset.columnSpace)) /
          Number(group.dataset.columnCount)
        : sheet.clientWidth - leftMargin - rightMargin;
  const vbase =
    vertical.relative === "page" ? py : vertical.relative === "margin" ? py + top : rect.top;
  const vh =
    vertical.relative === "page"
      ? ph
      : vertical.relative === "margin"
        ? ph - top - bottom
        : rect.height;
  const inline=g.drawingGroup?.position==='inline' ? inlineDrawingAnchor(root,g.drawingGroup.id)?.getBoundingClientRect():undefined;
  let left = inline ? inline.left+(horizontal.offset||0)/9525-g.drawingGroup!.x :
    hbase +
    (horizontal.align === "center"
      ? (hwidth - width) / 2
      : horizontal.align === "right"
        ? hwidth - width
        : (horizontal.offset || 0) / 9525);
  let y = inline ? inline.top+(vertical.offset||0)/9525-g.drawingGroup!.y :
    vbase +
    (vertical.align === "center"
      ? (vh - height) / 2
      : vertical.align === "bottom"
        ? vh - height
        : (vertical.offset || 0) / 9525);
  const paperEdgeColumnPicture = !offPage && isPaperEdgeColumnPicture(img, g, sheet, top);
  if (
    !paperEdgeColumnPicture && !offPage && g.wrap?.type !== 0 &&
    height <= ph - top - bottom &&
    !["paragraph", "line"].includes(vertical.relative || "paragraph")
  )
    y = Math.max(py + top, Math.min(y, py + ph - bottom - height));
  if (cellTop !== undefined) y = cellTop;
  if (recovery) {
    left = Math.max(rect.left, Math.min(left, rect.right - width));
    y = rect.top + (g.verticalPosition.offset || 0) / 9525;
  }
  // A proven full-width HWP square is aligned to the paper body, not
  // limited to the width of its indented marker paragraph.
  const hwpFullBodyWidth = !!g.hwpBodyFlow && !offPage && g.wrap?.type === 1 && !g.behindDocument && !g.drawingGroup &&
    g.layoutInCell === false && horizontal.relative === 'page' && vertical.relative === 'page' &&
    !Object.values(g.margins || {}).some(n => n) && !cell && !group && !p.closest('.document-textbox,.document-furniture,[data-document-note]') &&
    Math.abs(height - g.hwpBodyFlow.height) <= .16 && Math.abs(left - sr.left - leftMargin) <= .16 &&
    Math.abs(left + width - sr.right + rightMargin) <= .16 && y + height <= py + ph - bottom + .16;
  // A non-wrapping HWP bitmap can be anchored to all four paper edges.
  // The body max-width would shrink its actual drawing without moving it.
  const paper = g.hwpPaperDrawing;
  const paperCover = Boolean(paper && !offPage && g.wrap?.type === 0 && !g.drawingGroup && g.layoutInCell === false &&
    !cell && !group && !p.closest('.document-textbox,.document-furniture,[data-document-note]') &&
    horizontal.relative === 'page' && vertical.relative === 'page' && !horizontal.align && !vertical.align &&
    horizontal.offset === 0 && vertical.offset === 0 && !Object.values(g.margins || {}).some(n => n) &&
    Math.abs(paper.width - sheet.clientWidth) <= 1 && Math.abs(paper.height - ph) <= 1 &&
    Math.abs(paper.width - width) <= 1 && Math.abs(paper.height - height) <= 1);
  const unconstrainedWidth = Boolean(hwpFullBodyWidth || paperEdgeColumnPicture || paperCover);
  if (unconstrainedWidth) {
    img.dataset.hwpFullWidthPicture = 'true';
    img.style.maxWidth = 'none';
  } else if (img.dataset.hwpFullWidthPicture === 'true') {
    delete img.dataset.hwpFullWidthPicture;
    img.style.removeProperty('max-width');
  }
  Object.assign(img.style, {
    display: offPage ? "none" : "",
    position: "absolute",
    float: "none",
    margin: "0",
    left: `${left - rect.left}px`,
    top: `${y - rect.top}px`,
    zIndex: g.behindDocument ? "-1" : "1",
  });
  return {
    index,
    minimumHeight,
    hwpFullBodyWidth,
    unconstrainedWidth,
    hwpPaperCover: paperCover,
    wrapCell: g.layoutInCell !== false ? cell : null,
    cellRelative: !!cellFrame,
    offPage,
    left,
    top: y,
    width,
    height,
    cssLeft: left - rect.left,
    cssTop: y - rect.top,
    geometry: g,
  };
}

function isPaperEdgeColumnPicture(image: HTMLImageElement, floating: ImageFloating, paper: HTMLElement, bodyTop: number) {
  const source = floating.hwpPaperDrawing;
  if (!source) return false;
  const columns = image.closest<HTMLElement>('[data-column-count]');
  const { horizontalPosition: horizontal, verticalPosition: vertical } = floating;
  const x = Number(horizontal.offset) / 9525, y = Number(vertical.offset) / 9525;
  const width = Number(image.getAttribute('width')), height = Number(image.getAttribute('height'));
  // The two DOCX extents can round independently. Retain the source
  // placement only while the drawing is still the same size and on paper.
  const tolerance = .16;
  return !!columns && Number(columns.dataset.columnCount) >= 2 &&
    !image.closest('td,th,.document-textbox') && floating.layoutInCell === false &&
    floating.wrap?.type === 3 && !floating.behindDocument && !floating.drawingGroup &&
    horizontal.relative === 'page' && vertical.relative === 'page' && !horizontal.align && !vertical.align &&
    Math.abs(width - source.width) <= tolerance && Math.abs(height - source.height) <= tolerance &&
    x >= 0 && x + width <= paper.clientWidth + tolerance &&
    y >= 0 && y < bodyTop && y + height <= Number(paper.dataset.pageHeight) + tolerance;
}

function isPaperRelativeColumnImage(image: HTMLImageElement, floating: ImageFloating) {
  const columns = image.closest<HTMLElement>('[data-column-count]');
  const { horizontalPosition: horizontal, verticalPosition: vertical } = floating;
  return !!columns && Number(columns.dataset.columnCount) >= 2 && !image.closest('td,th,.document-textbox') &&
    floating.wrap?.type === 0 && !floating.drawingGroup &&
    horizontal.relative === 'page' && vertical.relative === 'page' && !horizontal.align && !vertical.align &&
    horizontal.offset !== undefined && vertical.offset !== undefined;
}

function canCoalesceColumnImageLayout(root: HTMLElement) {
  let hasPaperColumnImage = false;
  for (const image of floatingImages(root)) {
    const floating = imageFloating(image.dataset.imageFloating);
    if (!floating) continue;
    // Wrapped images depend on settled text geometry. Keep their scheduling
    // even in a document with an unrelated nonwrapping image in columns.
    if (floating.wrap?.type !== 0) return false;
    if (isPaperRelativeColumnImage(image, floating)) hasPaperColumnImage = true;
  }
  return hasPaperColumnImage;
}

function paperColumnTranslation(image: HTMLImageElement, floating: ImageFloating, left: number, top: number, scale: number) {
  const paper = image.closest<HTMLElement>('[data-page-height]');
  const bodyTop = paperInset(paper?.dataset.margin, 96);
  if ((!isPaperRelativeColumnImage(image, floating) && !(paper && isPaperEdgeColumnPicture(image, floating, paper, bodyTop))) ||
      image.style.display === 'none' || image.style.position !== 'absolute') return;
  // An absolute child can paint against a different column fragment from
  // the one used to calculate its CSS offsets. Translate its paint without
  // moving its anchor paragraph or affecting the column flow.
  const painted = image.getBoundingClientRect();
  const dx = (left - painted.left) / scale, dy = (top - painted.top) / scale;
  if (Math.abs(dx) > .5 || Math.abs(dy) > .5) return `${dx}px ${dy}px`;
}

/** Plan in a detached copy of the live geometry. The editable document receives
 * only ProseMirror decorations: no split paragraphs, copied text or DOM moves. */
const floatSpaceRule = 'data-word-floating-spaces';
const floatSpaceProperty = '--word-floating-space-adjustment';
const floatSpaceStyle = (delta: number) => `${floatSpaceProperty}:${delta}px`;
function setFloatSpaceRules(p: HTMLElement, delta: number) {
  p.setAttribute(floatSpaceRule, 'true');
  p.style.setProperty(floatSpaceProperty, `${delta}px`);
}
function clearFloatSpaceRule(p: HTMLElement) {
  p.removeAttribute(floatSpaceRule);
  p.style.removeProperty(floatSpaceProperty);
}
function clearFloatSpaceRules(root: HTMLElement) {
  for (const p of root.querySelectorAll<HTMLElement>(`[${floatSpaceRule}]`)) clearFloatSpaceRule(p);
}
function usesFloatSpaceRules(p: HTMLElement, delta: number) {
  const matches = (element: HTMLElement) => {
    const css = getComputedStyle(element);
    return css.whiteSpace === 'pre-wrap' && Math.abs(parseFloat(css.wordSpacing) - delta) < 0.001;
  };
  if (!matches(p)) return false;
  const runs = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
  while (runs.nextNode()) {
    const run = runs.currentNode;
    if (run.textContent && run.parentElement && !matches(run.parentElement)) return false;
  }
  return true;
}
export function planFloatingLayout(root: HTMLElement, pageFlow?: PageFlow, wordHangingSpaces = false): ParagraphPlan[] {
  const measure = (flow?: PageFlow, spaces?: WordFloatingSpacePass) =>
    planFloatingLayoutPass(root, flow, true, spaces) ?? planFloatingLayoutPass(root, flow, false, spaces)!;
  if (!wordHangingSpaces) return measure(pageFlow);
  const probe = (skip: ReadonlySet<number>) => {
    const rejected = new Set<number>();
    // A rejected view pass must not carry altered paper constraints into
    // the existing float planner. Only publish the accepted evidence.
    const flow = pageFlow && { starts: new Set(pageFlow.starts), backwards: new Set<number>(), candidates: new Set<number>() };
    return { plan: measure(flow, { skip, rejected }), flow, rejected };
  };
  let result = probe(new Set());
  if (result.rejected.size) result = probe(result.rejected);
  if (result.rejected.size) return measure(pageFlow);
  if (pageFlow && result.flow) Object.assign(pageFlow, result.flow);
  return result.plan;
}

function planFloatingLayoutPass(root: HTMLElement, pageFlow: PageFlow | undefined, allowRebase: boolean, wordSpaces?: WordFloatingSpacePass): ParagraphPlan[] | null {
  if (!root.querySelector("img[data-image-floating],.document-textbox")) return [];
  const owner = root.closest<HTMLElement>(".document-stack,.print-page");
  const host = (owner?.cloneNode(false) || document.createElement("div")) as HTMLElement;
  Object.assign(host.style, {
    position: "fixed",
    left: "-100000px",
    top: "0",
    transform: "none",
    zoom: "1",
    visibility: "hidden",
    pointerEvents: "none",
    width: `${parseFloat(getComputedStyle(owner || root).width)}px`,
  });
  const liveImages = pageFlow ? floatingImages(root) : [];
  const ownerStarts = liveImages.map(img => {
    const at = ownerPaper(img, root);
    return at ? at.page + (at.y > at.page * (at.height + 24) + at.top + .5 ? 1 : 0) : 0;
  });
  const copy = root.cloneNode(true) as HTMLElement;
  // Caret separators can leave a phantom line before newly inserted block
  // widgets in a clone. They have no authored width, height or model offset.
  copy.querySelectorAll("[data-flow-widget],.ProseMirror-separator").forEach((el) => el.remove());
  restoreCellPictureReservations(copy);
  clearFloatSpaceRules(copy);
  clearFloatingBottomBorders(copy);
  copy.removeAttribute("contenteditable");
  host.append(copy);
  document.body.append(host);
  const measure = document.createElement("span");
  Object.assign(measure.style, {
    position: "fixed",
    left: "-100000px",
    top: "0",
    visibility: "hidden",
    display: "inline-block",
    whiteSpace: "pre",
    width: "max-content",
    padding: "0",
    margin: "0",
    textIndent: "0",
  });
  measure.className = "flow-text-measure";
  host.append(measure);
  try {
    const ps = paragraphs(copy),
      images = floatingImages(copy);
    if (pageFlow) for (const image of images) image.style.removeProperty('translate');
    // Stripping old wrap widgets changes the clone's physical coordinates.
    // A page-anchored picture must still belong to its source page segment:
    // an explicit body break prevents it from wrapping preceding paper.
    const pageSegments = new Map<Element, number>();
    let segment = 0;
    for (const element of copy.querySelectorAll('p,h1,h2,h3,h4,h5,h6,[data-page-break],table')) {
      if (element.hasAttribute('data-page-break') && !element.closest('td,th,[data-column-count],.document-textbox')) segment++;
      pageSegments.set(element, segment);
    }
    // An empty anchor has no painted line of its own. In a source segment
    // containing a table, cell metrics and row pagination can still change
    // the apparent before/after paper. Require the existing repeated-cycle
    // evidence there instead of committing a transient one-pass constraint.
    const tableSegments = new Set([...copy.querySelectorAll('table')]
      .filter(table => !table.closest('td,th,[data-column-count],.document-textbox'))
      .map(table => pageSegments.get(table)));
    const tableDependentParagraphs = new Set(ps.filter(p => tableSegments.has(pageSegments.get(p))));
    const samePageSegment = (picture: Picture, paragraph: HTMLElement) => {
      if (picture.index < 0) return true;
      const image = images[picture.index];
      if (image.closest('td,th,[data-column-count],.document-textbox')) return true;
      return pageSegments.get(image.closest('p,h1,h2,h3,h4,h5,h6')!) === pageSegments.get(paragraph);
    };
    const imageOwners = images.map(image => image.closest<HTMLElement>('p,h1,h2,h3,h4,h5,h6')!);
    const forcedOwners = new Map<HTMLElement, number>();
    for (const index of pageFlow?.starts || []) if (imageOwners[index]) forcedOwners.set(imageOwners[index], index);
    if (pageFlow) {
      pageFlow.candidates.clear(); pageFlow.backwards.clear();
      for (const [index, image] of images.entries()) {
        const geometry = imageFloating(image.dataset.imageFloating), paragraph = imageOwners[index];
        if (!geometry || !paragraph || geometry.wrap?.type === 0 || geometry.behindDocument || geometry.drawingGroup || imageOutsidePaper(image) ||
          !['page', 'margin'].includes(geometry.verticalPosition.relative || '') || paragraph.closest('td,th,[data-column-count],.document-textbox') ||
          !(paragraph.parentElement === copy || paragraph.parentElement?.classList.contains('section-editable')) ||
          paragraph.querySelectorAll('img[data-image-floating]').length !== 1) continue;
        const paper = ownerPaper(image, copy)!, box = paragraph.getBoundingClientRect();
        const line = parseFloat(getComputedStyle(paragraph).lineHeight) || 24;
        if (box.height <= line * 1.5 && (Number(image.getAttribute('height')) || 0) + box.height <= paper.height - paper.top - paper.bottom + .5)
          pageFlow.candidates.add(index);
      }
      // An automatic anchor break has the same wrap boundary as a source break.
      let segment = 0;
      for (const element of copy.querySelectorAll('p,h1,h2,h3,h4,h5,h6,[data-page-break]')) {
        if ((element.hasAttribute('data-page-break') && !element.closest('td,th,[data-column-count],.document-textbox')) || forcedOwners.has(element as HTMLElement)) segment++;
        pageSegments.set(element, segment);
      }
    }
    prepareDrawingGroupAnchors(copy);
    layoutTextBoxes(copy);
    const pictures = images
      .map((img, index) => imagePosition(img, index, copy))
      .filter((p): p is Picture => !!p);
    for (const [index, frame] of [...copy.querySelectorAll<HTMLElement>('.document-textbox[data-document-textbox]')].entries()) {
      const data=textBoxFrame(frame.dataset.documentTextbox),box=frame.firstElementChild?.getBoundingClientRect();
      if(!data || data.position!=='floating' || data.wrap!=='square' || !box)continue;
      pictures.push({ index:-1-index,left:box.left,top:box.top,width:box.width,height:box.height,cssLeft:0,cssTop:0,
        geometry:{ horizontalPosition:{relative:'page',offset:0},verticalPosition:{relative:'page',offset:0},wrap:{type:1},margins:{top:(data.wrapMargins?.[0]||0)*9525,right:(data.wrapMargins?.[1]||0)*9525,bottom:(data.wrapMargins?.[2]||0)*9525,left:(data.wrapMargins?.[3]||0)*9525} } });
    }
    const initialPictures = pageFlow ? new Map(pictures.filter(p=>p.index>=0).map(p=>[p.index,p.top])) : undefined;

    // Removing the old floating widgets can move a blank anchor before a
    // table gap in this copy. When its live paper is later, its temporary
    // position must not wrap content already wholly on an earlier paper.
    const laterLivePapers = new Map<number, {
      sheet: HTMLElement; previousPaperBottom: number; anchor: NonNullable<ReturnType<typeof ownerPaper>>;
    }>();
    for (const picture of pictures) {
      if (!pageFlow?.candidates.has(picture.index) || pageFlow.starts.has(picture.index) ||
          picture.geometry.wrap?.type !== 1 || picture.geometry.verticalPosition.relative !== 'page') continue;
      const owner = imageOwners[picture.index];
      if (!tableDependentParagraphs.has(owner) || owner.textContent?.replace(/[\s\u200b\u2060\ufeff]/gu, '')) continue;
      const actual = ownerPaper(liveImages[picture.index], root);
      const planned = ownerPaper(images[picture.index], copy);
      if (actual && planned && actual.page > planned.page) {
        const pageEnd = actual.page * (actual.height + 24) - 24;
        const previousPaperBottom = actual.sheet.getBoundingClientRect().top + (pageEnd + .5) * documentScale(actual.sheet);
        laterLivePapers.set(picture.index, { sheet: actual.sheet, previousPaperBottom, anchor: actual });
      }
    }
    const liveParagraphs = laterLivePapers.size ? paragraphs(root) : [];

    type Projection = {
      picture: Picture; anchorY: number; recipientY: number; page: number; shift: number; bottom: number;
    };
    const proposals = new Map<HTMLElement, Projection>();
    const applied = new Map<number, Projection>();
    const overlapsOrdinaryFlow = () => {
      if (copy.querySelector('.table-overlay-active,.table-wrap-frame,[data-textbox-position="floating"]')) return true;
      if ([...copy.querySelectorAll<HTMLElement>('table[data-table-anchor]')].some(table => {
        const anchor = tableAnchor(table.dataset.tableAnchor);
        return !!(anchor?.wrap || anchor?.overlay);
      })) return true;
      return [...copy.querySelectorAll<HTMLElement>('[style*="margin"]')].some(block =>
        parseFloat(block.style.marginTop) < 0 || parseFloat(block.style.marginBottom) < 0);
    };
    const projectOrdinaryBody = allowRebase && laterLivePapers.size > 0 && !overlapsOrdinaryFlow();
    function firstTextAfterAnchor(anchor: HTMLElement | null | undefined) {
      let p = anchor?.nextElementSibling;
      while (p instanceof HTMLElement && p.tagName === 'P' &&
             !p.querySelector('img,svg,table,video,canvas,[data-note-reference],[data-page-number]')) {
        if (p.textContent?.trim()) return p;
        p = p.nextElementSibling;
      }
    }
    function proposeProjection(picture: Picture) {
      if (!projectOrdinaryBody) return;
      const liveAnchor = laterLivePapers.get(picture.index)?.anchor;
      const img = liveImages[picture.index];
      if (!liveAnchor || !img || getComputedStyle(img).position !== 'absolute') return;
      const owner = imageOwners[picture.index];
      const cloneAnchor = ownerPaper(owner, copy);
      if (!cloneAnchor || cloneAnchor.page !== liveAnchor.page) return;
      const receiver = firstTextAfterAnchor(owner), liveReceiver = firstTextAfterAnchor(img.closest('p'));
      if (!receiver || !liveReceiver || proposals.has(receiver)) return;
      const inCopy = ownerPaper(receiver, copy), inLive = ownerPaper(liveReceiver, root);
      if (!inCopy || !inLive || inCopy.sheet !== cloneAnchor.sheet || inLive.sheet !== liveAnchor.sheet ||
          inCopy.page !== cloneAnchor.page || inLive.page !== liveAnchor.page ||
          Math.abs((inLive.y - liveAnchor.y) - (inCopy.y - cloneAnchor.y)) > .5) return;
      const shift = liveAnchor.y - cloneAnchor.y;
      if (shift <= .5) return;
      const translated = { ...picture, top: picture.top - shift, cssTop: picture.cssTop - shift };
      const margin = picture.geometry.margins;
      const top = Math.min(translated.top, picture.top) - (margin?.top || 0) / 9525;
      const bottom = Math.max(translated.top, picture.top) + picture.height + (margin?.bottom || 0) / 9525;
      const rect = receiver.getBoundingClientRect(), height = liveReceiver.getBoundingClientRect().height / inLive.scale;
      if (top < rect.top - .5 || bottom > rect.top + height + .5 ||
          inLive.y + height > inLive.page * (inLive.height + 24) + inLive.height + .5) return;
      proposals.set(receiver, { picture: translated, anchorY: cloneAnchor.y, recipientY: inCopy.y,
        page: cloneAnchor.page, shift, bottom });
    }

    const groupIds=new Set<string>();
    for(const picture of [...pictures])if(picture.index>=0){const group=groupExclusion(picture);if(group && !groupIds.has(group.geometry.drawingGroup!.id)){pictures.push(group);groupIds.add(group.geometry.drawingGroup!.id);}}
    const result: ParagraphPlan[] = [];
    for (let pi = 0; pi < ps.length; pi++) {
      const p = ps[pi];
      const forced = forcedOwners.get(p);
      if (forced !== undefined) {
        const sheet = p.closest<HTMLElement>('[data-page-height],.print-page') || copy, paper = ownerPaper(images[forced], copy)!;
        const current = paper.page + (paper.y > paper.page * (paper.height + 24) + paper.top + .5 ? 1 : 0);
        const target = Math.max(current, ownerStarts[forced] || 0);
        const gap = Math.max(0, Math.ceil((sheet.getBoundingClientRect().top + target * (paper.height + 24) + paper.top - p.getBoundingClientRect().top) * 64) / 64);
        if (gap > .1) {
          const spacer = document.createElement('div'); spacer.dataset.flowWidget = 'true'; spacer.style.height = `${gap}px`; p.before(spacer);
        }
      }
      const own = images
          .flatMap((img, index) =>
            img.closest("p,h1,h2,h3,h4,h5,h6") === p ? [imagePosition(img, index, copy)!] : [],
          )
          .filter(Boolean);
      for (const picture of own) {
        const i = pictures.findIndex((v) => v.index === picture.index);
        if (i >= 0) pictures[i] = picture;
        proposeProjection(picture);
        const group=groupExclusion(picture);if(group){const gi=pictures.findIndex(v=>v.index<=-100000 && v.geometry.drawingGroup?.id===group.geometry.drawingGroup!.id);if(gi>=0)pictures[gi]=group;}
      }
      // A body paragraph containing only floating anchors has no painted
      // inline content for its own artwork to wrap. Preserve its natural
      // source line; paper fitting still includes the drawing's real extent.
      if (own.length && (!p.closest('td,th,.document-textbox,[data-column-count]') || own.every(picture => picture.cellRelative)) &&
          !p.textContent?.replace(/[\s\u200b\u2060\ufeff]/gu, '') &&
          !p.querySelector('img:not([data-image-floating]):not(.ProseMirror-separator),svg,video,canvas,[data-note-reference],[data-page-number],[data-symbol-font],[data-word-overlap],[data-docx-tab]') &&
          own.every(picture => !picture.geometry.drawingGroup)) {
        if (own.length === 1) reserveHwpBodyPicture(p, own[0], copy);
        result.push({ index: pi, widgets: [], pictures: own });
        continue;
      }
      // Behind/in-front artwork needs positioning but does not exclude text.
      // Its often page-sized bounds must not trigger a manual re-layout of
      // unrelated paragraphs, which changes justification and mixed-font
      // line endings even though every line still has its full width.
      // A cell-bound drawing must not reflow paragraphs in another cell.
      // Otherwise page-anchored artwork can chase an earlier row as its page moves.
      const liveParagraph = liveParagraphs[pi];
      const liveSheet = liveParagraph?.closest('[data-page-height],.print-page');
      let liveBottom: number | undefined;
      const wrapsOnlyLaterPaper = (picture: Picture) => {
        const later = laterLivePapers.get(picture.index);
        if (!liveParagraph || !later || liveSheet !== later.sheet || p === imageOwners[picture.index] ||
            !(p.compareDocumentPosition(images[picture.index]) & Node.DOCUMENT_POSITION_FOLLOWING)) return false;
        liveBottom ??= liveParagraph.getBoundingClientRect().bottom;
        return liveBottom <= later.previousPaperBottom;
      };
      const projection = proposals.get(p);
      if (projection) {
        const receiver = ownerPaper(p, copy);
        if (!receiver || receiver.page !== projection.page || Math.abs(receiver.y - projection.recipientY) > .5) return null;
      }
      const flowing = pictures.filter(pic => samePageSegment(pic, p) && !pic.offPage && pic.geometry.wrap?.type !== 0 &&
        !wrapsOnlyLaterPaper(pic) &&
        (!pic.wrapCell || pic.wrapCell === p.closest('td,th')) && !(pic.index<0 && p.closest('.document-textbox')))
        .map(pic => projection && projection.picture.index === pic.index ? projection.picture : pic);
      const style = getComputedStyle(p),
        box = p.getClientRects()[0] || p.getBoundingClientRect();
      const lh = parseFloat(style.lineHeight) || parseFloat(style.fontSize) * 1.3,
        width = parseFloat(style.width) || box.width;
      const group = p.closest<HTMLElement>("[data-column-count]");
      const groupBox = group?.getBoundingClientRect(),
        colHeight = group ? parseFloat(group.style.getPropertyValue("column-height")) : 0;
      const colGap = Number(group?.dataset.columnSpace) || 0,
        colWidth = group
          ? (group.clientWidth - (Number(group.dataset.columnCount) - 1) * colGap) /
            Number(group.dataset.columnCount)
          : width;
      const rowGap = group ? parseFloat(group.style.rowGap) || 0 : 0;
      let x = box.left,
        y = box.top;
      const relevant = () =>
        flowing.some((pic) => exclusion(pic, y, lh).some(([a, b]) => b > x && a < x + width));
      if (
        !flowing.some((pic) => pic.top + pic.height >= y - 0.5 && pic.top < box.bottom + 0.5)
      ) {
        if (projection && flowing.includes(projection.picture)) return null;
        if (own.length) result.push({ index: pi, widgets: [], pictures: own });
        continue;
      }
      const firstIndent = parseFloat(style.textIndent) || 0;
      const firstSlots = wordSpaces && !wordSpaces.skip.has(pi) && p.parentElement === copy && !own.length &&
        (!copy.classList.contains('print-content') || p.dataset.wordFloatingSpaceSource === 'true') &&
        !p.hasAttribute('data-print-paragraph-fragment') && style.whiteSpace === 'break-spaces' &&
        style.textAlignLast === 'auto' && ['left', 'start', 'justify'].includes(style.textAlign) &&
        isOrdinaryWordBodyParagraph(p, copy) && relevant()
          ? slots(flowing, x + firstIndent, y, width - firstIndent, lh) : [];
      const isRightSlot = (spans: [number, number][], indent: number) =>
        spans.length === 1 && spans[0][0] > x + indent + .2 && Math.abs(spans[0][1] - (x + width)) < .1;
      let wordSpaceAdjustment = isRightSlot(firstSlots, firstIndent) ? measuredBodySpaceAdjustment(p) : null;
      if (wordSpaceAdjustment !== null) {
        setFloatSpaceRules(p, wordSpaceAdjustment);
        if (!usesFloatSpaceRules(p, wordSpaceAdjustment)) {
          clearFloatSpaceRule(p);
          wordSpaceAdjustment = null;
        }
      }
      const tokens = units(p),
        widgets: Widget[] = [],
        cache = new Map<string, { width: number; height: number }>();
      for (const property of [
        "fontFamily",
        "fontSize",
        "fontWeight",
        "fontStyle",
        "fontStretch",
        "fontVariant",
        "letterSpacing",
        "wordSpacing",
        "textTransform",
        "lineHeight",
      ] as const)
        measure.style[property] = style[property];
      // Keep unitless leading unitless: a larger inline font must enlarge
      // its actual line box just as it does in the editable paragraph.
      measure.style.lineHeight = p.style.lineHeight || String(lh / parseFloat(style.fontSize));
      const metrics = (start: number, end: number) => {
        if (start === end) return { width: 0, height: lh };
        const key = `${start}:${end}`;
        if (cache.has(key)) return cache.get(key)!;
        const range = document.createRange();
        range.setStart(...tokens[start].start);
        range.setEnd(...tokens[end - 1].end);
        let fragment: globalThis.Node = range.cloneContents();
        let ancestor =
          range.commonAncestorContainer.nodeType === Node.TEXT_NODE
            ? range.commonAncestorContainer.parentElement
            : (range.commonAncestorContainer as HTMLElement);
        // A range entirely inside a bold/link/style run does not clone that
        // common ancestor. Restore its inherited styles for faithful widths.
        while (ancestor && ancestor !== p) {
          const wrapper = ancestor.cloneNode(false);
          wrapper.appendChild(fragment);
          fragment = wrapper;
          ancestor = ancestor.parentElement;
        }
        measure.replaceChildren(fragment);
        measure
          .querySelectorAll("img[data-image-floating],[data-flow-widget]")
          .forEach((el) => el.remove());
        const rect = measure.getBoundingClientRect();
        const value = { width: rect.width, height: Math.max(lh, rect.height) };
        cache.set(key, value);
        return value;
      };
      const size = (start: number, end: number, hangSpace = false) => {
        if (hangSpace && end > start) {
          const originalEnd = end;
          while (end > start && tokens[end - 1].text === ' ') end--;
          // Retain actual fit requirements for an all-space range.
          if (end === start) return metrics(start, originalEnd).width;
        }
        return metrics(start, end).width;
      };
      // A measured range must really fit. Allowing a positive tolerance can
      // wrap its final whole word onto an extra native line without that
      // line's picture-exclusion widget. Reserve one Chromium layout unit;
      // text size and the document's wrap geometry remain unchanged.
      const fitReserve = 1 / 64;
      const fit = (start: number, end: number, available: number, hangSpace = false) => {
        let low = start,
          high = end,
          best = start;
        while (low <= high) {
          const mid = (low + high) >> 1;
          if (size(start, mid, hangSpace) <= available - fitReserve) {
            best = mid;
            low = mid + 1;
          } else high = mid - 1;
        }
        if (
          best < end &&
          best > start &&
          /[\p{Script=Latin}\p{N}]/u.test(tokens[best - 1].text) &&
          /[\p{Script=Latin}\p{N}]/u.test(tokens[best].text)
        ) {
          let boundary = best;
          while (boundary > start && /[\p{Script=Latin}\p{N}]/u.test(tokens[boundary - 1].text))
            boundary--;
          if (boundary > start) best = boundary;
        }
        return best;
      };
      let bottomBorderInset: number | undefined;
      let bottomBorderLineHeight = 0;
      const bottomBorder = !own.length && !p.hasAttribute('data-print-paragraph-fragment') && !p.hasAttribute('data-print-soft-end') &&
        p.hasAttribute(wordBottomBorderEnd) && firstIndent === 0 && ['left', 'start'].includes(style.textAlign) &&
        style.direction === 'ltr' && style.writingMode === 'horizontal-tb' && !projection &&
        !root.querySelector('[data-note-reference],table[data-table-anchor],[data-column-count]:not([data-column-count="1"])')
        ? floatingBottomBorderDimensions(p) : null;
      let start = 0,
        line = 0;
      while (start < tokens.length && line++ < 10000) {
        if (
          !relevant() &&
          y >
            Math.max(
              ...flowing.map(
                (pic) => pic.top + pic.height + (pic.geometry.margins?.bottom || 0) / 9525,
              ),
            )
        )
          break;
        const indent = line === 1 ? firstIndent : 0;
        const hard = tokens.findIndex((token, i) => i >= start && token.hard);
        const end = hard < 0 ? tokens.length : hard,
          lineStart = start;
        let lineHeight = metrics(start, fit(start, end, width - indent)).height;
        const arrange = (height: number) => {
          // A previous float can move a line into a later float. Observe the
          // actual attempted line, not only the paragraph's original box.
          if (pageFlow) for (const picture of flowing) {
            const image = picture.index >= 0 ? images[picture.index] : undefined;
            if (image && !p.contains(image) && (p.compareDocumentPosition(image) & Node.DOCUMENT_POSITION_FOLLOWING) &&
              exclusion(picture, y, height).some(([left, right]) => left < x + width && right > x)) pageFlow.backwards.add(picture.index);
          }
          const lineWidgets: Widget[] = [];
          const spans = slots(flowing, x + indent, y, width - indent, height);
          let nextStart = lineStart,
            cursor = x + indent, hanging = false;
          const supported = isRightSlot(spans, indent);
          for (const [left, right] of spans) {
            const hangSpace = wordSpaceAdjustment !== null && supported;
            const next = fit(nextStart, end, right - left, hangSpace);
            if (
              next === nextStart &&
              right - left < Math.min(width, size(nextStart, Math.min(nextStart + 1, end)))
            )
              continue;
            // A tiny strip beside a float must not split an ordinary Latin
            // word into one-letter columns when that word fits the body width.
            if(next<end && next>nextStart && /[\p{Script=Latin}\p{N}]/u.test(tokens[next-1].text) && /[\p{Script=Latin}\p{N}]/u.test(tokens[next].text)) {
              let wordEnd=next;while(wordEnd<end && /[\p{Script=Latin}\p{N}]/u.test(tokens[wordEnd].text))wordEnd++;
              if(size(nextStart,wordEnd)<=width-indent && size(nextStart,wordEnd)>right-left-fitReserve)continue;
            }
            const last = next === nextStart && nextStart < end ? nextStart + 1 : next;
            hanging = hangSpace && last > nextStart && tokens[last - 1].text === ' ';
            const textWidth = size(nextStart, last, hanging);
            const alignment =
              style.textAlign === "right" || style.textAlign === "end"
                ? 1
                : style.textAlign === "center"
                  ? 0.5
                  : 0;
            const placedLeft = left + Math.max(0, right - left - textWidth) * alignment;
            if (placedLeft > cursor + 0.2)
              lineWidgets.push({
                offset: tokens[nextStart]?.from ?? tokens.at(-1)!.to,
                width: placedLeft - cursor,
              });
            cursor = placedLeft + textWidth;
            nextStart = last;
            if (nextStart >= end) break;
          }
          return { lineWidgets, nextStart, cursor, hanging, blocked: !spans.length, supported, spans };
        };
        let arranged = arrange(lineHeight);
        // Exclusions depend on line height, and the selected words determine
        // that height. Re-evaluate when a font-size boundary changes the fit.
        for (let attempt = 0; attempt < 3 && !arranged.blocked; attempt++) {
          const actual = metrics(lineStart, arranged.nextStart).height;
          if (Math.abs(actual - lineHeight) < 0.1) break;
          lineHeight = actual;
          arranged = arrange(lineHeight);
        }
        if (wordSpaceAdjustment !== null && !arranged.supported) wordSpaces?.rejected.add(pi);
        if (bottomBorder && line === 1 && arranged.supported && arranged.nextStart === tokens.length) {
          bottomBorderInset = arranged.spans[0][0] - x;
          bottomBorderLineHeight = lineHeight;
        }
        if (arranged.blocked) {
          const jump = Math.max(
            lineHeight,
            Math.min(
              ...flowing
                .filter((pic) => pic.top + pic.height > y)
                .map(
                  (pic) => pic.top + pic.height - y + (pic.geometry.margins?.bottom || 0) / 9525,
                ),
            ),
          );
          widgets.push({ offset: tokens[start].from, height: jump });
          y += jump;
          continue;
        }
        widgets.push(...arranged.lineWidgets);
        start = arranged.nextStart;
        // Reserve unused width so native center/right alignment cannot shift
        // a measured gap over a picture. The unplanned tail keeps its alignment.
        const remaining = x + width - arranged.cursor - 0.4;
        if (remaining > 0.2 && !arranged.hanging)
          widgets.push({
            offset: tokens[start]?.from ?? tokens.at(-1)!.to,
            width: remaining,
          });
        if (start === end && hard >= 0) {
          start++;
          y += lineHeight;
        } else if (start < tokens.length) {
          widgets.push({ offset: tokens[start].from, break: true });
          y += lineHeight;
        } else break;
        if (group && groupBox && colHeight) {
          const row = Math.floor((y - groupBox.top) / (colHeight + rowGap)),
            rowTop = groupBox.top + row * (colHeight + rowGap);
          if (y + lh > rowTop + colHeight + 0.5) {
            const col = Math.round((x - groupBox.left) / (colWidth + colGap));
            if (col + 1 < Number(group.dataset.columnCount)) {
              x += colWidth + colGap;
              y = rowTop;
            } else {
              x = groupBox.left;
              y = rowTop + colHeight + rowGap;
            }
          }
        }
      }
      insertWidgets(p, widgets);
      if (bottomBorderInset !== undefined && bottomBorder) {
        const paper = ownerPaper(p, copy), rect = p.getBoundingClientRect();
        const strokeTop = rect.bottom + bottomBorder.gap;
        const strokeLeft = x + bottomBorderInset;
        const band = slots(flowing, x, strokeTop, width, bottomBorder.stroke);
        // Pictures with no text wrap do not contribute to slots, but their
        // painted rectangles can still cross the proposed line.
        const crossesPicture = pictures.some(picture => !picture.offPage &&
          picture.left < x + width && picture.left + picture.width > strokeLeft &&
          picture.top < strokeTop + bottomBorder.stroke && picture.top + picture.height > strokeTop);
        const reserve = bottomBorder.gap + bottomBorder.stroke + Math.max(0, parseFloat(style.marginBottom) || 0);
        // Use the verified interval at the stroke, not a block-wide border
        // through a nearby picture or a line whose clear interval changes.
        if (paper && !crossesPicture && rect.height / paper.scale <= bottomBorderLineHeight + .5 &&
            isRightSlot(band, 0) && Math.abs(band[0][0] - x - bottomBorderInset) < .1 &&
            paper.y + rect.height / paper.scale + reserve <=
              paper.page * (paper.height + 24) + paper.height - paper.bottom - .5)
          applyFloatingBottomBorder(p, bottomBorderInset);
        else bottomBorderInset = undefined;
      }
      if (wordSpaceAdjustment !== null) {
        const paper = ownerPaper(p, copy);
        const after = Math.max(0, parseFloat(style.marginBottom) || 0);
        // A split would clone a plan into a different slot on the next page.
        // Only retain a paragraph whose complete adjusted flow needs no split.
        if (start < tokens.length || !paper ||
            paper.y + p.getBoundingClientRect().height / paper.scale + after >
              paper.page * (paper.height + 24) + paper.height - paper.bottom - 0.5)
          wordSpaces?.rejected.add(pi);
      }
      if (projection && flowing.includes(projection.picture)) {
        // If either band can reach the next block, keep placement and wrap
        // together by discarding this pass, before applying any decorations.
        const receiver = ownerPaper(p, copy), rect = p.getBoundingClientRect();
        if (!receiver || receiver.page !== projection.page || rect.bottom < projection.bottom - .5 ||
            receiver.y + rect.height / receiver.scale > receiver.page * (receiver.height + 24) + receiver.height + .5) return null;
        applied.set(projection.picture.index, projection);
      }
      result.push({ index: pi, widgets, pictures: own, ...(wordSpaceAdjustment !== null ? { wordSpaceAdjustment } : {}),
        ...(bottomBorderInset !== undefined ? { bottomBorderInset } : {}) });
    }
    // Flow widgets can change a vertically aligned cell's content height and
    // move its anchor paragraph. Resolve absolute picture offsets against
    // that final geometry, not the paragraph's pre-wrap position.
    for (const paragraph of result) {
      const resolved: Picture[] = [];
      for (const picture of paragraph.pictures) {
        const image = images[picture.index];
        const placed = imagePosition(image, picture.index, copy) || picture;
        if (pageFlow) placed.cssTranslation = paperColumnTranslation(image, placed.geometry, placed.left, placed.top, 1);
        const projection = applied.get(picture.index);
        if (projection) {
          const paper = ownerPaper(images[picture.index], copy);
          if (!paper || paper.page !== projection.page || Math.abs(paper.y - projection.anchorY) > 1 / 64) return null;
        }
        resolved.push(projection ? { ...placed, cssTop: placed.cssTop - projection.shift } : placed);
      }
      paragraph.pictures = resolved;
    }
    for (const paragraph of result) {
      paragraph.minimumHeight = Math.max(0, ...paragraph.pictures.map(p => p.minimumHeight || 0)) || undefined;
    }
    for (const [index, block] of ps.entries()) {
      const reservation = cellPictureReservation(block);
      if (!reservation) continue;
      let paragraph = result.find(value => value.index === index);
      if (!paragraph) { paragraph = { index, widgets: [], pictures: [] }; result.push(paragraph); }
      paragraph.minimumHeight = Math.max(paragraph.minimumHeight || 0, reservation.minimumHeight || 0) || undefined;
      paragraph.originalMinimum = reservation.originalMinimum;
    }
    // Short anchors must not move to another paper after their own
    // earlier-text exclusion has already been calculated. Start that anchor
    // on fresh paper and plan again with the new page boundary in place.
    if (pageFlow) for (const paragraph of result) {
      for (const picture of paragraph.pictures) {
        if (pageFlow.starts.has(picture.index)) paragraph.startPage = true;
        else if ((ps[paragraph.index].textContent?.replace(/[\s\u200b\u2060\ufeff]/gu, "") || !tableDependentParagraphs.has(ps[paragraph.index])) && pageFlow.candidates.has(picture.index) && pageFlow.backwards.has(picture.index)) {
          const paper = ownerPaper(images[picture.index], copy)!, initial = initialPictures!.get(picture.index)!;
          if (picture.top - initial >= paper.height + 24 - .5) pageFlow.starts.add(picture.index);
        }
      }
    }
    return result;
  } finally {
    host.remove();
  }
}

export function applyFloatingLayout(root: HTMLElement, wordHangingSpaces = false) {
  root.querySelectorAll("[data-flow-widget]").forEach((el) => el.remove());
  clearFloatSpaceRules(root);
  clearFloatingBottomBorders(root);
  restoreCellPictureReservations(root);
  const plan = planFloatingLayout(root, undefined, wordHangingSpaces),
    ps = paragraphs(root),
    images = floatingImages(root);
  for (const paragraph of plan) {
    const p = ps[paragraph.index];
    if (paragraph.wordSpaceAdjustment !== undefined) setFloatSpaceRules(p, paragraph.wordSpaceAdjustment);
    if (paragraph.bottomBorderInset !== undefined) applyFloatingBottomBorder(p, paragraph.bottomBorderInset);
    if (paragraph.pictures.length) p.style.position = "relative";
    if (paragraph.originalMinimum !== undefined) p.setAttribute('data-floating-cell-original-min-height', paragraph.originalMinimum);
    if (paragraph.minimumHeight) {
      p.style.minHeight = `${Math.max(parseFloat(p.style.minHeight) || 0, paragraph.minimumHeight)}px`;
    }
    for (const picture of paragraph.pictures) {
      const img = images[picture.index];
      const own = imagePosition(img, picture.index, root);
      if (own) img.style.zIndex = own.geometry.behindDocument ? "-1" : "1";
    }
    insertWidgets(p, paragraph.widgets);
  }
  images.forEach((img, index) => imagePosition(img, index, root));
}

export function positionPaperImagesInPrintColumns(root: HTMLElement) {
  for (const image of root.querySelectorAll<HTMLImageElement>('.print-content [data-column-count] img[data-image-floating]')) {
    const paper = image.closest<HTMLElement>('.print-page');
    const floating = imageFloating(image.dataset.imageFloating);
    if (!paper || !floating) continue;
    const {horizontalPosition: horizontal, verticalPosition: vertical} = floating;
    if (horizontal.offset === undefined || vertical.offset === undefined) continue;
    const paperRect = paper.getBoundingClientRect();
    const scale = documentScale(paper);
    const translation = paperColumnTranslation(image, floating,
      paperRect.left + horizontal.offset / 9525 * scale, paperRect.top + vertical.offset / 9525 * scale, scale);
    if (translation) image.style.translate = translation;
  }
}

const key = new PluginKey("floatingLayout");
export const FloatingLayout = Extension.create({
  name: "floatingLayout",
  addProseMirrorPlugins() {
    return [
      new Plugin({
        key,
        state: {
          init: () => DecorationSet.empty,
          apply: (tr, old) => tr.getMeta(key) ?? old.map(tr.mapping, tr.doc),
        },
        props: { decorations: (state) => key.getState(state) },
        view(view) {
          let frame = 0,
            signature = "",
            lastDocument = view.state.doc;
          const pageFlow: PageFlow = { starts: new Set(), backwards: new Set(), candidates: new Set() };
          const backwardsSeen = new Set<number>(), paperHistory = new Map<number, OwnerPageStep[]>();
          let layoutPass = 0;
          let flowDocument: typeof view.state.doc | undefined = view.state.doc, flowShape = '';
          let scheduledDocument: typeof view.state.doc | undefined;
          let coalesceColumnLayout = false;
          const calculate = () => {
            if (frame) {
              if (scheduledDocument !== view.state.doc) {
                scheduledDocument = view.state.doc;
                coalesceColumnLayout = canCoalesceColumnImageLayout(view.dom);
              }
              // Column image paint can starve while the editor keeps updating.
              // Measure its latest anchor in the pass already queued.
              if (coalesceColumnLayout) return;
              cancelAnimationFrame(frame);
            }
            frame = requestAnimationFrame(() => {
              frame = 0;
              if (
                editorLayoutSuspended() || !view.dom.isConnected ||
                view.composing ||
                view.dom.classList.contains("table-row-resizing")
              )
                return;
              const ps = paragraphs(view.dom),
                pos = new Map<HTMLElement, number>();
              view.state.doc.descendants((node, position) => {
                if (node.isTextblock) {
                  const dom = view.nodeDOM(position);
                  if (dom instanceof HTMLElement) pos.set(dom, position);
                }
              });
              const sheet=view.dom.closest<HTMLElement>('.document-stack'),shape=sheet?[sheet.dataset.pageHeight,sheet.dataset.margin,sheet.dataset.bottomMargin,sheet.dataset.leftMargin,sheet.dataset.rightMargin,getComputedStyle(sheet).width].join('|'):'';
              if(flowDocument!==view.state.doc || flowShape!==shape) {
                pageFlow.starts.clear();pageFlow.backwards.clear();pageFlow.candidates.clear();backwardsSeen.clear();paperHistory.clear();flowDocument=view.state.doc;flowShape=shape;
              }
              const layout = layoutHash(`${sheet?.dataset.pages}:${view.dom.getBoundingClientRect().height}:${signature}`);
              const pass = ++layoutPass, repeated: RepeatedOwner[] = [];
              floatingImages(view.dom).forEach((image, index) => {
                if (!pageFlow.candidates.has(index) || pageFlow.starts.has(index)) return;
                const at = ownerPaper(image, view.dom); if (!at) return;
                const history = paperHistory.get(index) || [], current = { page: at.page, layout, pass };
                if (history.at(-1)?.page === at.page) history[history.length - 1] = current;
                else { history.push(current); if (history.length > 5) history.shift(); }
                paperHistory.set(index, history);
                // Require two complete repetitions of the same page/layout
                // cycle. A transient return during initial pagination is not
                // evidence that a backward wrap needs an anchor break.
                const equal = (a: number, b: number) => history[a].page === history[b].page && history[a].layout === history[b].layout;
                if (backwardsSeen.has(index) && history.length === 5 && equal(0, 2) && equal(2, 4) && equal(1, 3) && history[0].page !== history[1].page)
                  repeated.push({ index, image, paper: at, history });
              });
              if (repeated.length === 2 && coupledOwnerPair(repeated[0], repeated[1], pos, view.state.doc)) {
                const [earlier, later] = repeated;
                pageFlow.starts.add(later.index);
                paperHistory.delete(earlier.index);
                backwardsSeen.delete(earlier.index);
              } else {
                for (const owner of repeated) pageFlow.starts.add(owner.index);
              }
              let plan: ParagraphPlan[];
              const wordHangingSpaces = documentMetadata.getState(view.state)?.wordHangingSpaces === true &&
                !view.dom.querySelector('[data-note-reference]');
              // Only add proven source-anchor constraints. The set grows
              // monotonically and is bounded by the document's pictures;
              // no iteration cap freezes an unresolved layout.
              for (;;) {
                const count = pageFlow.starts.size; plan = planFloatingLayout(view.dom, pageFlow, wordHangingSpaces);
                for (const index of pageFlow.backwards) backwardsSeen.add(index);
                if (pageFlow.starts.size === count) break;
              }
              const decorations: Decoration[] = [],
                values: string[] = [];
              for (const paragraph of plan) {
                const p = ps[paragraph.index],
                  position = pos.get(p);
                if (position === undefined) continue;
                if (paragraph.wordSpaceAdjustment !== undefined) {
                  decorations.push(Decoration.node(position, position + view.state.doc.nodeAt(position)!.nodeSize,
                    { [floatSpaceRule]: 'true', style: floatSpaceStyle(paragraph.wordSpaceAdjustment) }));
                  values.push(`float-word-space:${position}:${paragraph.wordSpaceAdjustment}`);
                }
                if (paragraph.bottomBorderInset !== undefined) {
                  const paint = floatingBottomBorderPaint(p, paragraph.bottomBorderInset);
                  if (paint) {
                    decorations.push(Decoration.node(position, position + view.state.doc.nodeAt(position)!.nodeSize, paint));
                    values.push(`bottom-border:${position}:${paint.style}`);
                  }
                }
                if (paragraph.pictures.length || paragraph.originalMinimum !== undefined)
                  decorations.push(
                    Decoration.node(
                      position,
                      position + view.state.doc.nodeAt(position)!.nodeSize,
                      { style: `position:relative${paragraph.minimumHeight ? `;min-height:${paragraph.minimumHeight}px` : ""}`, ...(paragraph.startPage?{'data-floating-start-page':'true'}:{}), ...(paragraph.originalMinimum !== undefined ? {'data-floating-cell-original-min-height': paragraph.originalMinimum} : {}) },
                    ),
                  );
                if(paragraph.originalMinimum !== undefined)values.push(`cell-picture-min:${position}:${paragraph.minimumHeight}:${paragraph.originalMinimum}`);
                if(paragraph.startPage)values.push(`float-start-page:${position}`);
                const index = units(p);
                for (const picture of paragraph.pictures) {
                  const token = index.find(
                    (u) =>
                      u.floating &&
                      u.start[0].childNodes[u.start[1]] === floatingImages(view.dom)[picture.index],
                  );
                  if (!token) continue;
                  const g = picture.geometry;
                  const img = token.start[0].childNodes[token.start[1]] as HTMLImageElement;
                  const top = hwpPaperCoverEditorTop(picture, img, view.dom) ??
                    (paragraph.startPage ? picture.cssTop : isolatedPagePictureTop(picture, img, view, position));
                  const translation = picture.cssTranslation ? `;translate:${picture.cssTranslation}` : "";
                  const style = `position:absolute;float:none;margin:0;left:${picture.cssLeft}px;top:${top}px;z-index:${g.behindDocument ? -1 : 1}${picture.offPage ? ";display:none" : ""}${picture.unconstrainedWidth ? ";max-width:none" : ""}${translation}`;
                  decorations.push(
                    Decoration.node(position + 1 + token.from, position + 1 + token.to,
                      { style, ...(picture.unconstrainedWidth ? { 'data-hwp-full-width-picture': 'true' } : {}) }),
                  );
                  values.push(`image:${position + token.from}:${style}`);
                }
                for (const [order, widget] of paragraph.widgets.entries()) {
                  const at = position + 1 + widget.offset;
                  decorations.push(
                    Decoration.widget(at, () => widgetDOM(widget), {
                      side: -10000 + order,
                      key: `${position}:${order}:${JSON.stringify(widget)}`,
                    }),
                  );
                  values.push(`${at}:${JSON.stringify(widget)}`);
                }
              }
              const next = values.join("|");
              if (next !== signature || lastDocument !== view.state.doc) {
                signature = next;
                lastDocument = view.state.doc;
                const hadFocus=view.hasFocus();
                view.dispatch(
                  view.state.tr
                    .setMeta(key, DecorationSet.create(view.state.doc, decorations))
                    .setMeta("addToHistory", false),
                );
                // Rebuilding inline flow widgets can leave Chromium's native
                // range in a removed/split text node while the model caret is
                // already at the document end. Reapply the current selection
                // without stealing focus from a toolbar, dialog or IME.
                if(hadFocus && !view.composing)view.focus();
              }
            });
          };
          const invalidate=()=>{flowDocument=undefined;calculate();};
          document.addEventListener(editorLayoutResume, calculate);
          const observer = new ResizeObserver(calculate);
          observer.observe(view.dom);
          view.dom.addEventListener("load", invalidate, true);
          view.dom.addEventListener("compositionend", calculate);
          document.fonts.addEventListener("loadingdone", invalidate);
          document.addEventListener(wordFontMetricsChanged, invalidate);
          calculate();
          return {
            update: calculate,
            destroy() {
              cancelAnimationFrame(frame);
              observer.disconnect();
              document.removeEventListener(editorLayoutResume, calculate);
              view.dom.removeEventListener("load", invalidate, true);
              view.dom.removeEventListener("compositionend", calculate);
              document.fonts.removeEventListener("loadingdone", invalidate);
              document.removeEventListener(wordFontMetricsChanged, invalidate);
            },
          };
        },
      }),
    ];
  },
});
