import {drawingGroup,type DrawingGroup} from './drawing-group';
import type { IFloating } from "docx";
import type JSZip from "jszip";
import { hwpCellRasterFillNamespace, hwpRasterPictureFilter, isHwpCellRasterBrightness } from './hwp-cell-raster-fill';
export type HwpBodyPictureFlow = { anchorY: number; height: number; followingEm: number };
type HwpPaperDrawing = { width: number; height: number };
export type ImageFloating = IFloating & {
  polygon?: [number, number][];
  drawingGroup?: DrawingGroup;
  hwpBodyFlow?: HwpBodyPictureFlow;
  hwpPaperDrawing?: HwpPaperDrawing;
  hwpCellRasterFill?: true;
  hwpCellRasterBrightness?: number;
  hwpRasterPictureBrightness?: number;
};
const hwpBodyFlowNamespace = 'urn:wordstudio:hwp-body-picture-flow:1';
const compatibilityNamespace = 'http://schemas.openxmlformats.org/markup-compatibility/2006';

const drawingNamespace = 'http://schemas.openxmlformats.org/drawingml/2006/main';

function cellRasterBlip(anchor: Element) {
  const [image, duplicate] = [...anchor.getElementsByTagNameNS(drawingNamespace, 'blip')];
  return duplicate ? undefined : image;
}

function cellRasterBrightness(anchor: Element): number | undefined {
  const image = cellRasterBlip(anchor);
  if (!image) return;
  const [effect, extra] = [...image.children];
  if (!effect) return 0;
  if (extra || effect.namespaceURI !== drawingNamespace || effect.localName !== 'lum') return;
  const raw = effect.getAttribute('bright') ?? '0', contrast = effect.getAttribute('contrast') ?? '0';
  const brightness = Number(raw) / 1000;
  if (contrast === '0' && /^-?\d{1,6}$/.test(raw) && isHwpCellRasterBrightness(brightness)) return brightness;
}

function cellRasterSignature(anchor: Element) {
  const child = (name: string) => [...anchor.children].find(node => node.localName === name);
  const extent = child('extent');
  const offset = (name: string) => [...(child(name)?.children ?? [])].find(el => el.localName === 'posOffset')?.textContent;
  const dimensions = [extent?.getAttribute('cx'), extent?.getAttribute('cy'), offset('positionH'), offset('positionV')];
  return dimensions.every(value => /^-?\d{1,10}$/.test(value || '')) ? dimensions.join(':') : undefined;
}
function hwpBodyPictureFlow(input: unknown): HwpBodyPictureFlow | undefined {
  if (!input || typeof input !== 'object') return;
  const { anchorY, height, followingEm } = input as HwpBodyPictureFlow;
  if (![anchorY, height, followingEm].every(n => typeof n === 'number' && Number.isFinite(n)) ||
      anchorY < 0 || anchorY > 3000 || height < 1 || height > 3000 ||
      followingEm < 4 || followingEm > 120 || followingEm >= height) return;
  return { anchorY, height, followingEm };
}

function hwpPaperDrawing(input: unknown): HwpPaperDrawing | undefined {
  if (!input || typeof input !== 'object') return;
  const { width, height } = input as HwpPaperDrawing;
  if (![width, height].every(n => typeof n === 'number' && Number.isFinite(n) && n > 0 && n <= 4000)) return;
  return { width, height };
}

export type ImageGeometryRecord = { name: string; geometry: ImageFloating };

/** docx emits a rectangular wrapPolygon. Replace it with the document's
 * contour in every story and remove the temporary matching name. */
export async function restoreImageGeometry(zip: JSZip, records: Map<string, ImageGeometryRecord>) {
  const ns = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
  for (const file of Object.keys(zip.files).filter((name) => /^word\/.*\.xml$/.test(name))) {
    const source = await zip.file(file)!.async("string");
    if (![...records.keys()].some((key) => source.includes(key))) continue;
    const xml = new DOMParser().parseFromString(source, "application/xml");
    for (const properties of xml.getElementsByTagNameNS(ns, "docPr")) {
      const record = records.get(properties.getAttribute("name") || "");
      if (!record) continue;
      properties.setAttribute("name", record.name);
      if(record.geometry.drawingGroup)properties.parentElement!.setAttribute("data-ws-drawing-group",JSON.stringify(record.geometry.drawingGroup));
      const anchor = properties.parentElement;
      const brightness = record.geometry.hwpCellRasterFill
        ? record.geometry.hwpCellRasterBrightness : record.geometry.hwpRasterPictureBrightness;
      if (brightness && anchor) {
        const image = cellRasterBlip(anchor);
        if (image && !image.children.length) {
          const luminance = xml.createElementNS(drawingNamespace, 'a:lum');
          luminance.setAttribute('bright', String(brightness * 1000));
          luminance.setAttribute('contrast', '0');
          image.append(luminance);
        }
      }
      const sourceGeometry = record.geometry.hwpBodyFlow
        ? { ns: hwpBodyFlowNamespace, prefix: 'wspf', name: 'clearance', value: record.geometry.hwpBodyFlow }
        : record.geometry.hwpPaperDrawing
          ? { ns: hwpBodyFlowNamespace, prefix: 'wspf', name: 'paper', value: record.geometry.hwpPaperDrawing }
          : record.geometry.hwpCellRasterFill && anchor
            ? { ns: hwpCellRasterFillNamespace, prefix: 'wscf', name: 'fill', value: cellRasterSignature(anchor) }
            : record.geometry.hwpRasterPictureBrightness
              ? { ns: hwpCellRasterFillNamespace, prefix: 'wscf', name: 'picture', value: true } : undefined;
      if (anchor?.localName === 'anchor' && sourceGeometry?.value !== undefined) {
        const { ns, prefix, name, value } = sourceGeometry;
        anchor.setAttributeNS('http://www.w3.org/2000/xmlns/', 'xmlns:' + prefix, ns);
        anchor.setAttributeNS(ns, prefix + ':' + name, JSON.stringify(value));
        for (const [attribute, value] of [['Ignorable',prefix], ['PreserveAttributes',prefix + ':' + name]]) {
          const earlier = (anchor.getAttributeNS(compatibilityNamespace, attribute) || '').trim();
          anchor.setAttributeNS(compatibilityNamespace, 'mc:' + attribute, earlier ? earlier + ' ' + value : value);
        }
      }
      const wrap = anchor && [...anchor.children].find(el => el.localName === 'wrapTight' || el.localName === 'wrapThrough');
      if (!wrap || !record.geometry.polygon) continue;
      const polygon = xml.createElementNS(ns, "wp:wrapPolygon");
      polygon.setAttribute("edited", "0");
      record.geometry.polygon.forEach(([x, y], index) => {
        const point = xml.createElementNS(ns, index ? "wp:lineTo" : "wp:start");
        point.setAttribute("x", String(Math.round(x * 21600)));
        point.setAttribute("y", String(Math.round(y * 21600)));
        polygon.append(point);
      });
      const previous = wrap.getElementsByTagNameNS(ns, "wrapPolygon")[0];
      if (previous) previous.replaceWith(polygon);
      else wrap.append(polygon);
    }
    zip.file(file, new XMLSerializer().serializeToString(xml));
  }
}

export function imageFloating(value: unknown): ImageFloating | undefined {
  if (typeof value !== "string" || value.length > 10000) return;
  try {
    const data = JSON.parse(value);
    const position = (p: any, vertical: boolean) =>
      p &&
      typeof p === "object" &&
      (p.offset === undefined ||
        (typeof p.offset === "number" &&
          Number.isFinite(p.offset) &&
          Math.abs(p.offset) <= 100000000)) &&
      (p.align === undefined ||
        (vertical
          ? ["top", "center", "bottom", "inside", "outside"]
          : ["left", "center", "right", "inside", "outside"]
        ).includes(p.align)) &&
      (p.relative === undefined ||
        (vertical
          ? [
              "page",
              "margin",
              "paragraph",
              "line",
              "topMargin",
              "bottomMargin",
              "insideMargin",
              "outsideMargin",
            ]
          : [
              "page",
              "margin",
              "column",
              "character",
              "leftMargin",
              "rightMargin",
              "insideMargin",
              "outsideMargin",
            ]
        ).includes(p.relative));
    if (
      !position(data.horizontalPosition, false) ||
      !position(data.verticalPosition, true) ||
      ![0, 1, 2, 3, 4].includes(data.wrap?.type)
    )
      return;
    const polygon =
      Array.isArray(data.polygon) &&
      data.polygon.length >= 3 &&
      data.polygon.length <= 256 &&
      data.polygon.every(
        (p: unknown) =>
          Array.isArray(p) &&
          p.length === 2 &&
          p.every((n) => typeof n === "number" && Number.isFinite(n) && Math.abs(n) <= 4),
      )
        ? (data.polygon as [number, number][])
        : undefined;
    if (data.polygon !== undefined && !polygon) return;
    const group=drawingGroup(data.drawingGroup);if(data.drawingGroup!==undefined && !group)return;
    const hwpBodyFlow = hwpBodyPictureFlow(data.hwpBodyFlow);
    const paperDrawing = hwpPaperDrawing(data.hwpPaperDrawing);
    const brightness = data.hwpCellRasterBrightness ?? 0;
    const cellRasterFill = isHwpCellRasterBrightness(brightness) && data.hwpCellRasterFill === true && data.behindDocument === true && data.layoutInCell === true &&
      data.wrap.type === 0 && data.horizontalPosition.relative === 'column' && data.verticalPosition.relative === 'paragraph' &&
      !group && !polygon && !paperDrawing && !hwpBodyFlow;
    const pictureBrightness = data.hwpRasterPictureBrightness;
    return {
      horizontalPosition: data.horizontalPosition,
      verticalPosition: data.verticalPosition,
      wrap: {
        type: data.wrap.type,
        ...(["bothSides", "left", "right", "largest"].includes(data.wrap.side)
          ? { side: data.wrap.side }
          : {}),
      },
      ...(polygon ? { polygon } : {}),
      ...(group ? {drawingGroup:group} : {}),
      ...(hwpBodyFlow ? { hwpBodyFlow } : {}),
      ...(paperDrawing ? { hwpPaperDrawing: paperDrawing } : {}),
      ...(cellRasterFill ? { hwpCellRasterFill: true as const, ...(brightness && { hwpCellRasterBrightness: brightness }) } : {}),
      ...(!cellRasterFill && !group && !paperDrawing && !hwpBodyFlow && pictureBrightness &&
        isHwpCellRasterBrightness(pictureBrightness) ? { hwpRasterPictureBrightness: pictureBrightness } : {}),
      ...(data.margins && typeof data.margins === "object"
        ? {
            margins: Object.fromEntries(
              ["top", "right", "bottom", "left"].map((k) => [
                k,
                Math.max(0, Math.min(10000000, Number(data.margins[k]) || 0)),
              ]),
            ),
          }
        : {}),
      behindDocument: data.behindDocument === true,
      allowOverlap: true,
      layoutInCell: data.layoutInCell !== false,
    };
  } catch {
    return;
  }
}

export function floatingCss(value: unknown, width: number) {
  const data = imageFloating(value);
  if (!data) return "";
  const horizontal = data.horizontalPosition,
    vertical = data.verticalPosition;
  const left =
    horizontal.align === "center"
      ? `calc(50% - ${width / 2}px)`
      : horizontal.align === "right"
        ? `calc(100% - ${width}px)`
        : `${(horizontal.offset || 0) / 9525}px`;
  const paperOverlay = data.layoutInCell === false && data.wrap?.type === 0 && !data.behindDocument && !data.drawingGroup &&
    horizontal.relative === "page" && vertical.relative === "page";
  const filter = hwpRasterPictureFilter(data.hwpRasterPictureBrightness ?? 0);
  return `position:absolute;float:none;left:${left};top:${vertical.relative === "paragraph" ? (vertical.offset || 0) / 9525 : 0}px;margin:0;z-index:${data.behindDocument ? -1 : 1}${paperOverlay ? ";--table-cell-image-max-width:none" : ""}${filter ? `;filter:${filter}` : ""}`;
}

export function drawingFloating(drawing: Element) {
  const normalized = drawing.getAttribute("data-ws-floating");
  if (normalized) return imageFloating(normalized);
  const anchor = [...drawing.children].find((el) => el.localName === "anchor");
  if (!anchor) return;
  const read = (kind: string) => {
    const p = [...anchor.children].find((el) => el.localName === kind);
    const align = [...(p?.children || [])].find((el) => el.localName === "align")?.textContent;
    const offset = Number(
      [...(p?.children || [])].find((el) => el.localName === "posOffset")?.textContent || 0,
    );
    return {
      relative: p?.getAttribute("relativeFrom") || (kind === "positionH" ? "column" : "paragraph"),
      ...(align ? { align } : { offset }),
    };
  };
  const type = [...anchor.children].find((el) => /^wrap/.test(el.localName))?.localName;
  const wrapping = [...anchor.children].find((el) => el.localName === type);
  const storedFlow = anchor.getAttributeNS(hwpBodyFlowNamespace, 'clearance');
  let hwpBodyFlow: HwpBodyPictureFlow | undefined;
  let paperDrawing: HwpPaperDrawing | undefined;
  if (storedFlow && storedFlow.length < 120) try {
    hwpBodyFlow = hwpBodyPictureFlow(JSON.parse(storedFlow));
  } catch { /* Optional source geometry must not prevent drawing import. */ }
  const storedPaperDrawing = anchor.getAttributeNS(hwpBodyFlowNamespace, 'paper');
  if (storedPaperDrawing && storedPaperDrawing.length < 120) try {
    paperDrawing = hwpPaperDrawing(JSON.parse(storedPaperDrawing));
  } catch { /* Optional source geometry must not prevent drawing import. */ }
  const points = [...(wrapping?.getElementsByTagNameNS("*", "wrapPolygon")[0]?.children || [])].map(
    (p) => [Number(p.getAttribute("x")) / 21600, Number(p.getAttribute("y")) / 21600],
  );
  const horizontalPosition = read("positionH"), verticalPosition = read("positionV");
  const signature = cellRasterSignature(anchor);
  const previousFill = anchor.getAttributeNS(hwpCellRasterFillNamespace, 'fill');
  const remainsFill = signature && previousFill === JSON.stringify(signature) && type === 'wrapNone' &&
    anchor.getAttribute('behindDoc') === '1' && anchor.getAttribute('layoutInCell') === '1' &&
    horizontalPosition.relative === 'column' && verticalPosition.relative === 'paragraph';
  // Keep the standard luminance when an external edit turns the cell fill into a normal picture.
  const scopedPicture = anchor.getAttributeNS(hwpCellRasterFillNamespace, 'picture') === 'true' ||
    /^"-?\d{1,10}(?::-?\d{1,10}){3}"$/.test(previousFill ?? '');
  const rasterBrightness = remainsFill || scopedPicture ? cellRasterBrightness(anchor) : undefined;
  return imageFloating(
    JSON.stringify({
      horizontalPosition,
      verticalPosition,
      wrap: {
        type:
          (
            {
              wrapNone: 0,
              wrapSquare: 1,
              wrapTight: 2,
              wrapTopAndBottom: 3,
              wrapThrough: 4,
            } as Record<string, number>
          )[type || ""] ?? 1,
        side: wrapping?.getAttribute("wrapText"),
      },
      ...(points.length >= 3 ? { polygon: points } : {}),
      margins: Object.fromEntries(
        ["top", "right", "bottom", "left"].map((k) => [
          k,
          Number(anchor.getAttribute(`dist${k[0].toUpperCase()}`)) || 0,
        ]),
      ),
      behindDocument: anchor.getAttribute("behindDoc") === "1",
      layoutInCell: !["0", "false", "off"].includes(anchor.getAttribute("layoutInCell") || ""),
      ...(hwpBodyFlow ? { hwpBodyFlow } : {}),
      ...(paperDrawing ? { hwpPaperDrawing: paperDrawing } : {}),
      ...(remainsFill && rasterBrightness !== undefined ? { hwpCellRasterFill: true, hwpCellRasterBrightness: rasterBrightness }
        : rasterBrightness ? { hwpRasterPictureBrightness: rasterBrightness } : {}),
      ...(drawing.hasAttribute("data-ws-drawing-group")?{drawingGroup:JSON.parse(drawing.getAttribute("data-ws-drawing-group")!)}:{}),
    }),
  );
}

/** Keep off-paper drawings and their authored coordinates in the document.
 * A rectangle wholly outside its sheet has no paper ink or text-wrap budget;
 * a partly visible or paragraph-relative drawing follows its normal layout. */
function rawOutsidePaper(img: HTMLElement, g: ImageFloating): boolean {
  const sheet = img.closest<HTMLElement>('[data-page-height],.print-page');
  if (!sheet) return false;
  const width = Number(img.getAttribute('width')), height = Number(img.getAttribute('height'));
  const style = getComputedStyle(sheet), paperWidth = parseFloat(style.width);
  const paperHeight = Number(sheet.dataset.pageHeight) || parseFloat(style.height);
  if (!(width > 0 && height > 0 && paperWidth > 0 && paperHeight > 0)) return false;
  const horizontal = g.horizontalPosition, vertical = g.verticalPosition;
  const x = (horizontal.offset || 0) / 9525, y = (vertical.offset || 0) / 9525;
  return (horizontal.relative === 'page' && !horizontal.align && (x >= paperWidth || x + width <= 0)) ||
    (vertical.relative === 'page' && !vertical.align && (y >= paperHeight || y + height <= 0));
}

/** A fitting foreground drawing in a single-paragraph cell stays in that
 * cell even when Word's stored page-relative X lies outside the paper.
 * Keep unsupported multi-block/group layouts on their existing path. */
export function recoverableCellImage(img: HTMLElement): HTMLElement | null {
  const g = imageFloating(img.dataset.imageFloating);
  const cell = img.closest<HTMLElement>('td,th');
  const p = img.closest<HTMLElement>('p,h1,h2,h3,h4,h5,h6');
  if (!g || g.layoutInCell !== true || !g.wrap?.type || g.behindDocument ||
      g.drawingGroup || !cell || !p || !rawOutsidePaper(img, g)) return null;

  const blocks = [...cell.children].filter(n => !n.classList.contains('ProseMirror-widget'));
  const width = Number(img.getAttribute('width'));
  const height = Number(img.getAttribute('height'));
  const sheet = img.closest<HTMLElement>('[data-page-height],.print-page');
  if (!sheet) return null;
  const style = getComputedStyle(sheet);
  const paperHeight = Number(sheet.dataset.pageHeight) || parseFloat(style.height);
  const paperWidth = parseFloat(style.width);
  const horizontal = g.horizontalPosition, vertical = g.verticalPosition;
  const x = (horizontal.offset || 0) / 9525, y = (vertical.offset || 0) / 9525;
  return blocks.length === 1 && blocks[0] === p &&
    cell.querySelectorAll('img[data-image-floating]').length === 1 && width > 0 && height > 0 &&
    horizontal.relative === 'page' && !horizontal.align && (x >= paperWidth || x + width <= 0) &&
    vertical.relative === 'page' && !vertical.align && y >= 0 &&
    width <= p.clientWidth + .5 && y + height <= paperHeight ? p : null;
}

export function imageOutsidePaper(img: HTMLElement): boolean {
  const g = imageFloating(img.dataset.imageFloating);
  return !!g && rawOutsidePaper(img, g) && !recoverableCellImage(img);
}
