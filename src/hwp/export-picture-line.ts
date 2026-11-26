import type { HwpDocument } from '@rhwp/core';
import { strFromU8, strToU8, zipSync } from 'fflate';
import { boundedNativeParts, nativeBodyRuns, nativePaintedControls } from './export-binary-pages';
import { elementRanges } from './markup-ranges';

export const nativeStandalonePictureAttribute = 'data-native-standalone-picture-line';

export function measuredStandaloneBitmapSources(body: Element, paper: DOMRect) {
  const sources: string[] = [];
  for (const image of body.querySelectorAll<HTMLImageElement>('img:not([data-image-floating])')) {
    const paragraph = image.parentElement;
    if (paragraph?.tagName !== 'P' || paragraph.children.length !== 1 || paragraph.textContent?.trim()) continue;
    const column = paragraph.parentElement;
    if (column !== body && !(column?.parentElement === body && column.matches('[data-column-count]'))) continue;
    const source = image.getAttribute('src');
    if (!source?.startsWith('data:image/bmp;base64,') || !image.naturalWidth || !image.naturalHeight) continue;
    const box = image.getBoundingClientRect();
    if (box.width > 0 && box.height > 0 && box.top >= paper.top && box.bottom <= paper.bottom &&
        box.left >= paper.left && box.right <= paper.right) sources.push(source);
  }
  return sources;
}

const same = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b);

/** A single inline picture needs a real line in an otherwise empty native
 * paragraph. Give the observed, restored bitmap a line without adding text or
 * changing its inline anchoring. Keep the normal save if the line reflows ink. */
export function restoreStandalonePictureLine(Core: typeof HwpDocument, document: HwpDocument, assetIds: readonly string[]) {
  if (assetIds.length !== 1) return document;
  let source: HwpDocument | undefined, result: HwpDocument | undefined;
  try {
    const [asset] = assetIds, name = 'Contents/section0.xml', bytes = document.exportHwpx();
    const archive = boundedNativeParts(bytes);
    if (!archive[name] || !archive[`BinData/${asset}.bmp`]) return document;
    const xml = strFromU8(archive[name]), paragraphs = elementRanges(xml, 'hp:p');
    const matching = paragraphs.flatMap((range, index) => {
      const solePicture = /^<hp:run\b[^>]*><hp:pic\b[^>]*>[\s\S]*?<\/hp:pic><\/hp:run>$/.test(range.inner);
      return solePicture && range.inner.includes(`binaryItemIDRef="${asset}"`) ? [{range, index}] : [];
    });
    if (matching.length !== 1) return document;
    const [{range, index}] = matching;
    const previous = source = new Core(bytes), pages = previous.pageCount();
    const metrics = JSON.parse(previous.getParagraphLineMetrics(0, index, '[]'));
    const picture = JSON.parse(previous.getPictureProperties(0, index, 0));
    const paragraph = JSON.parse(previous.getParaPropertiesAt(0, index));
    const font = JSON.parse(previous.getCharPropertiesAt(0, index, 0)).fontSize;
    const {width, height} = picture;
    const gap = Math.round(font * (paragraph.lineSpacing / 100 - 1));
    if (pages !== document.pageCount() || metrics.stored || metrics.controlCount !== 1 ||
        !picture.treatAsChar || picture.hasCaption || picture.rotationAngle ||
        paragraph.lineSpacingType !== 'Percent' || ![width, height, font, gap].every(Number.isSafeInteger) ||
        width <= 0 || height <= 0 || font <= 0 || gap < 0) return document;
    const at = range.end - '</hp:p>'.length;
    const flags = 0x60000; // First and last segment in this line.
    const line = `<hp:linesegarray><hp:lineseg textpos="0" vertpos="0" vertsize="${height}" textheight="${height}" baseline="${Math.round(height / 2)}" spacing="${gap}" horzpos="0" horzsize="${width}" flags="${flags}"/></hp:linesegarray>`;
    archive[name] = strToU8(xml.slice(0, at) + line + xml.slice(at));
    const candidate = result = new Core(zipSync(archive, {level:1}));
    if (candidate.pageCount() !== pages) return document;
    const controls = nativePaintedControls(candidate, 0);
    const image = controls.find(control => control.type === 'image' && control.secIdx === 0 &&
      control.paraIdx === index && control.controlIdx === 0 && !control.cellPath?.length);
    const paper = JSON.parse(candidate.getPageInfo(0));
    if (!image || Math.abs(image.w - width / 75) > .1 || Math.abs(image.h - height / 75) > .1 ||
        image.x < 0 || image.y < 0 || image.x + image.w > paper.width || image.y + image.h > paper.height ||
        !same(nativePaintedControls(previous, 0), controls.filter(control => control !== image))) return document;
    const ink = (doc: HwpDocument, page: number) => nativeBodyRuns(doc, page).filter(run => run.text.trim());
    for (let page = 0; page < pages; page++) {
      if (!same(ink(previous, page), ink(candidate, page)) ||
          (page && !same(nativePaintedControls(previous, page), nativePaintedControls(candidate, page)))) return document;
    }
    result = undefined;
    return candidate;
  } catch {
    return document;
  } finally {
    source?.free();
    result?.free();
  }
}
