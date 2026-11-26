import type { HwpDocument } from '@rhwp/core';
import { strFromU8, strToU8, zipSync } from 'fflate';
import { boundedNativeParts, nativeBodyRuns, nativePaintedControls } from './export-binary-pages';
import { elementRanges } from './markup-ranges';

export const nativeStandalonePictureAttribute = 'data-native-standalone-picture-line';

export type StandalonePictureMeasurement = { source: string; before: number };
export type StandalonePictureAsset = { asset: string; before: number };

export function measuredStandalonePictureSources(body: Element, paper: DOMRect) {
  const sources: StandalonePictureMeasurement[] = [];
  const bodyTop = body.getBoundingClientRect().top;
  for (const image of body.querySelectorAll<HTMLImageElement>('img:not([data-image-floating])')) {
    const paragraph = image.parentElement;
    if (paragraph?.tagName !== 'P' || paragraph.children.length !== 1 || paragraph.textContent?.trim()) continue;
    const column = paragraph.parentElement;
    if (column !== body && !(column?.parentElement === body && column.matches('[data-column-count]'))) continue;
    const source = image.getAttribute('src');
    if (!source || !/^data:image\/(?:bmp|png|jpeg);base64,/i.test(source) || !image.naturalWidth || !image.naturalHeight) continue;
    const box = image.getBoundingClientRect();
    if (!(box.width > 0 && box.height > 0 && box.top >= paper.top && box.bottom <= paper.bottom &&
        box.left >= paper.left && box.right <= paper.right)) continue;
    const margin = parseFloat(paragraph.style.marginTop);
    const offset = box.top - bodyTop;
    const before = column === body && paragraph === body.firstElementChild && margin > 0 &&
      Math.abs(offset - margin) <= .1 ? Math.round(offset * 75) : 0;
    sources.push({ source, before });
  }
  return sources;
}

const same = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b);
type PictureLine = { section: number; paragraph: number; page: number; width: number; height: number; before: number };

function missingPictureLine(doc: HwpDocument, section: number, paragraph: number, before: number) {
  const metrics = JSON.parse(doc.getParagraphLineMetrics(section, paragraph, '[]'));
  if (metrics.stored || metrics.lines?.length || metrics.controlCount !== 1 || doc.getParagraphLength(section, paragraph)) return;
  const picture = JSON.parse(doc.getPictureProperties(section, paragraph, 0));
  const format = JSON.parse(doc.getParaPropertiesAt(section, paragraph));
  const font = JSON.parse(doc.getCharPropertiesAt(section, paragraph, 0)).fontSize;
  const { width, height } = picture;
  const gap = Math.round(font * (format.lineSpacing / 100 - 1));
  if (!picture.treatAsChar || picture.hasCaption || picture.rotationAngle || format.lineSpacingType !== 'Percent' ||
      ![width, height, font, gap].every(Number.isSafeInteger) || width <= 0 || height <= 0 || font <= 0 || gap < 0) return;
  const { page } = JSON.parse(doc.getPageOfPosition(section, paragraph));
  if (!Number.isInteger(page) || page < 0 || page >= doc.pageCount() ||
      nativePaintedControls(doc, page).some(control => control.secIdx === section && control.paraIdx === paragraph)) return;
  // A single line has both first and last segment flags. Its height comes only
  // from the stored picture and the original paragraph spacing.
  const lineHeight = height + before;
  const line = `<hp:linesegarray><hp:lineseg textpos="0" vertpos="0" vertsize="${lineHeight}" textheight="${lineHeight}" baseline="${Math.round(lineHeight / 2)}" spacing="${gap}" horzpos="0" horzsize="${width}" flags="393216"/></hp:linesegarray>`;
  return { proof: { section, paragraph, page, width, height, before } as PictureLine, line };
}

function tryRestoreStandalonePictureLines(Core: typeof HwpDocument, document: HwpDocument, assetIds: readonly StandalonePictureAsset[]) {
  if (!assetIds.length) return document;
  let source: HwpDocument | undefined, result: HwpDocument | undefined;
  try {
    const bytes = document.exportHwpx(), archive = boundedNativeParts(bytes);
    const assets = new Map(assetIds.filter(({ asset }) => /^image\d+$/.test(asset) &&
      ['bmp', 'png', 'jpeg', 'jpg'].some(extension => !!archive[`BinData/${asset}.${extension}`]))
      .map(({ asset, before }) => [asset, before]));
    if (!assets.size) return document;
    const previous = source = new Core(bytes), pages = previous.pageCount();
    if (pages !== document.pageCount()) return document;
    const proofs: PictureLine[] = [];
    const seen = new Set<string>();
    for (const [name, data] of Object.entries(archive)) {
      const filename = /^Contents\/section(\d+)\.xml$/.exec(name);
      if (!filename) continue;
      const section = Number(filename[1]);
      let xml = strFromU8(data);
      const paragraphs = elementRanges(xml, 'hp:p');
      for (let paragraph = paragraphs.length - 1; paragraph >= 0; paragraph--) {
        const range = paragraphs[paragraph];
        if (range.inner.includes('<hp:linesegarray')) continue;
        const pictures = elementRanges(range.inner, 'hp:pic');
        if (pictures.length !== 1) continue;
        const asset = /<hc:img\b[^>]*\bbinaryItemIDRef="([^"]+)"/.exec(pictures[0].inner)?.[1];
        if (!asset) continue;
        const before = assets.get(asset);
        if (before === undefined) continue;
        const repair = missingPictureLine(previous, section, paragraph, before);
        if (!repair) continue;
        if (seen.has(asset)) return document;
        seen.add(asset); proofs.push(repair.proof);
        let content = range.inner;
        if (before) {
          const picture = pictures[0];
          const sized = picture.inner.replace(/<hp:outMargin\b[^>]*\/>/, tag => tag.replace(/\btop="0"/, `top="${before}"`));
          if (sized === picture.inner) return document;
          const at = picture.start + picture.open.length;
          content = content.slice(0, at) + sized + content.slice(at + picture.inner.length);
        }
        const at = range.start + range.open.length;
        xml = xml.slice(0, at) + content + repair.line + xml.slice(range.end - '</hp:p>'.length);
      }
      archive[name] = strToU8(xml);
    }
    if (!proofs.length) return document;
    const candidate = result = new Core(zipSync(archive, { level: 1 }));
    if (candidate.pageCount() !== pages) return document;
    const byPage = new Map<number, PictureLine[]>();
    for (const proof of proofs) {
      const list = byPage.get(proof.page);
      if (list) list.push(proof); else byPage.set(proof.page, [proof]);
    }
    for (let page = 0; page < pages; page++) {
      const controls = nativePaintedControls(candidate, page);
      const expected = byPage.get(page);
      if (expected) {
        const paper = JSON.parse(candidate.getPageInfo(page));
        for (const proof of expected) {
          const index = controls.findIndex(control => control.type === 'image' && control.secIdx === proof.section &&
            control.paraIdx === proof.paragraph && control.controlIdx === 0 && !control.cellPath?.length);
          if (index < 0) return document;
          const [image] = controls.splice(index, 1);
          const bodyTop = paper.marginTop + paper.marginHeader;
          if (proof.before && (Math.abs(image.y - bodyTop - proof.before / 75) > .12 ||
              image.y + image.h > paper.footerArea.y)) return document;
          if (Math.abs(image.w - proof.width / 75) > .1 || Math.abs(image.h - proof.height / 75) > .1 ||
              image.x < 0 || image.y < 0 || image.x + image.w > paper.width || image.y + image.h > paper.height) return document;
        }
      }
      const ink = (doc: HwpDocument) => nativeBodyRuns(doc, page).filter(run => run.text.trim());
      if (!same(ink(previous), ink(candidate)) || !same(nativePaintedControls(previous, page), controls)) return document;
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


/** Restore measured inline pictures. Keep the ordinary picture line if the
 * optional space above an image would change the rest of the native document. */
export function restoreStandalonePictureLines(Core: typeof HwpDocument, document: HwpDocument, assetIds: readonly StandalonePictureAsset[]) {
  const positioned = tryRestoreStandalonePictureLines(Core, document, assetIds);
  return positioned === document && assetIds.some(({ before }) => before > 0)
    ? tryRestoreStandalonePictureLines(Core, document, assetIds.map(({ asset }) => ({ asset, before: 0 }))) : positioned;
}
