import type { HwpDocument } from '@rhwp/core';
import { strFromU8, unzipSync } from 'fflate';
import { elementRanges } from './markup-ranges';

/** Older binary pictures omit originalSizeHu in the drawing API. Resolve their
 * declared creation size from the normalized source, requiring exact asset
 * bytes and one unambiguous size. Never infer a size from the page bbox. */
export function objectImageSizes(doc: HwpDocument) {
  let metadataBytes = 0, imageBytes = 0;
  const parts = unzipSync(doc.exportHwpx(), { filter: file => {
    if (/^Contents\/(?:section\d+\.xml|content\.hpf)$/.test(file.name)) {
      if (file.originalSize > 32 * 1024 * 1024 || (metadataBytes += file.originalSize) > 96 * 1024 * 1024) throw Error('Object image metadata too large');
      return true;
    }
    if (!/^BinData\/[^/]+\.(?:png|jpe?g|gif|webp|bmp)$/i.test(file.name) || file.originalSize > 4 * 1024 * 1024 || imageBytes + file.originalSize > 4 * 1024 * 1024) return false;
    imageBytes += file.originalSize; return true;
  } });
  const attr = (tag: string, name: string) => tag.match(new RegExp(`\\b${name}="([^"]*)"`))?.[1] ?? '';
  const assets = new Map<string, Uint8Array>();
  for (const item of strFromU8(parts['Contents/content.hpf'] ?? new Uint8Array()).matchAll(/<opf:item\b[^>]*>/g)) {
    const href = attr(item[0], 'href');
    if (/^BinData\/[^/]+$/.test(href) && parts[href]) assets.set(attr(item[0], 'id'), parts[href]);
  }
  const pictures: { bytes: Uint8Array; size: [number, number] }[] = [];
  for (const [name, bytes] of Object.entries(parts)) if (/^Contents\/section\d+\.xml$/.test(name)) {
    for (const pic of elementRanges(strFromU8(bytes), 'hp:pic')) {
      const image = pic.inner.match(/<hc:img\b[^>]*>/)?.[0] ?? '', asset = assets.get(attr(image, 'binaryItemIDRef'));
      const dim = pic.inner.match(/<hp:imgDim\b[^>]*>/)?.[0] ?? '', org = pic.inner.match(/<hp:orgSz\b[^>]*>/)?.[0] ?? '';
      let size = [Number(attr(dim, 'dimwidth')), Number(attr(dim, 'dimheight'))];
      if (!size.every(n => Number.isFinite(n) && n > 0)) size = [Number(attr(org, 'width')), Number(attr(org, 'height'))];
      if (asset && size.every(n => Number.isFinite(n) && n > 0 && n <= 1_000_000)) pictures.push({ bytes: asset, size: size as [number, number] });
      if (pictures.length > 10000) throw Error('Too many object image sizes');
    }
  }
  return (bytes: Uint8Array): [number, number] | undefined => {
    let found: [number, number] | undefined;
    for (const pic of pictures) {
      if (pic.bytes.length !== bytes.length || !pic.bytes.every((value, i) => value === bytes[i])) continue;
      if (found && (found[0] !== pic.size[0] || found[1] !== pic.size[1])) return;
      found = pic.size;
    }
    return found;
  };
}
