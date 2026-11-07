import type { HwpDocument } from '@rhwp/core';

const blank = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQIHWP4z8DwHwAFgAI/ScLbtAAAAABJRU5ErkJggg==';
const limit = 4 * 1024 * 1024;
const attribute = (text: string) => text.replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]!);

/** The native clipboard bridge labels EMF as octet-stream. Convert the asset
 * alone with the same native SVG renderer used by the reader, never a page
 * screenshot or a similarly positioned picture from the original document.
 * MS-EMF 2.3.4.2: EMR_HEADER type 1, ENHMETA_SIGNATURE at offset 40. */
export function metafilePreview(html: string, source: HwpDocument, cache: Map<string, string>): string {
  return html.replace(/<img\b[^>]*>/g, image => {
    const encoded = image.match(/\bsrc="data:(?:application\/octet-stream|image\/(?:x-)?emf);base64,([A-Za-z0-9+/=]+)"/);
    if (!encoded) return image;
    const head = Uint8Array.from(atob(encoded[1].slice(0, 128)), c => c.charCodeAt(0));
    if (head.length < 88) return image;
    const view = new DataView(head.buffer);
    if (view.getUint32(0, true) !== 1 || view.getUint32(40, true) !== 0x464d4520) return image;
    if (encoded[1].length > Math.ceil(limit / 3) * 4) throw Error('벡터 그림이 너무 커서 안전하게 변환하지 못했습니다.');
    let svg = cache.get(encoded[1]);
    if (!svg) {
      const bytes = Uint8Array.from(atob(encoded[1]), c => c.charCodeAt(0));
      const header = view.getUint32(4, true);
      if (header < 88 || header > bytes.length || view.getUint32(48, true) !== bytes.length) throw Error('벡터 그림의 파일 구조가 올바르지 않습니다.');
      const isolated = (source.constructor as typeof HwpDocument).createEmpty();
      try {
        isolated.createBlankDocument();
        isolated.insertPicture(0, 0, 0, '[]', bytes, 15000, 15000, 200, 200, 'emf', 'EMF asset');
        const images = [...isolated.renderPageSvgWithProfile(0, 'print').matchAll(/<image\b[^>]*>/g)];
        const embedded = images.length === 1 && images[0][0].match(/\bhref="data:image\/svg\+xml;base64,([A-Za-z0-9+/=]+)"/);
        if (!embedded || embedded[1].length > Math.ceil(limit / 3) * 4) throw Error('벡터 그림의 모양을 변환하지 못했습니다.');
        svg = new TextDecoder().decode(Uint8Array.from(atob(embedded[1]), c => c.charCodeAt(0)));
        if (!/^<svg\b/.test(svg)) throw Error('벡터 그림의 모양을 변환하지 못했습니다.');
        if (cache.size >= 128) throw Error('벡터 그림이 너무 많습니다.');
        cache.set(encoded[1], svg);
      } finally { isolated.free(); }
    }
    return image.replace(encoded[0], `src="${blank}" data-hwp-metafile="true" data-hwp-svg="${attribute(svg)}"`);
  });
}
