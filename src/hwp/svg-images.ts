import DOMPurify from 'dompurify';
import type { EditableHwp } from './conversion';
import type { HwpProgressReporter } from './progress';
import { RASTER_DATA } from '../document-html';
import { rasterizeHwpObject } from './object-preview';
import type { PageSettings } from '../types';

/** Rasterize only the parser's explicit equation and shape previews. The common image
 * importer continues to reject arbitrary SVGs and external image resources. */
export async function prepareEditableHwp(source: EditableHwp, onProgress?: HwpProgressReporter, signal?: AbortSignal): Promise<EditableHwp> {
  const result = await prepareBody(source, onProgress, signal);
  const warnings = new Set(result.warnings);
  const preparePage = async (input: PageSettings) => {
    const page = { ...input };
    for (const key of ['headerHtml', 'footerHtml', 'firstHeaderHtml', 'firstFooterHtml', 'evenHeaderHtml', 'evenFooterHtml'] as const) {
      if (!page[key]) continue;
      const story = await prepareBody({ ...source, html: page[key]! }, onProgress, signal);
      page[key] = story.html; story.warnings.forEach(w => warnings.add(w));
    }
    return page;
  };
  const root = document.createElement('template'); root.innerHTML = result.html;
  for (const marker of root.content.querySelectorAll('[data-section-page],[data-section-content]')) {
    const attribute = marker.hasAttribute('data-section-page') ? 'data-section-page' : 'data-section-content';
    const page = await preparePage(JSON.parse(marker.getAttribute(attribute)!));
    marker.setAttribute(attribute, JSON.stringify(page));
  }
  return { ...result, html: root.innerHTML, page: await preparePage(result.page), warnings: [...warnings] };
}

async function prepareBody(source: EditableHwp, onProgress?: HwpProgressReporter, signal?: AbortSignal): Promise<EditableHwp> {
  signal?.throwIfAborted();
  if (!/data-hwp-(?:svg|object|zero-line|image-crop)=/.test(source.html)) return source;
  const root = document.createElement('template'); root.innerHTML = source.html;
  for (const image of root.content.querySelectorAll<HTMLImageElement>('img[data-hwp-image-crop]')) {
    signal?.throwIfAborted();
    const crop = image.getAttribute('data-hwp-image-crop')!.split(',').map(Number);
    if (crop.length !== 4 || !crop.every(n => Number.isFinite(n) && n >= 0 && n < 1) || crop[0] + crop[2] >= 1 || crop[1] + crop[3] >= 1 || !RASTER_DATA.test(image.src))
      throw Error('한글 그림의 자르기 정보가 올바르지 않습니다.');
    const bitmap = new Image(); bitmap.src = image.src; await bitmap.decode(); signal?.throwIfAborted();
    const width = Number(image.getAttribute('width')), height = Number(image.getAttribute('height'));
    if (!(width > 0 && height > 0 && width <= 10000 && height <= 10000)) throw Error('한글 그림의 크기가 올바르지 않습니다.');
    const scale = Math.min(2, Math.sqrt(4_000_000 / (width * height))), canvas = document.createElement('canvas');
    canvas.width = Math.max(1, Math.ceil(width * scale)); canvas.height = Math.max(1, Math.ceil(height * scale));
    canvas.getContext('2d')!.drawImage(bitmap, crop[0] * bitmap.naturalWidth, crop[1] * bitmap.naturalHeight,
      (1 - crop[0] - crop[2]) * bitmap.naturalWidth, (1 - crop[1] - crop[3]) * bitmap.naturalHeight, 0, 0, canvas.width, canvas.height);
    image.src = canvas.toDataURL('image/png'); image.removeAttribute('data-hwp-image-crop'); canvas.width = canvas.height = 0;
  }
  // A zero-advance heading can share its source baseline with a following
  // whitespace paragraph that reserves its height. CSS zero leading instead
  // raises the ink into the previous line. Preserve that reserved space as a
  // visible line plus after-space; retain the spacer's characters and marks.
  for (const paragraph of root.content.querySelectorAll<HTMLElement>('p[data-hwp-zero-line]')) {
    const [height, after] = paragraph.getAttribute('data-hwp-zero-line')!.split(',').map(Number);
    const next = paragraph.nextElementSibling;
    if ([height, after].every(n => Number.isFinite(n) && n >= 0 && n <= 10000) && height > 0 &&
      next?.tagName === 'P' && !next.textContent?.trim() && !next.querySelector('img,table,br,[data-page-break],[data-column-break]')) {
      paragraph.style.lineHeight = `${height}px`; paragraph.style.marginBottom = `${after}px`;
      paragraph.append(...next.childNodes); next.remove();
    }
    paragraph.removeAttribute('data-hwp-zero-line');
  }
  const images = [...root.content.querySelectorAll<HTMLImageElement>('img[data-hwp-svg],img[data-hwp-object]')];
  const warnings = new Set(source.warnings);
  let completed = 0;
  for (const image of images) {
    signal?.throwIfAborted();
    try {
      if (image.hasAttribute('data-hwp-object')) {
        image.src = await rasterizeHwpObject(image.getAttribute('data-hwp-object')!, Number(image.getAttribute('width')), Number(image.getAttribute('height')), signal);
        image.removeAttribute('data-hwp-object');
      } else {
        const raw = image.getAttribute('data-hwp-svg')!;
        if (raw.length > 4 * 1024 * 1024) throw new Error('Object preview too large');
        const clean = DOMPurify.sanitize(raw, { USE_PROFILES: { svg: true },
          FORBID_TAGS: ['style', 'script', 'foreignObject', 'use', 'a', 'animate', 'set'] });
        const xml = new DOMParser().parseFromString(clean, 'image/svg+xml');
        if (xml.documentElement.localName !== 'svg' || xml.querySelector('parsererror')) throw new Error('Invalid equation SVG');
        for (const element of xml.querySelectorAll('*')) for (const attr of [...element.attributes])
          if ((/href$/i.test(attr.name) && !(element.localName === 'image' && RASTER_DATA.test(attr.value))) || /url\s*\(/i.test(attr.value)) element.removeAttributeNode(attr);
        const width = Number(image.width), height = Number(image.height);
        if (![width, height].every(n => Number.isFinite(n) && n > 0 && n <= 10000)) throw new Error('Invalid equation size');
        const scale = Math.min(2, Math.sqrt(4_000_000 / (width * height)));
        const canvas = document.createElement('canvas');
        canvas.width = Math.max(1, Math.ceil(width * scale)); canvas.height = Math.max(1, Math.ceil(height * scale));
        const bitmap = new Image();
        bitmap.src = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(new XMLSerializer().serializeToString(xml));
        const cancel = () => { bitmap.src = ""; };
        signal?.addEventListener('abort', cancel, { once: true });
        try { await bitmap.decode(); } finally { signal?.removeEventListener('abort', cancel); }
        signal?.throwIfAborted();
        canvas.getContext('2d')!.drawImage(bitmap, 0, 0, canvas.width, canvas.height);
        image.src = canvas.toDataURL('image/png'); image.removeAttribute('data-hwp-svg');
        canvas.width = canvas.height = 0;
      }
    } catch {
      signal?.throwIfAborted();
      const object = image.hasAttribute('data-hwp-object');
      image.replaceWith(document.createTextNode(`[${image.alt || (object ? '도형' : '수식')}]`));
      warnings.add(object ? '일부 도형·글상자를 그림으로 만들지 못해 설명을 표시했습니다. 원본 보기에서 확인하세요.' : '일부 수식을 그림으로 만들지 못해 원문 명령어를 표시했습니다.');
    } finally {
      onProgress?.({ label: '수식·도형의 원래 모양을 준비하는 중…', completed: ++completed, total: images.length, unit: '개' });
    }
    if (completed % 8 === 0) await new Promise(resolve => setTimeout(resolve, 0));
  }
  return { ...source, html: root.innerHTML, warnings: [...warnings] };
}
