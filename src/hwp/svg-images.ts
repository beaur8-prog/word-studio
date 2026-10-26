import DOMPurify from 'dompurify';
import type { EditableHwp } from './conversion';
import type { HwpProgressReporter } from './progress';
import { RASTER_DATA } from '../document-html';

/** Rasterize only the parser's explicit equation previews. The common image
 * importer continues to reject arbitrary SVGs and external image resources. */
export async function prepareEditableHwp(source: EditableHwp, onProgress?: HwpProgressReporter, signal?: AbortSignal): Promise<EditableHwp> {
  signal?.throwIfAborted();
  if (!source.html.includes('data-hwp-svg=')) return source;
  const root = document.createElement('template'); root.innerHTML = source.html;
  const images = [...root.content.querySelectorAll<HTMLImageElement>('img[data-hwp-svg]')];
  const warnings = new Set(source.warnings);
  let completed = 0;
  for (const image of images) {
    signal?.throwIfAborted();
    try {
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
    } catch {
      signal?.throwIfAborted();
      image.replaceWith(document.createTextNode(`[${image.alt || '수식'}]`));
      warnings.add('일부 수식을 그림으로 만들지 못해 원문 명령어를 표시했습니다.');
    }
    onProgress?.({ label: '수식·도형의 원래 모양을 준비하는 중…', completed: ++completed, total: images.length, unit: '개' });
    if (completed % 8 === 0) await new Promise(resolve => setTimeout(resolve, 0));
  }
  return { ...source, html: root.innerHTML, warnings: [...warnings] };
}
