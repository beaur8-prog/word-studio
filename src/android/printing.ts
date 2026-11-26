import { generateHTML } from '@tiptap/core';
import { editorExtensions } from '../extensions';
import { createPrintLayout, type PrintLayoutProgress } from '../print-layout';
import { withoutEditorReflow } from '../editor-layout-task';
import type { StudioDocument } from '../types';
import { androidServices, type AndroidPrintDocument, type AndroidPrintResult } from './platform';

export const ANDROID_PRINT_HELP = '시스템 인쇄창에서 프린터 또는 PDF로 저장을 선택하세요. 용지 크기와 방향도 그 창에서 설정합니다.';
const LIMIT = 128 * 1024 * 1024;
const escape = (value: string) => value.replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]!);
export function androidPrintMessage(result: AndroidPrintResult) {
  return result === 'cancelled' ? '인쇄를 취소했습니다.' : result === 'submitted' ? '시스템에 인쇄 요청을 전달했습니다.' : '인쇄창을 닫았습니다.';
}
/** A new print snapshot, independent of viewer virtualization and display zoom. */
export async function prepareAndroidPrintDocument(doc: StudioDocument, signal?: AbortSignal, progress?: (value: PrintLayoutProgress) => void) {
  const layout = await withoutEditorReflow(() => createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page, progress, signal));
  try { return await prepareAndroidPrintPages([...layout.element.querySelectorAll<HTMLElement>(':scope > .print-page')], doc.title, signal, progress); }
  finally { layout.dispose(); }
}
export async function prepareAndroidPrintPages(pages: readonly HTMLElement[], title: string, signal?: AbortSignal, progress?: (value: PrintLayoutProgress) => void): Promise<AndroidPrintDocument> {
  if (!pages.length) throw new Error('인쇄할 쪽이 없습니다.');
  const root = document.createElement('main'); root.className = 'print-layout android-print-layout';
  const embedded = new Map<string, Promise<string>>();
  const resource = (url: string) => {
    if (/^data:/i.test(url)) return Promise.resolve(url);
    if (!/^blob:/i.test(url) && new URL(url, document.baseURI).origin !== location.origin) throw new Error('외부 그림을 인쇄할 수 없습니다. 그림을 문서에 포함한 뒤 다시 시도하세요.');
    let result = embedded.get(url);
    if (!result) { result = fetch(new URL(url, document.baseURI), { signal }).then(response => {
      if (!response.ok) throw new Error('인쇄할 그림을 읽지 못했습니다.'); return response.blob();
    }).then(blob => new Promise<string>((resolve, reject) => { const reader = new FileReader(); reader.onload = () => resolve(String(reader.result)); reader.onerror = () => reject(new Error('인쇄할 그림을 읽지 못했습니다.')); reader.readAsDataURL(blob); })); embedded.set(url, result); }
    return result;
  };
  let first = { width: 0, height: 0 };
  for (let index = 0; index < pages.length; index++) {
    signal?.throwIfAborted(); const source = pages[index];
    const width = parseFloat(source.style.width) || source.offsetWidth, height = parseFloat(source.style.height) || source.offsetHeight;
    if (!(width > 0 && height > 0 && width <= 4608 && height <= 4608)) throw new Error('인쇄할 용지 크기를 확인할 수 없습니다.');
    if (!index) first = { width, height };
    const clone = source.cloneNode(true) as HTMLElement;
    clone.querySelectorAll('script,iframe,object,embed,link,meta,base,audio,video,source').forEach(node => node.remove());
    for (const form of clone.querySelectorAll('form')) form.replaceWith(...form.childNodes);
    for (const node of [clone, ...clone.querySelectorAll('*')]) for (const attr of [...node.attributes]) if (/^on/i.test(attr.name) || ['srcdoc', 'autofocus', 'contenteditable', 'tabindex', 'formaction', 'action'].includes(attr.name)) node.removeAttributeNode(attr);
    const originals = [...source.querySelectorAll('canvas')];
    for (const [n, canvas] of [...clone.querySelectorAll('canvas')].entries()) { const image = document.createElement('img'); image.src = originals[n].toDataURL('image/png'); image.style.cssText = canvas.style.cssText; image.width = originals[n].width; image.height = originals[n].height; canvas.replaceWith(image); }
    for (const image of clone.querySelectorAll<HTMLImageElement>('img')) { image.removeAttribute('srcset'); image.loading = 'eager'; if (image.getAttribute('src')) image.src = await resource(image.getAttribute('src')!); }
    for (const image of clone.querySelectorAll('svg image')) for (const attr of ['href', 'xlink:href']) { const url = image.getAttribute(attr); if (url && !url.startsWith('#')) image.setAttribute(attr, await resource(url)); }
    for (const node of clone.querySelectorAll<HTMLElement>('[style]')) {
      let style = node.getAttribute('style')!;
      for (const match of [...style.matchAll(/url\(\s*['"]?(blob:[^)'"\s]+)['"]?\s*\)/g)]) style = style.replace(match[0], `url("${await resource(match[1])}")`);
      node.setAttribute('style', style);
    }
    clone.classList.add('android-print-source'); Object.assign(clone.style, { position: 'absolute', top: '0', left: '0', margin: '0', transform: 'none', transformOrigin: 'top left', page: 'auto' });
    const sheet = document.createElement('section'); sheet.className = 'android-print-sheet'; sheet.dataset.sourceWidth = String(width); sheet.dataset.sourceHeight = String(height); sheet.append(clone); root.append(sheet);
    progress?.({ label: '인쇄할 쪽을 준비하는 중…', completed: index + 1, total: pages.length });
    if (index % 8 === 0) await new Promise(resolve => setTimeout(resolve, 0));
  }
  signal?.throwIfAborted();
  const styles = [...document.head.querySelectorAll<HTMLStyleElement | HTMLLinkElement>('style,link[rel="stylesheet"]')].filter(node => !(node instanceof HTMLLinkElement) || new URL(node.href).origin === location.origin).map(node => node.outerHTML).join('\n');
  const css = `@page{size:auto!important;margin:0!important}html,body{margin:0!important;padding:0!important;background:white!important;color:#242932!important;overflow:visible!important}body .android-print-layout{position:static!important;margin:0!important;padding:0!important;visibility:visible!important;overflow:visible!important}.android-print-sheet{position:relative;overflow:hidden;box-sizing:border-box;break-after:page;page-break-after:always}.android-print-sheet:last-child{break-after:auto;page-break-after:auto}.android-print-source{position:absolute!important;margin:0!important;box-shadow:none!important;break-after:auto!important;page-break-after:auto!important;page:auto!important;transform-origin:top left!important;-webkit-print-color-adjust:exact;print-color-adjust:exact}.android-print-source *{-webkit-print-color-adjust:exact;print-color-adjust:exact}`;
  const html = `<!doctype html><html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><base href="${escape(document.baseURI)}"><meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'unsafe-inline' ${escape(location.origin)}; font-src data: ${escape(location.origin)}; img-src data: ${escape(location.origin)}; script-src 'none'; base-uri ${escape(location.origin)}; form-action 'none'"><title>${escape(title)}</title>${styles}<style>${css}</style></head><body class="studio-printing">${root.outerHTML}</body></html>`;
  const data = new TextEncoder().encode(html); if (data.byteLength > LIMIT) throw new Error('인쇄 데이터가 128 MiB를 넘습니다. 문서를 나눠 인쇄하세요.');
  return { name: `${title.replace(/[<>:"/\\|?*]/g, '-') || '문서'}.html`, type: 'text/html', data, pageWidthMils: Math.round(first.width / 96 * 1000), pageHeightMils: Math.round(first.height / 96 * 1000), pages: pages.length };
}
export async function printAndroidPrepared(document: AndroidPrintDocument): Promise<AndroidPrintResult> {
  const services = androidServices();
  if (!services?.printDocument) throw new Error('이 설치본에는 Android 인쇄 연결이 없습니다. 업데이트한 APK에서 다시 시도하세요.');
  const result = await services.printDocument(document);
  if (!['submitted', 'cancelled', 'closed'].includes(result)) throw new Error('인쇄창의 결과를 확인하지 못했습니다.');
  return result;
}
