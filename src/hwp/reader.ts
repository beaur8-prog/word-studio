import DOMPurify from 'dompurify';

export interface HwpInfo { count: number; fonts: string[]; version: string; warnings: number }
export interface HwpPage { svg: string; text: string; width: number; height: number }

/** Only inert SVG paint from the local parser reaches the document. */
export function safeHwpSvg(source: string): Pick<HwpPage, 'svg' | 'width' | 'height'> {
  const sanitized = DOMPurify.sanitize(source, { USE_PROFILES: { svg: true, svgFilters: true },
    FORBID_TAGS: ['style', 'script', 'foreignObject', 'a', 'animate', 'animateMotion', 'animateTransform', 'set'], ADD_ATTR: ['xml:space'] });
  const doc = new DOMParser().parseFromString(sanitized, 'image/svg+xml');
  const svg = doc.documentElement;
  if (svg.localName !== 'svg' || doc.querySelector('parsererror')) throw new Error('HWP 쪽을 표시하지 못했습니다.');
  // SVG fragment IDs are document-global, including across different inline
  // SVG roots. Multiple printed pages must not reuse another page's cell clip.
  const prefix = `hwp-${crypto.randomUUID()}-`, ids = new Map<string, string>();
  for (const el of [svg, ...svg.querySelectorAll('*')]) {
    const id = el.getAttribute('id');
    if (id) { const next = prefix + ids.size; ids.set(id, next); el.setAttribute('id', next); }
  }
  for (const el of [svg, ...svg.querySelectorAll('*')]) {
    for (const attribute of [...el.attributes]) {
      const value = attribute.value.trim();
      if (/^(?:href|xlink:href)$/i.test(attribute.name)) {
        if (value.startsWith('#') && ids.has(value.slice(1))) attribute.value = '#' + ids.get(value.slice(1));
        else if (!/^data:image\/(?:png|jpeg|gif|webp|bmp);base64,[a-z0-9+/=\s]+$/i.test(value)) el.removeAttributeNode(attribute);
      } else if (/url\s*\(/i.test(value)) {
        const reference = value.match(/^url\(\s*['"]?#([\w.-]+)['"]?\s*\)$/)?.[1];
        if (reference && ids.has(reference)) attribute.value = `url(#${ids.get(reference)})`; else el.removeAttributeNode(attribute);
      }
    }
  }
  const width = Number(svg.getAttribute('width')), height = Number(svg.getAttribute('height'));
  if (![width, height].every(value => Number.isFinite(value) && value >= 50 && value <= 5000)) throw new Error('지원하지 않는 HWP 용지 크기입니다.');
  return { svg: new XMLSerializer().serializeToString(svg), width, height };
}

export class HwpReaderSession {
  private worker = new Worker(new URL('./reader-worker.ts', import.meta.url), { type: 'module' });
  private next = 0;
  private closed = false;
  private pending = new Map<number, { resolve(value: unknown): void; reject(error: Error): void; timer: ReturnType<typeof setTimeout> }>();
  constructor() {
    this.worker.onmessage = ({ data }) => {
      const request = this.pending.get(data.id);
      if (!request) return;
      clearTimeout(request.timer); this.pending.delete(data.id);
      if (data.error) request.reject(new Error(data.error)); else request.resolve(data.result);
    };
    this.worker.onerror = () => this.close('HWP 해석 중 오류가 발생했습니다. 다른 문서를 열어 주세요.');
  }
  private call<T>(action: string, payload: object = {}, transfer: Transferable[] = []): Promise<T> {
    if (this.closed) return Promise.reject(new Error('읽기 작업이 취소됐습니다.'));
    const id = ++this.next;
    return new Promise<T>((resolve, reject) => {
      const timer = setTimeout(() => this.close('HWP 읽기 시간이 초과됐습니다. 더 작은 문서로 나눠 주세요.'), 45000);
      this.pending.set(id, { resolve: value => resolve(value as T), reject, timer });
      this.worker.postMessage({ id, action, ...payload }, transfer);
    });
  }
  async open(file: File): Promise<HwpInfo> {
    if (!file.size || file.size > 40 * 1024 * 1024) throw new Error('40MB 이하의 HWP/HWPX 문서를 선택하세요.');
    const bytes = await file.arrayBuffer();
    return this.call('open', { bytes, name: file.name, extension: file.name.split('.').pop()?.toLowerCase() }, [bytes]);
  }
  async page(index: number): Promise<HwpPage> {
    const page = await this.call<{ svg: string; text: string }>('page', { page: index });
    return { ...page, ...safeHwpSvg(page.svg) };
  }
  text(): Promise<string> { return this.call('text'); }
  close(message = '읽기 작업이 취소됐습니다.') {
    if (this.closed) return;
    this.closed = true; this.worker.terminate();
    for (const request of this.pending.values()) { clearTimeout(request.timer); request.reject(new Error(message)); }
    this.pending.clear();
  }
}
