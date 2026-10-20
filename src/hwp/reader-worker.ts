import init, { HwpDocument } from '@rhwp/core';
import wasmUrl from '@rhwp/core/rhwp_bg.wasm?url';
import JSZip from 'jszip';

let document: HwpDocument | undefined;
const ready = init({ module_or_path: fetch(wasmUrl).then(response => {
  if (!response.ok) throw new Error('HWP 읽기 모듈을 불러오지 못했습니다.');
  return response.arrayBuffer();
}) });
const scope = self as unknown as { onmessage: ((event: MessageEvent) => void) | null; postMessage(value: unknown): void };

scope.onmessage = async ({ data }) => {
  const { id, action } = data;
  try {
    await ready;
    let result: unknown;
    if (action === 'open') {
      const bytes = new Uint8Array(data.bytes);
      if (!bytes.length || bytes.length > 40 * 1024 * 1024) throw new Error('40MB 이하의 HWP/HWPX 문서를 선택하세요.');
      if (data.extension === 'hwpx') {
        const zip = await JSZip.loadAsync(bytes);
        const entries = Object.values(zip.files);
        const size = entries.reduce((sum, entry) => sum + ((entry as unknown as { _data?: { uncompressedSize?: number } })._data?.uncompressedSize ?? 0), 0);
        if (entries.length > 10000 || size > 100 * 1024 * 1024) throw new Error('압축을 푼 HWPX가 100MB 또는 내부 파일 10,000개를 넘습니다.');
        if (!zip.file('Contents/content.hpf') || !zip.file('Contents/header.xml')) throw new Error('유효한 HWPX 문서가 아닙니다.');
      } else if (![0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1].every((value, i) => bytes[i] === value)) {
        throw new Error('HWP 5.x 문서가 아닙니다. 구형 HWP 3.x는 한글에서 HWP 5.x 또는 HWPX로 변환하세요.');
      }
      document?.free();
      document = new HwpDocument(bytes);
      document.setFileName(data.name);
      const info = JSON.parse(document.getDocumentInfo());
      const count = document.pageCount();
      if (!count || count > 1000) throw new Error('1,000쪽 이하의 문서를 선택하세요.');
      result = { count, fonts: info.fontsUsed ?? [], version: info.version, warnings: JSON.parse(document.getValidationWarnings()).count ?? 0 };
    } else {
      if (!document) throw new Error('문서를 먼저 열어 주세요.');
      if (action === 'page') {
        const page = data.page;
        if (!Number.isInteger(page) || page < 0 || page >= document.pageCount()) throw new Error('쪽 번호가 올바르지 않습니다.');
        const svg = document.renderPageSvg(page);
        if (svg.length > 12 * 1024 * 1024) throw new Error('이 쪽의 그림·도형 데이터가 너무 큽니다.');
        const layout = JSON.parse(document.getPageTextLayout(page));
        result = { svg, text: (layout.runs ?? []).map((run: { text: string }) => run.text).join('\n') };
      } else if (action === 'text') {
        // The upstream source-text API omits text nested in some HWP table and
        // textbox controls. Copy the same layout runs that the reader paints.
        const pages: string[] = [];
        let size = 0;
        for (let i = 0; i < document.pageCount(); i++) {
          const runs: { text: string; x: number; y: number; w: number; h: number }[] = JSON.parse(document.getPageTextLayout(i)).runs ?? [];
          let text = '', previous: typeof runs[number] | undefined;
          for (const run of runs) {
            if (previous && Math.abs(run.y - previous.y) > Math.max(2, run.h * .35)) text += '\n';
            else if (previous && run.x - previous.x - previous.w > Math.max(4, run.h * .6)) text += '\t';
            text += run.text; previous = run;
          }
          size += new TextEncoder().encode(text).byteLength + 2;
          if (size > 4 * 1024 * 1024) throw new Error('복사할 문서 텍스트가 너무 큽니다. 필요한 쪽에서 직접 선택해 복사하세요.');
          pages.push(text);
        }
        result = pages.join('\n\n');
      }
      else throw new Error('지원하지 않는 읽기 작업입니다.');
    }
    scope.postMessage({ id, result });
  } catch (error) {
    const message = String(error instanceof Error ? error.message : error);
    scope.postMessage({ id, error: /password|비밀번호|암호화|encrypted/i.test(message)
      ? '암호화된 문서는 읽을 수 없습니다. 한글에서 암호를 해제한 사본을 여세요.' : message });
  }
};
