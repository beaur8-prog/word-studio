import type { HwpDocument } from '@rhwp/core';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import documentLimits from '../../electron/document-limits.json';

/** The HTML bridge flattens fields and the native field insertion API handles
 * only body text. Restore fields in the engine's own HWPX before either final
 * writer, so nested cells retain the same live control and character style. */
export function stageHwpFields(html: string) {
  const prefix = `WORDSTUDIO-FIELD-${crypto.randomUUID()}-`;
  const fields = new Map<string, 'PAGE' | 'TOTAL_PAGE'>();
  const staged = html.replace(/<span\b[^>]*\bdata-page-number=["'](PAGE|NUMPAGES)["'][^>]*>[^<]*<\/span>/g, (tag, kind: string) => {
    const token = prefix + fields.size + '-END'; fields.set(token, kind === 'PAGE' ? 'PAGE' : 'TOTAL_PAGE');
    return tag.replace(/>[^<]*<\/span>$/, '>' + token + '</span>');
  });
  return { html: staged, restore(Core: typeof HwpDocument, doc: HwpDocument): HwpDocument {
    if (!fields.size) return doc;
    let size = 0, restored = 0;
    const files = unzipSync(doc.exportHwpx(), { filter: file => {
      if ((size += file.originalSize) > documentLimits.documentInflatedBytes) throw Error('한글 필드 변환 데이터가 너무 큽니다.');
      return true;
    } });
    for (const name of Object.keys(files)) {
      if (!/^Contents\/section\d+\.xml$/.test(name)) continue;
      const xml = strFromU8(files[name]).replace(/<hp:t\b[^>]*>([\s\S]*?)<\/hp:t>/g, (tag: string) => tag.replace(new RegExp(prefix + '\\d+-END', 'g'), token => {
        const type = fields.get(token); if (!type) throw Error('저장할 한글 필드 정보가 올바르지 않습니다.');
        restored++;
        return `</hp:t><hp:ctrl><hp:autoNum num="0" numType="${type}"><hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar="" supscript="0"/></hp:autoNum></hp:ctrl><hp:t>`;
      })).replace(/<hp:linesegarray\b[^>]*>[\s\S]*?<\/hp:linesegarray>/g, '');
      files[name] = strToU8(xml);
    }
    if (restored !== fields.size) throw Error('일부 한글 자동 쪽 번호를 저장하지 못했습니다.');
    return new Core(zipSync(files, { level: 1 }));
  } };
}
