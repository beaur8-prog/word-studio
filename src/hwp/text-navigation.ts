import type { HwpDocument } from '@rhwp/core';
import type { DocumentTabStop } from '../document-tabs';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import documentLimits from '../../electron/document-limits.json';

// CSS can contain a quoted font family inside the opposite HTML delimiter.
const attr = (tag: string, key: string) => tag.match(new RegExp(`\\b${key}=(["'])([\\s\\S]*?)\\1`))?.[2] ?? '';
const escape = (text: string) => text.replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]!);
const decode = (text: string) => text.replace(/&(?:amp|lt|gt|quot|apos|#\d+|#x[\da-f]+);/gi, e => {
  const names: Record<string, string> = { '&amp;': '&', '&lt;': '<', '&gt;': '>', '&quot;': '"', '&apos;': "'" };
  if (names[e]) return names[e];
  const n = e[2].toLowerCase() === 'x' ? parseInt(e.slice(3, -1), 16) : Number(e.slice(2, -1));
  return n > 0 && n <= 0x10ffff ? String.fromCodePoint(n) : '\ufffd';
});
const types = ['left', 'right', 'center', 'decimal'] as const;
const leaders: Record<string, DocumentTabStop['leader']> = { SOLID: 'underscore', DASH: 'hyphen', DOT: 'dot' };

/** Read normalized engine HWPX: default branches retain the doubled native
 * units, unlike HWPUNIT switch cases. Paragraph IDs also cover nested cells. */
export function hwpTabDefinitions(files: Record<string, Uint8Array>) {
  const header = files['Contents/header.xml'] ? strFromU8(files['Contents/header.xml']) : '';
  const tabs = new Map<string, { stops: DocumentTabStop[]; right: boolean }>();
  for (const m of header.matchAll(/<hh:tabPr\b[^>]*\/>|<hh:tabPr\b[^>]*>[\s\S]*?<\/hh:tabPr>/g)) {
    const tag = m[0], branch = tag.match(/<hp:default>([\s\S]*?)<\/hp:default>/)?.[1] ?? tag;
    const stops: DocumentTabStop[] = [];
    for (const item of branch.matchAll(/<hh:tabItem\b[^>]*>/g)) {
      const type = attr(item[0], 'type').toLowerCase() as DocumentTabStop['type'];
      const position = Number(attr(item[0], 'pos')) / 150;
      if (types.includes(type) && Number.isFinite(position) && position >= 0 && position <= 10000)
        stops.push({ position: Math.round(position * 100) / 100, type, ...(leaders[attr(item[0], 'leader')] ? { leader: leaders[attr(item[0], 'leader')] } : {}) });
    }
    tabs.set(attr(tag, 'id'), { stops, right: attr(tag, 'autoTabRight') === '1' });
  }
  const paragraphs = new Map<string, string>();
  for (const m of header.matchAll(/<hh:paraPr\b[^>]*>/g)) paragraphs.set(attr(m[0], 'id'), attr(m[0], 'tabPrIDRef'));
  return (paragraph: string, spacing: number, width: number): DocumentTabStop[] => {
    const source = tabs.get(paragraphs.get(paragraph) ?? ''), stops = [...(source?.stops ?? [])];
    // Plain/default tabs may also carry stored inline widths in HWP. Keep
    // their existing text flow until those widths can be mapped together.
    if (!stops.length || stops.some(s => s.position > width + .5)) return [];
    // Materialize the source's default grid as standard paragraph tab stops.
    // This survives DOCX and section-specific defaults without hidden metadata.
    const step = Number.isFinite(spacing) && spacing >= 150 ? spacing / 150 : 4000 / 75;
    const last = Math.max(0, ...stops.map(s => s.position));
    for (let i = Math.floor(last / step) + 1; i * step <= Math.max(width + step, last) && stops.length < 100; i++)
      stops.push({ position: Math.round(i * step * 100) / 100, type: 'left' });
    return stops.sort((a, b) => a.position - b.position).slice(0, 100);
  };
}
export const hwpTabAttribute = (stops?: DocumentTabStop[]) => stops?.length ? ` data-tab-stops="${escape(JSON.stringify(stops))}"` : '';
export const hwpBookmark = (name?: string) => name && name.length <= 1000 ? `<span data-docx-bookmark="${escape(name)}"></span>` : undefined;

/** Keep point bookmarks, paragraph tabs and Korean word wrapping through the
 * native HTML bridge, including nested cells. Restore real DocInfo records. */
export function stageHwpTextNavigation(html: string) {
  const prefix = `WORDSTUDIO-NAV-${crypto.randomUUID()}-`;
  const bookmarks = new Map<string, string>();
  const paragraphMarkers = new Map<string, { stops?: DocumentTabStop[]; keepWord?: boolean }>();
  const bookmarkNames = new Set<string>();
  let sequence = 0;
  html = html.replace(/<span\b[^>]*\bdata-docx-bookmark=["'][^"']*["'][^>]*>\s*<\/span>/g, tag => {
    const name = decode(attr(tag, 'data-docx-bookmark'));
    if (!name || name.length > 1000) throw Error('저장할 책갈피 이름이 올바르지 않습니다.');
    // Printed repeating headers can clone a point bookmark onto every page.
    // A named destination must still occur once in the saved document.
    if (bookmarkNames.has(name)) return '';
    bookmarkNames.add(name);
    const token = `${prefix}${sequence++}-END`; bookmarks.set(token, name);
    return `<span>${token}</span>`;
  }).replace(/<p\b[^>]*>/g, tag => {
    const raw = decode(attr(tag, 'data-tab-stops'));
    const keepWord = /(?:^|;)\s*word-break\s*:\s*keep-all\s*(?:;|$)/i.test(decode(attr(tag, 'style')));
    if (!raw && !keepWord) return tag;
    if (raw.length > 10000) throw Error('저장할 탭 정보가 너무 큽니다.');
    const stops = raw ? JSON.parse(raw) as DocumentTabStop[] : [];
    if (!Array.isArray(stops) || stops.length > 100 || stops.some(s => !s || !types.includes(s.type) || !Number.isFinite(s.position) || Math.abs(s.position) > 10000))
      throw Error('저장할 탭 정보가 올바르지 않습니다.');
    const supported = stops.filter(s => s.position >= 0);
    if (!supported.length && !keepWord) return tag;
    const token = `${prefix}${sequence++}-END`;
    paragraphMarkers.set(token, { ...(supported.length ? { stops: supported } : {}), ...(keepWord ? { keepWord: true } : {}) });
    return tag + `<span>${token}</span>`;
  });
  return { html, restore(Core: typeof HwpDocument, doc: HwpDocument) {
    if (!sequence) return doc;
    let bytes = 0, found = 0;
    const files = unzipSync(doc.exportHwpx(), { filter: f => {
      if ((bytes += f.originalSize) > documentLimits.documentInflatedBytes) throw Error('한글 책갈피·탭 변환 데이터가 너무 큽니다.');
      return true;
    } });
    let header = strFromU8(files['Contents/header.xml']);
    const paraShapes = new Map([...header.matchAll(/<hh:paraPr\b[^>]*>[\s\S]*?<\/hh:paraPr>/g)].map(m => [attr(m[0], 'id'), m[0]]));
    let nextPara = Math.max(-1, ...[...paraShapes.keys()].map(Number)) + 1;
    let nextTab = Math.max(-1, ...[...header.matchAll(/<hh:tabPr\b[^>]*>/g)].map(m => Number(attr(m[0], 'id')))) + 1;
    const newParas: string[] = [], newTabs: string[] = [], reused = new Map<string, number>();
    for (const name of Object.keys(files).filter(n => /^Contents\/section\d+\.xml$/.test(n))) {
      const stack: { ref: string; replacement?: number }[] = [];
      let xml = strFromU8(files[name]).replace(/<hp:p\b[^>]*>|<\/hp:p>|<hp:t\b[^>]*>[\s\S]*?<\/hp:t>/g, tag => {
        if (tag.startsWith('<hp:p')) { stack.push({ ref: attr(tag, 'paraPrIDRef') }); return tag; }
        if (tag === '</hp:p>') { const p = stack.pop(); return p?.replacement === undefined ? tag : `<!--${prefix}PARA-${p.replacement}-->` + tag; }
        return tag.replace(new RegExp(`${prefix}\\d+-END`, 'g'), token => {
          const bookmark = bookmarks.get(token);
          if (bookmark !== undefined) { found++; return `</hp:t><hp:ctrl><hp:bookmark name="${escape(bookmark)}"/></hp:ctrl><hp:t>`; }
          const change = paragraphMarkers.get(token), paragraph = stack.at(-1);
          if (!change || !paragraph) throw Error('한글 문단·책갈피·탭 위치가 올바르지 않습니다.');
          found++;
          const key = paragraph.ref + ':' + JSON.stringify(change);
          let id = reused.get(key);
          if (id === undefined) {
            const base = paraShapes.get(paragraph.ref); if (!base) throw Error('한글 문단 서식을 찾지 못했습니다.');
            id = nextPara++; reused.set(key, id);
            let shape = base.replace(/\bid="[^"]*"/, `id="${id}"`);
            if (change.stops) {
              const tabId = nextTab++;
              newTabs.push(`<hh:tabPr id="${tabId}" autoTabLeft="0" autoTabRight="0">` + change.stops.map(s => `<hh:tabItem pos="${Math.round(s.position * 150)}" type="${s.type.toUpperCase()}" leader="${s.leader === 'dot' ? 'DOT' : s.leader === 'hyphen' ? 'DASH' : s.leader === 'underscore' ? 'SOLID' : 'NONE'}"/>`).join('') + '</hh:tabPr>');
              shape = shape.replace(/\btabPrIDRef="[^"]*"/, `tabPrIDRef="${tabId}"`);
            }
            if (change.keepWord) {
              // Native 0 / HWPX non-Latin BREAK_WORD keeps an eojeol whole.
              // The native HTML paste otherwise resets this setting to 1.
              shape = shape.replace(/<hh:breakSetting\b[^>]*>/, tag => tag.replace(/\bbreakNonLatinWord="[^"]*"/, 'breakNonLatinWord="BREAK_WORD"'));
            }
            newParas.push(shape);
          }
          paragraph.replacement = id; return '';
        });
      });
      // Apply each paragraph's new shape after processing its children. The
      // marker belongs to the closest unmatched P, so nested cells stay local.
      const open: number[] = [], edits: { from: number; to: number; value: string }[] = [];
      for (const m of xml.matchAll(new RegExp(`<hp:p\\b[^>]*>|<\\/hp:p>|<!--${prefix}PARA-(\\d+)-->`, 'g'))) {
        if (m[0].startsWith('<hp:p')) open.push(m.index!);
        else if (m[0] === '</hp:p>') open.pop();
        else { const from = open.at(-1); if (from === undefined) throw Error('한글 탭 문단 구조가 올바르지 않습니다.'); const end = xml.indexOf('>', from) + 1; edits.push({ from, to: end, value: xml.slice(from, end).replace(/\bparaPrIDRef="[^"]*"/, `paraPrIDRef="${m[1]}"`) }, { from: m.index!, to: m.index! + m[0].length, value: '' }); }
      }
      for (const e of edits.sort((a, b) => b.from - a.from)) xml = xml.slice(0, e.from) + e.value + xml.slice(e.to);
      files[name] = strToU8(xml.replace(/<hp:linesegarray\b[^>]*>[\s\S]*?<\/hp:linesegarray>/g, ''));
    }
    if (found !== sequence) throw Error('일부 책갈피·탭을 한글로 저장하지 못했습니다.');
    for (const [tag, added] of [['tabProperties', newTabs], ['paraProperties', newParas]] as const) if (added.length)
      header = header.replace(new RegExp(`(<hh:${tag}\\b[^>]*itemCnt=")(\\d+)("[^>]*>)([\\s\\S]*?)(<\\/hh:${tag}>)`), (_, a, count, b, contents, end) => a + (Number(count) + added.length) + b + contents + added.join('') + end);
    files['Contents/header.xml'] = strToU8(header);
    return new Core(zipSync(files, { level: 1 }));
  } };
}
