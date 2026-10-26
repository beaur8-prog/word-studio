import type { HwpDocument } from '@rhwp/core';
import { strFromU8, strToU8, zipSync, unzipSync } from 'fflate';
import documentLimits from '../../electron/document-limits.json';
import { entities } from './cell-content';

type Range = { start: number; end: number; open: string; inner: string };
/** Match complete elements, including paragraphs containing nested tables.
 * Inputs are sanitized HTML or the engine's bounded, normalized HWPX. */
function ranges(source: string, name: string): Range[] {
  const result: Range[] = [], stack: { start: number; inner: number; open: string }[] = [];
  for (const m of source.matchAll(new RegExp(`<!--[\\s\\S]*?-->|<!\\[CDATA\\[[\\s\\S]*?\\]\\]>|<\\/?${name}\\b[^>]*>`, 'g'))) {
    const tag = m[0]; if (tag.startsWith('<!')) continue;
    if (tag.startsWith('</')) {
      const start = stack.pop();
      if (start && !stack.length) result.push({ start: start.start, end: m.index! + tag.length, open: start.open, inner: source.slice(start.inner, m.index) });
    } else if (!tag.endsWith('/>')) stack.push({ start: m.index!, inner: m.index! + tag.length, open: tag });
  }
  return result;
}
const attr = (tag: string, name: string) => entities(tag.match(new RegExp(`\\b${name}=["']([^"']*)["']`))?.[1] ?? '');
const plain = (xml: string) => entities(xml.replace(/<[^>]*>/g, ''));
const paragraph = (text: string) => `<hp:p id="0" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0"><hp:run charPrIDRef="0"><hp:t>${text}</hp:t></hp:run></hp:p>`;
const withoutLines = (xml: string) => xml.replace(/<hp:linesegarray\b[^>]*>[\s\S]*?<\/hp:linesegarray>/g, '');
function archive(doc: HwpDocument) {
  let size = 0;
  return unzipSync(doc.exportHwpx(), { filter: entry => {
    if ((size += entry.originalSize) > documentLimits.documentInflatedBytes) throw Error('각주·미주 변환 데이터가 너무 큽니다.');
    return true;
  } });
}

/** Stage the note stories in the same paste as the body, so their character,
 * paragraph, table and image IDs belong to one native document. Move them into
 * real note controls afterwards, rather than leaving printed notes in body. */
export function stageHwpNotes(html: string) {
  const prefix = `WORDSTUDIO-NOTE-${crypto.randomUUID()}-`;
  const notes = new Map<string, { index: number; kind: 'footNote' | 'endNote'; number: number; html: string; refs: number; xml?: string }>();
  let cursor = 0, staged = '';
  for (const block of ranges(html, 'div')) {
    const id = attr(block.open, 'data-document-note'); if (!id) continue;
    if (notes.has(id)) throw Error('각주·미주 ID가 중복되어 저장할 수 없습니다.');
    const label = Number(attr(block.open, 'data-note-label'));
    const kind = attr(block.open, 'data-note-kind') === 'endnote' ? 'endNote' : 'footNote';
    const number = Number.isInteger(label) && label > 0 && label <= 65535 ? label : [...notes.values()].filter(n => n.kind === kind).length + 1;
    notes.set(id, { index: notes.size, kind, number, html: block.inner, refs: 0 });
    staged += html.slice(cursor, block.start); cursor = block.end;
  }
  staged += html.slice(cursor);
  staged = staged.replace(/<span\b[^>]*\bdata-note-reference=["'][^"']+["'][^>]*>[^<]*<\/span>/g, tag => {
    const note = notes.get(attr(tag, 'data-note-reference'));
    if (!note) throw Error('각주·미주 본문을 찾지 못해 저장을 중단했습니다.');
    if (++note.refs > 1) throw Error('같은 각주·미주를 여러 번 참조하는 문서는 한글로 저장할 수 없습니다. DOCX 또는 백업으로 저장하세요.');
    return tag.replace(/>[^<]*<\/span>$/, `>${prefix}${note.index}-REF</span>`);
  });
  for (const note of notes.values()) {
    if (!note.refs) throw Error('본문 참조가 없는 각주·미주가 있습니다. DOCX 또는 백업으로 저장하세요.');
    staged += `<p>${prefix}${note.index}-START</p>${note.html}<p>${prefix}${note.index}-END</p>`;
  }
  return { html: staged, restore(Core: typeof HwpDocument, doc: HwpDocument) {
    if (!notes.size) return doc;
    const files = archive(doc), section = 'Contents/section0.xml';
    let xml = strFromU8(files[section]), from = -1, active: typeof notes extends Map<string, infer N> ? N | undefined : never;
    const removals: { start: number; end: number }[] = [];
    for (const p of ranges(xml, 'hp:p')) {
      const text = plain(p.inner);
      const marker = [...notes.values()].find(n => text === `${prefix}${n.index}-START`);
      if (marker) { if (active) throw Error('각주 변환 범위가 올바르지 않습니다.'); active = marker; from = p.end; removals.push({ start: p.start, end: -1 }); }
      else if (active && text === `${prefix}${active.index}-END`) {
        active.xml = withoutLines(xml.slice(from, p.start));
        removals.at(-1)!.end = p.end; active = undefined;
      }
    }
    if (active || [...notes.values()].some(n => !n.xml)) throw Error('각주·미주 본문을 한글 구조로 변환하지 못했습니다.');
    for (const r of removals.reverse()) xml = xml.slice(0, r.start) + xml.slice(r.end);
    let restored = 0;
    xml = xml.replace(/<hp:t\b[^>]*>[\s\S]*?<\/hp:t>/g, tag => tag.replace(new RegExp(prefix + '\\d+-REF', 'g'), token => {
      const note = [...notes.values()].find(n => token === `${prefix}${n.index}-REF`)!;
      const type = note.kind === 'footNote' ? 'FOOTNOTE' : 'ENDNOTE';
      const auto = `<hp:ctrl><hp:autoNum num="${note.number}" numType="${type}"><hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar=")" supscript="0"/></hp:autoNum></hp:ctrl>`;
      const body = note.xml!.replace(/(<hp:run\b[^>]*>)/, '$1' + auto);
      restored++;
      return `</hp:t><hp:ctrl><hp:${note.kind} number="${note.number}" suffixChar="41" instId="${note.index}"><hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="TOP" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">${body}</hp:subList></hp:${note.kind}></hp:ctrl><hp:t>`;
    }));
    if (restored !== notes.size) throw Error('일부 각주·미주 참조를 저장하지 못했습니다.');
    files[section] = strToU8(withoutLines(xml));
    return new Core(zipSync(files, { level: 1 }));
  } };
}

/** Render note stories through the same native HTML conversion as ordinary
 * body content. This isolated document shares styles and binaries with source,
 * while never editing the user's document or treating a screenshot as text. */
export function importHwpNotes(doc: HwpDocument, files: Record<string, Uint8Array>, render: (doc: HwpDocument) => { html: string; warnings: string[] }) {
  const notes: { id: string; kind: 'footnote' | 'endnote'; label: string; xml: string; html?: string }[] = [];
  for (const [name, bytes] of Object.entries(files)) {
    const section = name.match(/^Contents\/section(\d+)\.xml$/)?.[1]; if (section === undefined) continue;
    const xml = strFromU8(bytes);
    for (const r of ranges(xml, 'hp:(?:footNote|endNote)')) {
      const kind = r.open.startsWith('<hp:footNote') ? 'footnote' : 'endnote';
      const id = `hwp-note-${section}-${notes.filter(n => n.id.startsWith(`hwp-note-${section}-`)).length}`;
      const inner = r.inner.match(/<hp:subList\b[^>]*>([\s\S]*)<\/hp:subList>/)?.[1];
      if (!inner) throw Error('각주·미주 본문 구조가 올바르지 않습니다.');
      notes.push({ id, kind, label: attr(r.open, 'number') || '1', xml: withoutLines(inner)
        .replace(/<hp:ctrl>\s*<hp:autoNum\b[^>]*numType="(?:FOOTNOTE|ENDNOTE)"[^>]*>[\s\S]*?<\/hp:autoNum>\s*<\/hp:ctrl>/g, '') });
    }
  }
  if (!notes.length) return { notes, warnings: [] as string[] };
  const prefix = 'WORDSTUDIO-NOTE-IMPORT-' + crypto.randomUUID() + '-';
  const copy = archive(doc);
  for (const name of Object.keys(copy).filter(n => /^Contents\/section\d+\.xml$/.test(n))) {
    const xml = strFromU8(copy[name]);
    const contents = name === 'Contents/section0.xml' ? notes.map((n, i) => paragraph(`${prefix}${i}-START`) + n.xml + paragraph(`${prefix}${i}-END`)).join('') : paragraph('');
    copy[name] = strToU8(xml.replace(/(<hs:sec\b[^>]*>)[\s\S]*?(<\/hs:sec>)/, '$1' + contents + '$2'));
  }
  const Core = doc.constructor as typeof HwpDocument;
  const isolated = new Core(zipSync(copy, { level: 1 }));
  try {
    const result = render(isolated);
    let active = -1, from = 0;
    for (const p of ranges(result.html, 'p')) {
      const text = plain(p.inner).trim();
      const start = notes.findIndex((_, i) => text === `${prefix}${i}-START`);
      if (start >= 0) { active = start; from = p.end; }
      else if (active >= 0 && text === `${prefix}${active}-END`) { notes[active].html = result.html.slice(from, p.start).trim() || '<p></p>'; active = -1; }
    }
    if (notes.some(n => n.html === undefined)) throw Error('일부 각주·미주 본문을 변환하지 못했습니다.');
    return { notes, warnings: result.warnings };
  } finally { isolated.free(); }
}
