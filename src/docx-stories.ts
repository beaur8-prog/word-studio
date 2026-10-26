import JSZip from 'jszip';
import mammoth from 'mammoth';
import { prepareDocxMetadata } from './docx-metadata';
import { W, all, val } from './docx-format';
import type { PageSettings } from './types';

/** Convert a header, comment or note through the same formatting path as body
 * text, with its own relationships. The original ZIP and story stay unchanged. */
export async function storyHtml(zip: JSZip, part: string, root: Element, page: PageSettings) {
  const copy = new JSZip(); copy.files = { ...zip.files };
  const body = new DOMParser().parseFromString(`<w:document xmlns:w="${W}"><w:body/></w:document>`, 'application/xml');
  const target = all(body, 'body')[0];
  for (const element of Array.from(root.children)) target.append(body.importNode(element, true));
  const relationships = zip.file(part.replace(/([^/]+)$/, '_rels/$1.rels'));
  copy.file('word/_rels/document.xml.rels', relationships ? await relationships.async('string') : '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>');
  const metadata = await prepareDocxMetadata(copy, body, [page]);
  const result = await mammoth.convertToHtml({ arrayBuffer: metadata.buffer }, { externalFileAccess: false,
    includeEmbeddedStyleMap: false, ignoreEmptyParagraphs: false, styleMap: ['u => u', 'strike => s'] });
  return metadata.decorate(result.value);
}

export async function documentNotes(zip: JSZip, page: PageSettings) {
  const notes: { id: string; kind: 'footnote' | 'endnote'; html: string }[] = [];
  for (const kind of ['footnote', 'endnote'] as const) {
    const part = `word/${kind}s.xml`, file = zip.file(part); if (!file) continue;
    const parsed = new DOMParser().parseFromString(await file.async('string'), 'application/xml');
    for (const note of all(parsed, kind)) if (/^\d+$/.test(val(note, 'id')) && !['separator','continuationSeparator','continuationNotice'].includes(val(note, 'type')))
      notes.push({ id: `${kind}-${val(note, 'id')}`, kind, html: await storyHtml(zip, part, note, page) });
  }
  return notes;
}

export function attachDocumentNotes(html: string, notes: Awaited<ReturnType<typeof documentNotes>>) {
  if (!notes.length) return html;
  const template = document.createElement('template'); template.innerHTML = html;
  for (const note of notes) {
    const original = [...template.content.querySelectorAll('li[id]')].find(el => el.id === note.id);
    const list = original?.parentElement; original?.remove(); if (list && !list.children.length) list.remove();
    let label = '';
    for (const reference of template.content.querySelectorAll<HTMLAnchorElement>('a[href]')) if (reference.getAttribute('href') === `#${note.id}`) {
      label ||= reference.textContent || '';
      const node = document.createElement('span'); node.dataset.noteReference = note.id; node.dataset.noteKind = note.kind;
      node.dataset.noteLabel = label; node.textContent = label; reference.replaceWith(node);
    }
    const node = document.createElement('div'); node.dataset.documentNote = note.id; node.dataset.noteKind = note.kind;
    node.dataset.noteLabel = label; node.innerHTML = note.html; template.content.append(node);
  }
  return template.innerHTML;
}

export function structureDocxSections(html: string, pages: PageSettings[]) {
  const template = document.createElement('template'); template.innerHTML = html;
  const output = document.createElement('div');
  const differentSizes = new Set(pages.map(p => `${p.size}:${p.orientation}`)).size > 1;
  let page = pages[0], children: Element[] = [];
  const notes: Element[] = [];
  const flush = () => {
    if (!children.length) return;
    let nodes = children; children = [];
    if ((page?.columns?.count || 1) > 1) {
      // A Word page break ends every column on the current sheet. Keeping it
      // inside a CSS multicol box merely advances/extends that box and loses
      // the physical page boundary on reopen. Preserve the section wrapper,
      // but put each break between its column groups.
      const output: Element[] = [];
      let columns: HTMLDivElement | undefined;
      for (const child of nodes) {
        if (child.hasAttribute('data-page-break')) { columns = undefined; output.push(child); continue; }
        if (!columns) {
          columns = document.createElement('div'); columns.dataset.columnCount = String(page.columns!.count);
          columns.dataset.columnSpace = String(page.columns!.space); columns.dataset.columnBalance = String(page.columns!.balance !== false);
          output.push(columns);
        }
        columns.append(child);
      }
      nodes = output;
    }
    if (differentSizes && page) {
      const section = document.createElement('div'); section.dataset.sectionContent = JSON.stringify(page); section.append(...nodes); nodes = [section];
    }
    output.append(...nodes);
  };
  for (const element of [...template.content.children]) {
    if (element.hasAttribute('data-document-note')) { notes.push(element); continue; }
    if (element.hasAttribute('data-section-page')) { flush(); output.append(element); page = JSON.parse(element.getAttribute('data-section-page')!); }
    else children.push(element);
  }
  flush(); output.append(...notes); return output.innerHTML;
}
