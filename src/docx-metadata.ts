import type JSZip from 'jszip';
import type { Comment, PageSettings } from './types';
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const all = (root: Element | Document, tag: string) => Array.from(root.getElementsByTagNameNS(W, tag));
const val = (el: Element | undefined, key = 'val') => el?.getAttributeNS(W, key) ?? '';
const parse = (text: string) => new DOMParser().parseFromString(text, 'application/xml');
const child = (root: Element | undefined, name: string) => root && Array.from(root.children).find(el => el.localName === name);
const hex = (value: string) => /^[\da-f]{6}$/i.test(value) ? `#${value}` : '';
function ancestor(element: Element, name: string): Element | null {
  let parent = element.parentElement;
  while (parent && parent.localName !== name) parent = parent.parentElement;
  return parent;
}

function decorateCell(cell: HTMLElement, source: Element) {
  const paragraphProps = child(all(source, 'p')[0], 'pPr');
  if (paragraphProps) {
    const defaults = document.createElement('span').style;
    defaults.cssText = runStyle(paragraphProps);
    for (const property of ['font-family', 'font-size', 'color']) {
      if (defaults.getPropertyValue(property)) cell.style.setProperty(property, defaults.getPropertyValue(property));
    }
  }
  const props = child(source, 'tcPr');
  if (!props) return;
  const fill = hex(val(child(props, 'shd'), 'fill'));
  if (fill) cell.style.backgroundColor = fill;
  else if (val(child(props, 'shd'), 'fill') === 'auto') cell.style.backgroundColor = 'transparent';
  const vertical = val(child(props, 'vAlign'));
  if (['top', 'center', 'bottom'].includes(vertical)) cell.style.verticalAlign = vertical === 'center' ? 'middle' : vertical;
  const borders = child(props, 'tcBorders'), margins = child(props, 'tcMar');
  for (const side of ['top', 'right', 'bottom', 'left']) {
    const border = child(borders, side) ?? child(borders, side === 'left' ? 'start' : side === 'right' ? 'end' : side);
    const kind = val(border);
    if (border) {
      const styles: Record<string, string> = { single: 'solid', dashed: 'dashed', dashSmallGap: 'dashed', dotted: 'dotted', double: 'double', nil: 'none', none: 'none' };
      const width = Number(val(border, 'sz')) / 8 || .75;
      cell.style.setProperty(`border-${side}`, `${styles[kind] === 'none' ? 0 : width}pt ${styles[kind] ?? 'solid'} ${hex(val(border, 'color')) || '#000000'}`);
    }
    const margin = child(margins, side) ?? child(margins, side === 'left' ? 'start' : side === 'right' ? 'end' : side);
    if (margin && val(margin, 'type') === 'dxa') cell.style.setProperty(`padding-${side}`, `${Number(val(margin, 'w')) / 15}px`);
  }
}

function runStyle(run: Element): string {
  const props = child(run, 'rPr');
  if (!props) return '';
  const style = document.createElement('span').style;
  const color = hex(val(child(props, 'color')));
  if (color) style.color = color;
  const font = child(props, 'rFonts');
  const family = val(font, 'eastAsia') || val(font, 'ascii') || val(font, 'hAnsi');
  if (family) style.fontFamily = family;
  const size = Number(val(child(props, 'sz')));
  if (size > 0) style.fontSize = `${size / 2}pt`;
  const highlight = val(child(props, 'highlight'));
  if (highlight && highlight !== 'none') style.backgroundColor = highlight;
  const shading = hex(val(child(props, 'shd'), 'fill'));
  if (shading) style.backgroundColor = shading;
  return style.cssText;
}

/** Temporary bookmarks let Mammoth preserve exact positions, including repeated text. */
export async function prepareDocxMetadata(zip: JSZip, body: Document, pages: PageSettings[]) {
  const comments: Comment[] = [];
  const idMap = new Map<string, string>();
  const commentFile = zip.file('word/comments.xml');
  const resolved = new Set<string>();
  const extended = zip.file('word/commentsExtended.xml');
  if (extended) for (const el of Array.from(parse(await extended.async('string')).getElementsByTagName('*'))) {
    const attrs = Array.from(el.attributes);
    if (attrs.some(a => a.localName === 'done' && a.value === '1')) {
      const para = attrs.find(a => a.localName === 'paraId')?.value;
      if (para) resolved.add(para);
    }
  }
  if (commentFile) for (const el of all(parse(await commentFile.async('string')), 'comment')) {
    const id = crypto.randomUUID();
    idMap.set(val(el, 'id'), id);
    const date = val(el, 'date');
    const lastParagraph = all(el, 'p').at(-1);
    const paraId = lastParagraph && Array.from(lastParagraph.attributes).find(a => a.localName === 'paraId')?.value;
    comments.push({ id, quote: '', text: all(el, 'p').map(p => all(p, 't').map(t => t.textContent).join('')).join('\n'),
      createdAt: Number.isFinite(Date.parse(date)) ? new Date(date).toISOString() : new Date().toISOString(),
      resolved: Boolean(paraId && resolved.has(paraId)) });
  }
  let sequence = 900000;
  const markers = new Map<string, { kind: string; id?: string; paragraph?: Element; cell?: Element; style?: string; nextPage?: PageSettings }>();
  const bookmark = (name: string) => {
    const fragment = body.createDocumentFragment();
    const id = String(++sequence);
    const start = body.createElementNS(W, 'w:bookmarkStart');
    start.setAttributeNS(W, 'w:id', id); start.setAttributeNS(W, 'w:name', name);
    const end = body.createElementNS(W, 'w:bookmarkEnd'); end.setAttributeNS(W, 'w:id', id);
    fragment.append(start, end);
    return fragment;
  };
  for (const kind of ['commentRangeStart', 'commentRangeEnd']) for (const el of all(body, kind)) {
    const id = idMap.get(val(el, 'id'));
    if (!id) continue;
    const name = `ws_marker_${sequence}`;
    markers.set(name, { kind: kind === 'commentRangeStart' ? 'start' : 'end', id });
    el.replaceWith(bookmark(name));
  }
  for (const run of all(body, 'r')) {
    const style = runStyle(run);
    if (!style) continue;
    const start = `ws_marker_${sequence}`;
    markers.set(start, { kind: 'runStart', style }); run.before(bookmark(start));
    const end = `ws_marker_${sequence}`;
    markers.set(end, { kind: 'runEnd' }); run.after(bookmark(end));
  }
  let section = 0;
  for (const paragraph of all(body, 'p')) {
    const name = `ws_marker_${sequence}`;
    const isEnd = all(paragraph, 'sectPr').length > 0;
    const parent = paragraph.parentElement;
    const cell = parent?.localName === 'tc' && all(parent, 'p')[0] === paragraph ? parent : undefined;
    markers.set(name, { kind: 'paragraph', paragraph, cell,
      ...(isEnd ? { nextPage: pages[++section] } : {}) });
    const properties = Array.from(paragraph.children).find(e => e.localName === 'pPr');
    paragraph.insertBefore(bookmark(name), properties ? properties.nextSibling : paragraph.firstChild);
  }
  zip.file('word/document.xml', new XMLSerializer().serializeToString(body));
  return {
    buffer: await zip.generateAsync({ type: 'arraybuffer' }), comments,
    decorate(html: string) {
      const template = document.createElement('template'); template.innerHTML = html;
      const sizedTables = new Set<HTMLTableElement>();
      // Attach paragraph properties before deleting marker anchors.
      for (const anchor of template.content.querySelectorAll('a[id]')) {
        const marker = markers.get(anchor.id);
        if (marker?.kind !== 'paragraph') continue;
        const p = anchor.closest<HTMLElement>('p,h1,h2,h3,h4,h5,h6,li');
        if (!p || !marker.paragraph) continue;
        const cell = p.closest<HTMLElement>('td,th');
        if (cell && marker.cell) {
          decorateCell(cell, marker.cell);
          const sourceTable = ancestor(marker.cell, 'tbl'), table = cell.closest('table');
          if (sourceTable && table && !sizedTables.has(table)) {
            sizedTables.add(table);
            const rows = all(sourceTable, 'tr').filter(row => ancestor(row, 'tbl') === sourceTable);
            // Vertical merges can leave an HTML row without any paragraph
            // anchors. Table row order still matches the source, including these rows.
            if (rows.length === table.rows.length) rows.forEach((source, index) => {
              const property = child(child(source, 'trPr'), 'trHeight'), height = Number(val(property));
              if (height > 0 && val(property, 'hRule') !== 'auto') table.rows[index].style.height = `${height / 15}px`;
            });
          }
        }
        const props = all(marker.paragraph, 'pPr')[0];
        if (props) {
          const align = val(all(props, 'jc')[0]);
          if (['left', 'right', 'center', 'both'].includes(align)) p.style.textAlign = align === 'both' ? 'justify' : align;
          const indent = Number(val(all(props, 'ind')[0], 'left'));
          if (indent > 0) p.style.marginLeft = `${indent / 15}px`;
          const spacing = all(props, 'spacing')[0];
          const line = Number(val(spacing, 'line'));
          if (line > 0 && (!val(spacing, 'lineRule') || val(spacing, 'lineRule') === 'auto')) p.style.lineHeight = String(line / 240);
        }
        if (marker.nextPage) {
          const sectionBreak = document.createElement('div');
          sectionBreak.setAttribute('data-section-page', JSON.stringify(marker.nextPage));
          p.after(sectionBreak);
        }
      }
      const active = new Set<string>();
      let activeStyle = '';
      const walker = document.createTreeWalker(template.content, NodeFilter.SHOW_ELEMENT | NodeFilter.SHOW_TEXT);
      const items: Node[] = []; while (walker.nextNode()) items.push(walker.currentNode);
      for (const item of items) {
        if (item instanceof HTMLAnchorElement && markers.has(item.id)) {
          const marker = markers.get(item.id)!;
          if (marker.kind === 'start') active.add(marker.id!);
          if (marker.kind === 'end') active.delete(marker.id!);
          if (marker.kind === 'runStart') activeStyle = marker.style || '';
          if (marker.kind === 'runEnd') activeStyle = '';
          item.remove();
        } else if (item.nodeType === Node.TEXT_NODE && (active.size || activeStyle) && item.textContent) {
          const span = document.createElement('span');
          if (active.size) span.setAttribute('data-comment-ids', [...active].join(' '));
          if (activeStyle) span.style.cssText = activeStyle;
          item.parentNode!.replaceChild(span, item); span.append(item);
          for (const id of active) { const comment = comments.find(c => c.id === id); if (comment) comment.quote += item.textContent; }
        }
      }
      return template.innerHTML;
    },
  };
}
