import { documentScale } from './document-scale';
import { Extension } from '@tiptap/core';
import { Plugin, PluginKey } from '@tiptap/pm/state';
import { Decoration, DecorationSet, type EditorView } from '@tiptap/pm/view';
import type { Node as PMNode } from '@tiptap/pm/model';
import { imageFloating } from './image-floating';
import { paginateCellRow, paginateMergedRows } from './editor-cell-pagination';
import { layoutColumns } from './column-layout';
import { layoutTableAnchors } from './table-anchor';
import { paginateParagraph } from './editor-paragraph-pagination';

const key = new PluginKey('studioPagination');
type RepeatingHeader = { table: HTMLTableElement; rows: HTMLTableRowElement[]; height: number };
type Piece = { pos: number; end: number; top: number; height: number; columns?: number; columnFlow?: boolean; cellRow?: HTMLTableRowElement; mergedRows?: HTMLTableRowElement[]; explicit?: 'pageBreak' | 'sectionBreak'; breakType?: string; empty?: boolean; references?: { id: string; offset: number }[]; header?: RepeatingHeader; tableTop?: number };

/** Use real block/row geometry. A form is often one large table cell containing
 * paragraphs and nested tables, so treating only top-level nodes as indivisible
 * blocks lets text run straight through page margins. */
function pieces(view: EditorView, usableHeight: number, scale: number, removed: { top: number; height: number }[], root: PMNode, start: number, domRoot: HTMLElement) {
  const origin = domRoot.getBoundingClientRect().top;
  const result: Piece[] = [];
  const rect = (element: HTMLElement) => {
    const box = element.getBoundingClientRect();
    const top = (box.top - origin) / scale, bottom = (box.bottom - origin) / scale;
    return { top: top - removed.filter(g => g.top < top - .5).reduce((n, g) => n + g.height, 0),
      height: (box.height / scale) - removed.filter(g => g.top >= top - .5 && g.top < bottom - .5).reduce((n, g) => n + g.height, 0) };
  };
  const walk = (node: PMNode, pos: number) => {
    const dom = view.nodeDOM(pos) as HTMLElement | null;
    if (!dom?.getBoundingClientRect) return;
    const box = rect(dom);
    if (node.type.name === 'columnGroup' && node.attrs.count === 1) { node.forEach((child, offset) => walk(child, pos + 1 + offset)); return; }
    if (node.type.name === 'columnGroup') {
      const prefix = parseFloat(dom.style.getPropertyValue('--column-prefix')) || 0;
      box.top += prefix; box.height = Math.max(0, box.height - prefix);
    }
    // Absolute floating pictures still occupy paper. Include their extent so
    // a short anchor near the bottom moves with its picture to the next page.
    for (const img of dom.querySelectorAll<HTMLElement>('img[data-image-floating]')) {
      if (imageFloating(img.dataset.imageFloating)?.wrap?.type === 0) continue;
      const picture = rect(img); box.height = Math.max(box.height, picture.top + picture.height - box.top);
    }
    if (node.type.name === 'documentNote' && node.attrs.kind === 'footnote') return;
    if (node.type.name === 'pageBreak' || node.type.name === 'sectionBreak') {
      let breakType = '';
      if (node.type.name === 'sectionBreak') { try { breakType = JSON.parse(node.attrs.page).breakType || ''; } catch { /* Validated on file input. */ } }
      result.push({ pos, end: pos + node.nodeSize, ...box, height: 0, explicit: node.type.name, breakType }); return;
    }
    if (node.type.name !== 'table') { result.push({ pos, end: pos + node.nodeSize, ...box, columnFlow: node.type.name === 'columnGroup', empty: !node.textContent.trim() && !dom.querySelector('img,[data-note-reference]') }); return; }
    const rows: { node: PMNode; pos: number; dom: HTMLElement }[] = [];
    node.forEach((row, offset) => {
      const rowPos = pos + 1 + offset, rowDom = view.nodeDOM(rowPos) as HTMLElement;
      if (rowDom) rows.push({ node: row, pos: rowPos, dom: rowDom });
    });
    let headerEnd = 0;
    while (headerEnd < rows.length && rows[headerEnd].node.childCount > 0 && rows[headerEnd].node.content.content.every(cell => cell.type.name === 'tableHeader')) headerEnd++;
    const header: RepeatingHeader | undefined = headerEnd && headerEnd < rows.length ? {
      table: rows[0].dom.closest('table')!, rows: rows.slice(0, headerEnd).map(r => r.dom as HTMLTableRowElement),
      height: rect(rows[headerEnd - 1].dom).top + rect(rows[headerEnd - 1].dom).height - rect(rows[0].dom).top,
    } : undefined;
    for (let i = 0; i < rows.length;) {
      let end = i === 0 && headerEnd ? Math.min(rows.length, headerEnd + 1) : i + 1;
      for (let j = i; j < end; j++) rows[j].node.forEach(cell => { end = Math.min(rows.length, Math.max(end, j + Number(cell.attrs.rowspan || 1))); });
      const first = rect(rows[i].dom), last = rect(rows[end - 1].dom);
      const height = last.top + last.height - first.top;
      const row = rows[i];
      if (i === 0 && headerEnd && height > usableHeight && rows[headerEnd]?.node.attrs.rowSplit === 'allow') {
        const headers = rect(rows[headerEnd - 1].dom);
        const paragraph = rows[headerEnd].dom.querySelector('p');
        const line = paragraph ? parseFloat(getComputedStyle(paragraph).lineHeight) || 24 : 24;
        let columns = 0; row.node.forEach(cell => { columns += Number(cell.attrs.colspan || 1); });
        result.push({ pos: row.pos, end: rows[headerEnd - 1].pos + rows[headerEnd - 1].node.nodeSize, top: first.top, height: headers.top + headers.height - first.top + line, columns });
        i = headerEnd; continue;
      }
      let containsTable = false;
      if (row.node.childCount === 1) row.node.firstChild!.forEach(block => { if (block.type.name === 'table') containsTable = true; });
      if ((height > usableHeight || containsTable) && end === i + 1 && row.node.childCount === 1) {
        const cell = row.node.firstChild!;
        cell.forEach((block, offset) => walk(block, row.pos + 2 + offset));
      } else {
        let columns = 0; row.node.forEach(cell => { columns += Number(cell.attrs.colspan || 1); });
        const cellRow = end === i + 1 && row.node.attrs.rowSplit === 'allow' && row.node.attrs.wordRowRule !== 'exact' && row.node.content.content.every(c => Number(c.attrs.rowspan || 1) === 1)
          ? row.dom as HTMLTableRowElement : undefined;
        result.push({ pos: row.pos, end: rows[end - 1].pos + rows[end - 1].node.nodeSize, top: first.top, height, columns, cellRow, header: i >= headerEnd ? header : undefined, tableTop: rect(rows[0].dom).top, mergedRows: end > i + 1 && height > usableHeight ? rows.slice(i, end).map(r => r.dom as HTMLTableRowElement) : undefined });
      }
      i = end;
    }
  };
  root.forEach((node, offset) => walk(node, start + offset));
  for (const piece of result) {
    const dom = view.nodeDOM(piece.pos) as HTMLElement;
    piece.references = [...dom.querySelectorAll<HTMLElement>('[data-note-reference]')].map(ref => ({ id: ref.dataset.noteReference!, offset: rect(ref).top - piece.top }));
  }
  return result;
}

export const Pagination = Extension.create({
  name: 'pagination',
  addProseMirrorPlugins() {
    return [new Plugin({
      key,
      state: { init: () => DecorationSet.empty, apply: (tr, old) => tr.getMeta(key) ?? old.map(tr.mapping, tr.doc) },
      props: { decorations: state => key.getState(state) },
      view(view) {
        let scheduled = 0, signature = '', lastDocument: PMNode | undefined;
        const overlays = document.createElement('div'); overlays.className = 'pagination-headers'; overlays.setAttribute('aria-hidden', 'true');
        const calculate = () => {
          cancelAnimationFrame(scheduled);
          scheduled = requestAnimationFrame(() => {
            if (!view.dom.isConnected || view.composing || view.dom.classList.contains('table-row-resizing')) return;
            const sheet = view.dom.closest<HTMLElement>('.document-stack'); if (!sheet) return;
            // Column prefixes depend on the previous page gaps. Resolve them
            // before measuring; min-height can keep the root's ResizeObserver
            // silent even when column fragments have moved inside the sheet.
            layoutTableAnchors(view.dom); layoutColumns(view.dom);
            view.dom.classList.add('pagination-measuring-cells');
            try {
            const dec: Decoration[] = [], signatures: string[] = [];
            const repeated: { header: RepeatingHeader; top: number; left: number; width: number }[] = [];
            const contexts: { node: PMNode; start: number; dom: HTMLElement; sheet: HTMLElement }[] = [];
            view.state.doc.forEach((node, pos) => {
              if (node.type.name === 'sectionContent') {
                const element = view.nodeDOM(pos) as HTMLElement, dom = element?.querySelector<HTMLElement>('.section-editable');
                if (dom) contexts.push({ node, start: pos + 1, dom, sheet: element });
              }
            });
            if (contexts.length) view.state.doc.forEach((node, pos) => {
              if (node.type.name === 'sectionBreak') dec.push(Decoration.node(pos, pos + node.nodeSize, { style: 'height:24px;margin:0' }));
            });
            else contexts.push({ node: view.state.doc, start: 0, dom: view.dom, sheet });
            let totalCount = 0;
            const placedNotes = new Set<string>();
            for (const context of contexts) {
              const h = Number(context.sheet.dataset.pageHeight) || Number(sheet.dataset.pageHeight) || 1122.52;
              const top = Number(context.sheet.dataset.margin ?? sheet.dataset.margin ?? 96);
              const bottom = Number(context.sheet.dataset.bottomMargin ?? sheet.dataset.bottomMargin ?? top), stride = h + 24, usable = h - top - bottom;
              const origin = context.dom.getBoundingClientRect().top, scale = documentScale(context.dom);
              const removed = [...context.dom.querySelectorAll<HTMLElement>('.pagination-gap,.page-break,.section-break')].map(el => {
                const r = el.getBoundingClientRect(); return { top: (r.top - origin) / scale, height: r.height / scale };
              });
              const noteNodes = new Map<string, { pos: number; end: number; height: number }>();
              view.state.doc.descendants((node, position) => {
                if (node.type.name !== 'documentNote' || node.attrs.kind !== 'footnote') return;
                const pos = position, dom = view.nodeDOM(pos) as HTMLElement;
                if (dom) {
                  const css = getComputedStyle(dom);
                  const height = dom.getBoundingClientRect().height / scale - ['paddingTop', 'paddingBottom', 'borderTopWidth', 'borderBottomWidth']
                    .reduce((sum, property) => sum + (parseFloat(css[property as 'paddingTop']) || 0), 0);
                  noteNodes.set(String(node.attrs.id), { pos, end: pos + node.nodeSize, height });
                }
                return false;
              });
              const assigned = new Map<string, number>(), onPage = new Map<number, Set<string>>();
              const assignNotes = (refs: { id: string; page: number }[]) => {
                for (const ref of refs) if (!assigned.has(ref.id)) {
                  assigned.set(ref.id, ref.page); const ids = onPage.get(ref.page) || new Set<string>(); ids.add(ref.id); onPage.set(ref.page, ids);
                }
              };
              const reserve = (page: number, candidate: { id: string; page: number }[]) => {
                const ids = new Set([...(onPage.get(page) || []), ...candidate.filter(ref => ref.page === page).map(ref => ref.id)]);
                // Match the printed note group: 12px separation, 1px rule,
                // 6px padding and a 4px gap before each note.
                return ids.size ? 19 + [...ids].reduce((sum, id) => sum + 4 + (noteNodes.get(id)?.height || 0), 0) : 0;
              };
              let extra = 0, sectionJustStarted = false;
              const tablePages = new Map<HTMLTableElement, number>(), headerPages = new Map<HTMLTableElement, Set<number>>();
              const repeat = (header: RepeatingHeader | undefined, page: number) => {
                if (!header || page <= (tablePages.get(header.table) ?? page)) return 0;
                const shown = headerPages.get(header.table) ?? new Set<number>();
                if (!shown.has(page)) {
                  shown.add(page); headerPages.set(header.table, shown);
                  const box = header.table.getBoundingClientRect(), sheetBox = sheet.getBoundingClientRect();
                  repeated.push({ header, top: (origin - sheetBox.top) / scale + page * stride + top,
                    left: (box.left - sheetBox.left) / scale, width: box.width / scale });
                  signatures.push(`header:${view.posAtDOM(header.table, 0)}:${page}:${header.height}`);
                }
                return header.height;
              };
              for (const piece of pieces(view, usable, scale, removed, context.node, context.start, context.dom)) {
                if (piece.explicit === 'sectionBreak' && piece.breakType === 'continuous') {
                  dec.push(Decoration.node(piece.pos, piece.end, { style: 'height:0;margin:0' })); signatures.push(`continuous:${piece.pos}`); continue;
                }
                if (piece.explicit === 'pageBreak' && sectionJustStarted) {
                  dec.push(Decoration.node(piece.pos, piece.end, { style: 'height:0;margin:0' })); signatures.push(`coalesced:${piece.pos}`); sectionJustStarted = false; continue;
                }
                const y = piece.top + extra, page = Math.max(0, Math.floor((y - top + .5) / stride));
                if (piece.header && !tablePages.has(piece.header.table)) tablePages.set(piece.header.table,
                  Math.max(0, Math.floor(((piece.tableTop ?? piece.top) + extra - top + .5) / stride)));
                const trialPages = new Set<number>();
                const repeatAt = (page: number) => {
                  if (!piece.header || page <= (tablePages.get(piece.header.table) ?? page)) return 0;
                  trialPages.add(page); return piece.header.height;
                };
                const commitHeaders = () => trialPages.forEach(page => repeat(piece.header, page));
                const candidate = () => (piece.references || []).filter(ref => noteNodes.has(ref.id) && !assigned.has(ref.id) && !placedNotes.has(ref.id))
                  .map(ref => ({ id: ref.id, page: Math.max(0, Math.floor((piece.top + extra + ref.offset - top + .5) / stride)) }));
                let refs = candidate();
                const end = page * stride + h - bottom - reserve(page, refs);
                if (piece.mergedRows) {
                  const flow = paginateMergedRows(view, piece.mergedRows, { origin, scale, top, bottom, height: h, extra, removed, repeatAt, reserveAt: page => reserve(page, refs) });
                  if (flow) { commitHeaders(); dec.push(...flow.decorations); signatures.push(...flow.signatures); extra += flow.extra; sectionJustStarted = false; continue; }
                }
                if (piece.cellRow && y + piece.height > end + .5) {
                  const flow = paginateCellRow(view, piece.cellRow, piece.pos, { origin, scale, top, bottom, height: h, extra, removed, repeatAt,
                    references: (piece.references ?? []).filter(ref => refs.some(r => r.id === ref.id)).map(ref => ({ id: ref.id, top: piece.top + ref.offset })),
                    reserveAt: (page, own) => reserve(page, own ?? []) });
                  if (flow) { assignNotes(flow.references); commitHeaders(); dec.push(...flow.decorations); signatures.push(...flow.signatures); extra += flow.extra; sectionJustStarted = false; continue; }
                }
                if (!piece.columns && !piece.explicit && y + piece.height > end + .5) {
                  const paragraph = view.nodeDOM(piece.pos);
                  if (paragraph instanceof HTMLElement && paragraph.tagName === 'P') {
                    const flow = paginateParagraph(view, paragraph, { origin, scale, top, bottom, height: h, extra, removed,
                      references: (piece.references ?? []).filter(ref => refs.some(r => r.id === ref.id)).map(ref => ({ id: ref.id, top: piece.top + ref.offset })),
                      reserveAt: (page, own) => reserve(page, own) });
                    if (flow) { assignNotes(flow.references); dec.push(...flow.decorations); signatures.push(...flow.signatures); extra += flow.extra; sectionJustStarted = false; continue; }
                  }
                }
                // Multi-column groups already fragment across paper rows in
                // layoutColumns. Moving their outer box as an atomic block
                // adds a second page gap and feeds it back into column-prefix.
                if (piece.explicit || (!piece.columnFlow && y + piece.height > end + .5 && piece.height <= usable)) {
                  let destination = page + 1;
                  if (piece.breakType === 'oddPage' || piece.breakType === 'evenPage') {
                    const parity = piece.breakType === 'oddPage' ? 1 : 0;
                    if ((totalCount + destination + 1) % 2 !== parity) destination++;
                  }
                  const gap = Math.max(0, Math.ceil((destination * stride + top + repeat(piece.header, destination) - y) * 64) / 64);
                  if (piece.explicit) dec.push(Decoration.node(piece.pos, piece.end, { style: `height:${gap}px` }));
                  else dec.push(Decoration.widget(piece.pos, () => {
                    const el = document.createElement(piece.columns ? 'tr' : 'div');
                    el.className = `pagination-gap${piece.columns ? ' pagination-row-gap' : ''}`;
                    el.style.height = `${gap}px`; el.contentEditable = 'false'; el.setAttribute('aria-hidden', 'true');
                    if (piece.columns) { const cell = document.createElement('td'); cell.colSpan = piece.columns; cell.style.height = `${gap}px`; el.append(cell); }
                    return el;
                  }, { side: -1, key: `gap-${piece.pos}-${gap}` }));
                  extra += gap; signatures.push(`${piece.pos}:${gap.toFixed(3)}`); refs = candidate();
                }
                assignNotes(refs);
                if (piece.explicit === 'sectionBreak') sectionJustStarted = true;
                else if (!piece.empty && !piece.explicit) sectionJustStarted = false;
              }
              const padding = getComputedStyle(context.dom), left = parseFloat(padding.paddingLeft) || 0;
              const width = context.dom.clientWidth - left - (parseFloat(padding.paddingRight) || 0);
              const globalTop = (context.dom.getBoundingClientRect().top - view.dom.getBoundingClientRect().top) / scale;
              const globalLeft = (context.dom.getBoundingClientRect().left - view.dom.getBoundingClientRect().left) / scale;
              for (const [page, ids] of onPage) {
                let y = page * stride + h - bottom - reserve(page, []) + 12, first = true;
                for (const id of ids) {
                  const note = noteNodes.get(id)!;
                  if (!first) y += 4;
                  dec.push(Decoration.node(note.pos, note.end, { class: 'pagination-footnote',
                    style: `position:absolute;top:${globalTop + y}px;left:${globalLeft + left}px;width:${width}px;margin:0;border:${first ? '1px solid #77818b' : '0'};border-width:${first ? '1px 0 0' : '0'};padding:${first ? '10px 0 0' : '0'}` }));
                  placedNotes.add(id);
                  y += note.height + (first ? 11 : 0); first = false; signatures.push(`note:${id}:${page}:${note.height}`);
                }
              }
              const last = [...context.dom.children].filter(el => !el.matches('[data-document-note][data-note-kind="footnote"]')).at(-1) as HTMLElement | undefined;
              const naturalBottom = last ? (last.getBoundingClientRect().bottom - origin) / scale - removed.reduce((n, g) => n + g.height, 0) : top;
              const count = Math.max(1, Math.ceil((naturalBottom + extra + bottom - .5) / stride), ...[...onPage.keys()].map(page => page + 1));
              if (context.sheet !== sheet) context.sheet.dataset.pages = String(count);
              totalCount += count; signatures.push(`${context.start}:${h}:${top}:${bottom}:${count}`);
            }
            sheet.dataset.pages = String(totalCount); sheet.dispatchEvent(new CustomEvent('pagecount', { bubbles: true, detail: totalCount }));
            if (!overlays.isConnected) sheet.append(overlays);
            overlays.replaceChildren(...repeated.map(item => {
              const box = document.createElement('div'); box.className = 'tiptap';
              Object.assign(box.style, { position: 'absolute', top: `${item.top}px`, left: `${item.left}px`, width: `${item.width}px`, minHeight: '0', padding: '0' });
              const table = item.header.table.cloneNode(false) as HTMLTableElement;
              table.style.width = `${item.width}px`; table.style.margin = '0';
              const columns = item.header.table.querySelector(':scope > colgroup'); if (columns) table.append(columns.cloneNode(true));
              const body = document.createElement('tbody'); body.append(...item.header.rows.map(row => row.cloneNode(true))); table.append(body); box.append(table);
              return box;
            }));
            const next = signatures.join('|');
            if (signature !== next || lastDocument !== view.state.doc) {
              signature = next; lastDocument = view.state.doc;
              view.dispatch(view.state.tr.setMeta(key, DecorationSet.create(view.state.doc, dec)).setMeta('addToHistory', false));
            }
            } finally { view.dom.classList.remove('pagination-measuring-cells'); }
          });
        };
        calculate(); const observer = new ResizeObserver(calculate); observer.observe(view.dom);
        document.fonts.addEventListener('loadingdone', calculate); view.dom.addEventListener('compositionend', calculate);
        return { update: calculate, destroy() { overlays.remove(); cancelAnimationFrame(scheduled); observer.disconnect(); document.fonts.removeEventListener('loadingdone', calculate); view.dom.removeEventListener('compositionend', calculate); } };
      },
    })];
  },
});
