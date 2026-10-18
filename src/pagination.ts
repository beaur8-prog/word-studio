import { Extension } from '@tiptap/core';
import { Plugin, PluginKey } from '@tiptap/pm/state';
import { Decoration, DecorationSet, type EditorView } from '@tiptap/pm/view';
import type { Node as PMNode } from '@tiptap/pm/model';

const key = new PluginKey('studioPagination');
type Piece = { pos: number; end: number; top: number; height: number; columns?: number; explicit?: 'pageBreak' | 'sectionBreak'; breakType?: string; empty?: boolean; references?: { id: string; offset: number }[] };

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
    if (node.type.name === 'documentNote' && node.attrs.kind === 'footnote') return;
    if (node.type.name === 'pageBreak' || node.type.name === 'sectionBreak') {
      let breakType = '';
      if (node.type.name === 'sectionBreak') { try { breakType = JSON.parse(node.attrs.page).breakType || ''; } catch { /* Validated on file input. */ } }
      result.push({ pos, end: pos + node.nodeSize, ...box, height: 0, explicit: node.type.name, breakType }); return;
    }
    if (node.type.name !== 'table') { result.push({ pos, end: pos + node.nodeSize, ...box, empty: !node.textContent.trim() && !dom.querySelector('img,[data-note-reference]') }); return; }
    const rows: { node: PMNode; pos: number; dom: HTMLElement }[] = [];
    node.forEach((row, offset) => {
      const rowPos = pos + 1 + offset, rowDom = view.nodeDOM(rowPos) as HTMLElement;
      if (rowDom) rows.push({ node: row, pos: rowPos, dom: rowDom });
    });
    for (let i = 0; i < rows.length;) {
      let end = i + 1;
      for (let j = i; j < end; j++) rows[j].node.forEach(cell => { end = Math.min(rows.length, Math.max(end, j + Number(cell.attrs.rowspan || 1))); });
      const first = rect(rows[i].dom), last = rect(rows[end - 1].dom);
      const height = last.top + last.height - first.top;
      const row = rows[i];
      let containsTable = false;
      if (row.node.childCount === 1) row.node.firstChild!.forEach(block => { if (block.type.name === 'table') containsTable = true; });
      if ((height > usableHeight || containsTable) && end === i + 1 && row.node.childCount === 1) {
        const cell = row.node.firstChild!;
        cell.forEach((block, offset) => walk(block, row.pos + 2 + offset));
      } else {
        let columns = 0; row.node.forEach(cell => { columns += Number(cell.attrs.colspan || 1); });
        result.push({ pos: row.pos, end: rows[end - 1].pos + rows[end - 1].node.nodeSize, top: first.top, height, columns });
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
        const calculate = () => {
          cancelAnimationFrame(scheduled);
          scheduled = requestAnimationFrame(() => {
            if (!view.dom.isConnected || view.composing || view.dom.classList.contains('table-row-resizing')) return;
            const sheet = view.dom.closest<HTMLElement>('.document-stack'); if (!sheet) return;
            const dec: Decoration[] = [], signatures: string[] = [];
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
              const origin = context.dom.getBoundingClientRect().top, scale = context.dom.getBoundingClientRect().width / context.dom.offsetWidth || 1;
              const removed = [...context.dom.querySelectorAll<HTMLElement>('.pagination-gap,.page-break,.section-break')].map(el => {
                const r = el.getBoundingClientRect(); return { top: (r.top - origin) / scale, height: r.height / scale };
              });
              const noteNodes = new Map<string, { pos: number; end: number; height: number }>();
              view.state.doc.descendants((node, position) => {
                if (node.type.name !== 'documentNote' || node.attrs.kind !== 'footnote') return;
                const pos = position, dom = view.nodeDOM(pos) as HTMLElement;
                if (dom) noteNodes.set(String(node.attrs.id), { pos, end: pos + node.nodeSize, height: dom.offsetHeight + 4 });
                return false;
              });
              const assigned = new Map<string, number>(), onPage = new Map<number, Set<string>>();
              const reserve = (page: number, candidate: { id: string; page: number }[]) => {
                const ids = new Set([...(onPage.get(page) || []), ...candidate.filter(ref => ref.page === page).map(ref => ref.id)]);
                return ids.size ? 12 + [...ids].reduce((sum, id) => sum + (noteNodes.get(id)?.height || 0), 0) : 0;
              };
              let extra = 0, sectionJustStarted = false;
              for (const piece of pieces(view, usable, scale, removed, context.node, context.start, context.dom)) {
                if (piece.explicit === 'sectionBreak' && piece.breakType === 'continuous') {
                  dec.push(Decoration.node(piece.pos, piece.end, { style: 'height:0;margin:0' })); signatures.push(`continuous:${piece.pos}`); continue;
                }
                if (piece.explicit === 'pageBreak' && sectionJustStarted) {
                  dec.push(Decoration.node(piece.pos, piece.end, { style: 'height:0;margin:0' })); signatures.push(`coalesced:${piece.pos}`); sectionJustStarted = false; continue;
                }
                const y = piece.top + extra, page = Math.floor((y - top + .5) / stride);
                const candidate = () => (piece.references || []).filter(ref => noteNodes.has(ref.id) && !assigned.has(ref.id) && !placedNotes.has(ref.id))
                  .map(ref => ({ id: ref.id, page: Math.max(0, Math.floor((piece.top + extra + ref.offset - top + .5) / stride)) }));
                let refs = candidate();
                const end = page * stride + h - bottom - reserve(page, refs);
                if (piece.explicit || (y + piece.height > end + .5 && piece.height <= usable)) {
                  let destination = page + 1;
                  if (piece.breakType === 'oddPage' || piece.breakType === 'evenPage') {
                    const parity = piece.breakType === 'oddPage' ? 1 : 0;
                    if ((totalCount + destination + 1) % 2 !== parity) destination++;
                  }
                  const gap = Math.max(0, destination * stride + top - y);
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
                for (const ref of refs) if (!assigned.has(ref.id)) {
                  assigned.set(ref.id, ref.page); const ids = onPage.get(ref.page) || new Set(); ids.add(ref.id); onPage.set(ref.page, ids);
                }
                if (piece.explicit === 'sectionBreak') sectionJustStarted = true;
                else if (!piece.empty && !piece.explicit) sectionJustStarted = false;
              }
              const padding = getComputedStyle(context.dom), left = parseFloat(padding.paddingLeft) || 0;
              const width = context.dom.clientWidth - left - (parseFloat(padding.paddingRight) || 0);
              const globalTop = (context.dom.getBoundingClientRect().top - view.dom.getBoundingClientRect().top) / scale;
              const globalLeft = (context.dom.getBoundingClientRect().left - view.dom.getBoundingClientRect().left) / scale;
              for (const [page, ids] of onPage) {
                let y = page * stride + h - bottom - reserve(page, []);
                for (const id of ids) {
                  const note = noteNodes.get(id)!;
                  dec.push(Decoration.node(note.pos, note.end, { style: `position:absolute;top:${globalTop + y}px;left:${globalLeft + left}px;width:${width}px;margin:0` }));
                  placedNotes.add(id);
                  y += note.height; signatures.push(`note:${id}:${page}:${note.height}`);
                }
              }
              const last = [...context.dom.children].filter(el => !el.matches('[data-document-note][data-note-kind="footnote"]')).at(-1) as HTMLElement | undefined;
              const naturalBottom = last ? (last.getBoundingClientRect().bottom - origin) / scale - removed.reduce((n, g) => n + g.height, 0) : top;
              const count = Math.max(1, Math.ceil((naturalBottom + extra + bottom - .5) / stride), ...[...onPage.keys()].map(page => page + 1));
              if (context.sheet !== sheet) context.sheet.dataset.pages = String(count);
              totalCount += count; signatures.push(`${context.start}:${h}:${top}:${bottom}:${count}`);
            }
            sheet.dataset.pages = String(totalCount); sheet.dispatchEvent(new CustomEvent('pagecount', { bubbles: true, detail: totalCount }));
            const next = signatures.join('|');
            if (signature !== next || lastDocument !== view.state.doc) {
              signature = next; lastDocument = view.state.doc;
              view.dispatch(view.state.tr.setMeta(key, DecorationSet.create(view.state.doc, dec)).setMeta('addToHistory', false));
            }
          });
        };
        calculate(); const observer = new ResizeObserver(calculate); observer.observe(view.dom);
        document.fonts.addEventListener('loadingdone', calculate); view.dom.addEventListener('compositionend', calculate);
        return { update: calculate, destroy() { cancelAnimationFrame(scheduled); observer.disconnect(); document.fonts.removeEventListener('loadingdone', calculate); view.dom.removeEventListener('compositionend', calculate); } };
      },
    })];
  },
});
