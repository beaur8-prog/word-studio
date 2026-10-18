import { Extension } from '@tiptap/core';
import { Plugin, PluginKey } from '@tiptap/pm/state';
import { Decoration, DecorationSet, type EditorView } from '@tiptap/pm/view';
import type { Node as PMNode } from '@tiptap/pm/model';

const key = new PluginKey('studioPagination');
type Piece = { pos: number; end: number; top: number; height: number; columns?: number; explicit?: boolean };

/** Use real block/row geometry. A form is often one large table cell containing
 * paragraphs and nested tables, so treating only top-level nodes as indivisible
 * blocks lets text run straight through page margins. */
function pieces(view: EditorView, usableHeight: number, scale: number, removed: { top: number; height: number }[]) {
  const origin = view.dom.getBoundingClientRect().top;
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
    if (node.type.name === 'pageBreak' || node.type.name === 'sectionBreak') {
      result.push({ pos, end: pos + node.nodeSize, ...box, height: 0, explicit: true }); return;
    }
    if (node.type.name !== 'table') { result.push({ pos, end: pos + node.nodeSize, ...box }); return; }
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
      if (height > usableHeight && end === i + 1 && row.node.childCount === 1) {
        const cell = row.node.firstChild!;
        cell.forEach((block, offset) => walk(block, row.pos + 2 + offset));
      } else {
        let columns = 0; row.node.forEach(cell => { columns += Number(cell.attrs.colspan || 1); });
        result.push({ pos: row.pos, end: rows[end - 1].pos + rows[end - 1].node.nodeSize, top: first.top, height, columns });
      }
      i = end;
    }
  };
  view.state.doc.forEach(walk);
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
            const h = Number(sheet.dataset.pageHeight) || 1122.52, top = Number(sheet.dataset.margin ?? 96);
            const bottom = Number(sheet.dataset.bottomMargin ?? top), stride = h + 24, usable = h - top - bottom;
            const origin = view.dom.getBoundingClientRect().top, scale = view.dom.getBoundingClientRect().width / view.dom.offsetWidth || 1;
            const removed = [...view.dom.querySelectorAll<HTMLElement>('.pagination-gap,.page-break,.section-break')].map(el => {
              const r = el.getBoundingClientRect(); return { top: (r.top - origin) / scale, height: r.height / scale };
            });
            const dec: Decoration[] = [], sizes: string[] = [];
            let extra = 0;
            for (const piece of pieces(view, usable, scale, removed)) {
              const y = piece.top + extra, page = Math.floor((y - top + .5) / stride);
              const end = page * stride + h - bottom;
              if (!piece.explicit && (y + piece.height <= end + .5 || piece.height > usable)) continue;
              const gap = Math.max(0, (page + 1) * stride + top - y);
              if (piece.explicit) dec.push(Decoration.node(piece.pos, piece.end, { style: `height:${gap}px` }));
              else dec.push(Decoration.widget(piece.pos, () => {
                const el = document.createElement(piece.columns ? 'tr' : 'div');
                el.className = `pagination-gap${piece.columns ? ' pagination-row-gap' : ''}`;
                el.style.height = `${gap}px`; el.contentEditable = 'false'; el.setAttribute('aria-hidden', 'true');
                if (piece.columns) { const cell = document.createElement('td'); cell.colSpan = piece.columns; cell.style.height = `${gap}px`; el.append(cell); }
                return el;
              }, { side: -1, key: `gap-${piece.pos}-${gap}` }));
              extra += gap; sizes.push(`${piece.pos}:${gap.toFixed(3)}`);
            }
            const last = view.dom.lastElementChild as HTMLElement | null;
            const naturalBottom = last ? (last.getBoundingClientRect().bottom - origin) / scale - removed.reduce((n, g) => n + g.height, 0) : top;
            const count = Math.max(1, Math.ceil((naturalBottom + extra + bottom - .5) / stride));
            sheet.dataset.pages = String(count); sheet.dispatchEvent(new CustomEvent('pagecount', { bubbles: true, detail: count }));
            const next = `${h}:${top}:${bottom}:${sizes.join(',')}`;
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
