import { Extension } from '@tiptap/core';
import { Plugin, PluginKey } from '@tiptap/pm/state';
import { Decoration, DecorationSet, type EditorView } from '@tiptap/pm/view';
import { closeHistory } from '@tiptap/pm/history';
import type { Node as PMNode } from '@tiptap/pm/model';

export const MAX_ROW_HEIGHT = 2000;
const MIN_DRAG_HEIGHT = 12;
const HANDLE_SIZE = 5;

export function rowHeight(value: unknown): number | null {
  if (typeof value !== 'number' || !Number.isFinite(value) || value <= 0 || value > MAX_ROW_HEIGHT) return null;
  return value;
}

function readHeight(element: HTMLElement): number | null {
  const raw = element.style.height || element.getAttribute('height') || '';
  const match = raw.trim().match(/^(\d+(?:\.\d+)?)\s*(px|pt|mm|cm|in)?$/i);
  if (!match) return null;
  const units: Record<string, number> = { px: 1, pt: 96 / 72, mm: 96 / 25.4, cm: 96 / 2.54, in: 96 };
  return rowHeight(Math.min(MAX_ROW_HEIGHT, Number(match[1]) * units[match[2]?.toLowerCase() || 'px']));
}

type Preview = { pos: number; height: number } | null;
type Hit = { pos: number; row: HTMLTableRowElement; table: HTMLTableElement };
type Drag = Hit & { startY: number; startHeight: number; scale: number; height: number; doc: PMNode };
const resizeKey = new PluginKey<Preview>('studioTableRowResize');

function hitAt(view: EditorView, event: MouseEvent): Hit | null {
  if (!view.editable || view.composing) return null;
  const doc = view.dom.ownerDocument;
  // Also allow grabbing the last border just outside the table itself.
  const targets = [event.target, doc.elementFromPoint(event.clientX, event.clientY - HANDLE_SIZE)];
  for (const target of targets) {
    if (!(target instanceof Element)) continue;
    const cell = target.closest<HTMLTableCellElement>('td,th');
    if (!cell || !view.dom.contains(cell)) continue;
    const table = cell.closest('table')!, sourceRow = cell.parentElement as HTMLTableRowElement;
    const bounds = cell.getBoundingClientRect();
    // At intersections, retain the existing column-width resize handles.
    if (event.clientX <= bounds.left + HANDLE_SIZE || event.clientX >= bounds.right - HANDLE_SIZE) continue;
    let index = -1;
    if (Math.abs(event.clientY - bounds.bottom) <= HANDLE_SIZE) index = sourceRow.rowIndex + cell.rowSpan - 1;
    else if (Math.abs(event.clientY - bounds.top) <= HANDLE_SIZE) index = sourceRow.rowIndex - 1;
    if (index < 0 || index >= table.rows.length) continue;
    const row = table.rows[index];
    const $pos = view.state.doc.resolve(view.posAtDOM(row, 0));
    for (let depth = $pos.depth; depth > 0; depth--) {
      if ($pos.node(depth).type.name === 'tableRow') return { pos: $pos.before(depth), row, table };
    }
  }
  return null;
}

export const TableRowResize = Extension.create({
  name: 'tableRowResize',
  // Handle horizontal edges before table selection and column dragging.
  priority: 200,
  addGlobalAttributes() {
    return [{ types: ['tableRow'], attributes: {
      rowHeight: {
        default: null,
        parseHTML: readHeight,
        renderHTML: attrs => rowHeight(attrs.rowHeight) ? { style: `height:${attrs.rowHeight}px` } : {},
      },
      wordRowRule: { default: null, parseHTML: el => el.getAttribute('data-word-row-rule'),
        renderHTML: a => a.wordRowRule === 'exact' ? { 'data-word-row-rule': 'exact' } : {} },
    } }];
  },
  addProseMirrorPlugins() {
    let drag: Drag | null = null, hovered: Hit | null = null;
    let guide: HTMLDivElement, label: HTMLSpanElement, view: EditorView;
    let destroyed = false;

    const hide = () => {
      hovered = null;
      if (guide) guide.hidden = true;
      view?.dom.classList.remove('row-resize-cursor');
    };
    const show = (hit: Hit) => {
      const bounds = hit.table.getBoundingClientRect(), row = hit.row.getBoundingClientRect();
      const scale = view.dom.getBoundingClientRect().width / view.dom.offsetWidth || 1;
      guide.hidden = false;
      guide.style.left = `${bounds.left}px`;
      guide.style.top = `${row.bottom}px`;
      guide.style.width = `${bounds.width}px`;
      label.textContent = `${(row.height / scale * 25.4 / 96).toFixed(1)} mm`;
      view.dom.classList.add('row-resize-cursor');
    };
    const cleanup = () => {
      const win = view.dom.ownerDocument.defaultView!;
      win.removeEventListener('mousemove', move, true);
      win.removeEventListener('mouseup', release, true);
      win.removeEventListener('keydown', key, true);
      win.removeEventListener('blur', cancel);
      win.removeEventListener('resize', cancel);
      view.dom.classList.remove('table-row-resizing');
      view.dom.ownerDocument.body.classList.remove('table-row-resizing');
      hide();
    };
    const finish = (commit: boolean) => {
      if (!drag) return;
      const current = drag;
      drag = null;
      cleanup();
      if (destroyed) return;
      let tr = view.state.tr.setMeta(resizeKey, null).setMeta('addToHistory', false);
      const node = view.state.doc.nodeAt(current.pos);
      if (commit && view.editable && view.state.doc === current.doc && node?.type.name === 'tableRow' && Math.abs(current.height - current.startHeight) >= 1) {
        tr = closeHistory(tr.setNodeMarkup(current.pos, null, { ...node.attrs, rowHeight: current.height, wordRowRule: null }).setMeta('addToHistory', true));
      }
      view.dispatch(tr);
    };
    const cancel = () => finish(false);
    const key = (event: KeyboardEvent) => {
      if (event.key === 'Escape') { event.preventDefault(); event.stopPropagation(); cancel(); }
    };
    const scrolled = () => { if (drag) cancel(); else hide(); };
    const move = (event: MouseEvent) => {
      if (!drag) return;
      if (!(event.buttons & 1)) { finish(true); return; }
      event.preventDefault(); event.stopPropagation();
      drag.height = Math.max(MIN_DRAG_HEIGHT, Math.min(MAX_ROW_HEIGHT, Math.round(drag.startHeight + (event.clientY - drag.startY) / drag.scale)));
      view.dispatch(view.state.tr.setMeta(resizeKey, { pos: drag.pos, height: drag.height }).setMeta('addToHistory', false));
      if (drag) show(drag);
    };
    const release = (event: MouseEvent) => {
      if (event.button !== 0) return;
      event.preventDefault(); event.stopPropagation();
      finish(true);
    };

    return [new Plugin<Preview>({
      key: resizeKey,
      state: {
        init: () => null,
        apply(tr, old) {
          const preview = tr.getMeta(resizeKey);
          if (preview !== undefined) return preview;
          // Concurrent editing ends the gesture instead of resizing a different row.
          return tr.docChanged ? null : old;
        },
      },
      props: {
        decorations(state) {
          const preview = resizeKey.getState(state);
          if (!preview) return DecorationSet.empty;
          const row = state.doc.nodeAt(preview.pos);
          return row?.type.name === 'tableRow' ? DecorationSet.create(state.doc, [Decoration.node(preview.pos, preview.pos + row.nodeSize,
            { class: 'row-height-preview', style: `--row-preview-height:${preview.height}px` })]) : DecorationSet.empty;
        },
        handleDOMEvents: {
          mousemove(current, event) {
            if (drag) return false;
            if (event.buttons) { hide(); return false; }
            hovered = hitAt(current, event);
            if (hovered) show(hovered); else hide();
            return false;
          },
          mouseleave() { if (!drag) hide(); return false; },
          mousedown(current, event) {
            if (event.button !== 0) return false;
            const hit = hitAt(current, event);
            if (!hit) return false;
            event.preventDefault();
            const scale = current.dom.getBoundingClientRect().width / current.dom.offsetWidth || 1;
            const height = hit.row.getBoundingClientRect().height / scale;
            drag = { ...hit, startY: event.clientY, startHeight: height, scale, height, doc: current.state.doc };
            current.dom.classList.add('table-row-resizing');
            current.dom.ownerDocument.body.classList.add('table-row-resizing');
            const win = current.dom.ownerDocument.defaultView!;
            win.addEventListener('mousemove', move, true);
            win.addEventListener('mouseup', release, true);
            win.addEventListener('keydown', key, true);
            win.addEventListener('blur', cancel);
            win.addEventListener('resize', cancel);
            show(hit);
            return true;
          },
          dblclick(current, event) {
            const hit = hitAt(current, event);
            if (!hit) return false;
            event.preventDefault();
            const node = current.state.doc.nodeAt(hit.pos)!;
            if (node.attrs.rowHeight != null) current.dispatch(closeHistory(current.state.tr.setNodeMarkup(hit.pos, null, { ...node.attrs, rowHeight: null, wordRowRule: null })));
            hide();
            return true;
          },
        },
      },
      view(current) {
        destroyed = false;
        view = current;
        guide = current.dom.ownerDocument.createElement('div');
        guide.className = 'table-row-resize-guide'; guide.hidden = true; guide.setAttribute('aria-hidden', 'true');
        label = current.dom.ownerDocument.createElement('span'); guide.append(label);
        current.dom.ownerDocument.body.append(guide);
        current.dom.ownerDocument.defaultView!.addEventListener('scroll', scrolled, true);
        return {
          update(next, previous) {
            view = next;
            if (drag && (!next.editable || next.state.doc !== drag.doc)) cancel();
            else if (!drag && hovered) {
              if (!next.editable || next.state.doc !== previous.doc) hide();
              else show(hovered);
            }
          },
          destroy() {
            destroyed = true; if (drag) cancel(); hide(); guide.remove();
            current.dom.ownerDocument.defaultView!.removeEventListener('scroll', scrolled, true);
          },
        };
      },
    })];
  },
});
