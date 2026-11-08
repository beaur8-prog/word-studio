import { useEffect, useLayoutEffect, useRef } from 'react';
import { usePagedDocumentZoom } from '../paged-document-zoom';
import type { PrintLayout } from '../print-layout';
import { hwpViewportLayout, hwpVisiblePages } from '../hwp/viewport-layout';

export type PaperDocument = { layout: PrintLayout; sheets: HTMLElement[]; sizes: { width: number; height: number }[]; text: string[] };
export function paperDocument(layout: PrintLayout): PaperDocument {
  const sheets = [...layout.element.querySelectorAll<HTMLElement>(':scope > .print-page')];
  return { layout, sheets, sizes: sheets.map(sheet => ({ width: parseFloat(sheet.style.width), height: parseFloat(sheet.style.height) })),
    text: sheets.map(sheet => sheet.textContent || '') };
}
export function PaperViewer({ document: doc, zoom, highlight = '', navigation, onPage, onWidth, onZoom }: {
  document: PaperDocument; zoom: number; highlight?: string; navigation: { page: number; id: number }; onPage(page: number): void; onWidth(width: number): void; onZoom(value: number): void;
}) {
  const viewport = useRef<HTMLDivElement>(null), previous = useRef({ doc, zoom, top: 0 });
  const hostRef = useRef(doc.layout.element); hostRef.current = doc.layout.element;
  const restoreZoom = usePagedDocumentZoom(viewport, hostRef, doc.sizes, zoom / 100, value => onZoom(value * 100));
  const search = useRef(highlight); search.current = highlight;
  const callbacks = useRef({ onPage, onWidth }); callbacks.current = { onPage, onWidth };
  useLayoutEffect(() => {
    const root = viewport.current!, host = doc.layout.element, scale = zoom / 100;
    const layout = hwpViewportLayout(doc.sizes, scale);
    const old = previous.current;
    let targetTop = old.doc === doc ? old.top : 0;
    if (old.doc === doc && old.zoom !== zoom) {
      const oldLayout = hwpViewportLayout(doc.sizes, old.zoom / 100);
      const current = hwpVisiblePages(oldLayout, Math.max(0, old.top - 12), root.clientHeight).current;
      const within = Math.max(0, old.top - 12 - oldLayout.offsets[current]);
      targetTop = 12 + layout.offsets[current] + within * zoom / old.zoom;
    }
    previous.current = { doc, zoom, top: targetTop };
    host.className = 'print-layout android-paper-layer';
    host.style.width = `${layout.width}px`; host.style.height = `${layout.height + Math.max(0, root.clientHeight - (layout.heights.at(-1) ?? 0) - 24)}px`;
    root.append(host);
    const slots = doc.sheets.map((sheet, i) => {
      sheet.remove();
      const slot = document.createElement('section');
      slot.className = 'android-paper-slot'; slot.dataset.androidPage = String(i + 1); slot.setAttribute('aria-label', `${i + 1}쪽`);
      slot.style.cssText = `top:${layout.offsets[i]}px;width:${doc.sizes[i].width * scale}px;height:${doc.sizes[i].height * scale}px`;
      sheet.style.transform = `scale(${scale})`; slot.append(sheet); host.append(slot); return slot;
    });
    const refresh = () => {
      const visible = hwpVisiblePages(layout, Math.max(0, root.scrollTop - 12), root.clientHeight);
      slots.forEach((slot, i) => {
        const near = i >= visible.first - 1 && i <= visible.last + 1;
        if (near && !doc.sheets[i].isConnected) slot.append(doc.sheets[i]);
        else if (!near) doc.sheets[i].remove();
        if (i === visible.current) slot.setAttribute('aria-current', 'page'); else slot.removeAttribute('aria-current');
      });
      highlightPapers(doc.sheets.filter(sheet => sheet.isConnected), search.current);
      callbacks.current.onPage(visible.current);
      callbacks.current.onWidth(root.clientWidth);
    };
    let frame = 0;
    const schedule = () => { cancelAnimationFrame(frame); frame = requestAnimationFrame(refresh); };
    root.scrollTop = targetTop; restoreZoom();
    root.addEventListener('scroll', schedule, { passive: true });
    const observer = new ResizeObserver(schedule); observer.observe(root); refresh();
    return () => { previous.current = { doc, zoom, top: root.scrollTop }; cancelAnimationFrame(frame); observer.disconnect(); root.removeEventListener('scroll', schedule); slots.forEach(slot => slot.remove()); host.remove(); };
  }, [doc, zoom]);
  useEffect(() => { highlightPapers(doc.sheets.filter(sheet => sheet.isConnected), highlight); return () => clearHighlight(); }, [doc, highlight]);
  useEffect(() => {
    const layout = hwpViewportLayout(doc.sizes, zoom / 100);
    if (viewport.current) viewport.current.scrollTop = 12 + (layout.offsets[navigation.page] || 0);
    // A scroll-originated page change never rewrites the navigation request.
  }, [doc, navigation.id]);
  return <main ref={viewport} className="android-paper-viewport" data-scale={zoom / 100} tabIndex={0} aria-label="문서 읽기 영역" />;
}

// CSS highlights paint the original text ranges without inserting DOM wrappers,
// changing line metrics or modifying the document model.
function clearHighlight() { CSS.highlights?.delete('word-studio-android-find'); }
function highlightPapers(sheets: HTMLElement[], query: string) {
  clearHighlight();
  if (!query.trim() || !CSS.highlights || typeof Highlight === 'undefined') return;
  const ranges: Range[] = [];
  for (const sheet of sheets) {
    const walker = document.createTreeWalker(sheet, NodeFilter.SHOW_TEXT, { acceptNode: node => node.parentElement?.closest('style,script') ? NodeFilter.FILTER_REJECT : NodeFilter.FILTER_ACCEPT });
    const nodes: { node: Text; start: number; end: number }[] = []; let text = '';
    while (walker.nextNode()) { const node = walker.currentNode as Text; nodes.push({ node, start: text.length, end: text.length + node.length }); text += node.data; }
    const folded = text.toLocaleLowerCase(), needle = query.trim().toLocaleLowerCase();
    // Unicode case folding can change string length. Never apply wrong offsets.
    if (folded.length !== text.length) continue;
    let at = 0;
    while ((at = folded.indexOf(needle, at)) !== -1 && ranges.length < 2000) {
      const begin = nodes.find(item => item.end > at), end = nodes.find(item => item.end >= at + needle.length);
      if (begin && end) { const range = document.createRange(); range.setStart(begin.node, at - begin.start); range.setEnd(end.node, at + needle.length - end.start); ranges.push(range); }
      at += Math.max(1, needle.length);
    }
  }
  CSS.highlights.set('word-studio-android-find', new Highlight(...ranges));
}
