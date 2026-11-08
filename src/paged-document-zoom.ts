import { useRef, type RefObject } from 'react';
import { hwpPageAtOffset, hwpViewportLayout } from './hwp/viewport-layout';
import { useZoomInput } from './zoom-input';

type Size = { width: number; height: number };
/** Virtual readers keep page gaps in screen pixels. Anchor in page coordinates
 * rather than multiplying the scroll offset (which would scale every gap). */
export function usePagedDocumentZoom(viewport: RefObject<HTMLElement | null>, host: RefObject<HTMLElement | null>, pages: readonly Size[], scale: number, onScale: (scale: number) => void, enabled = true) {
  const anchor = useRef<{ pages: readonly Size[]; page: number; x: number; y: number; cx: number; cy: number } | null>(null);
  const change = (next: number, cx: number, cy: number) => {
    const v = viewport.current, content = host.current; if (!v || !content || !pages.length) return;
    const vr = v.getBoundingClientRect(), cr = content.getBoundingClientRect();
    const layout = hwpViewportLayout(pages, scale), page = hwpPageAtOffset(layout, Math.max(0, vr.top + cy - cr.top));
    anchor.current = { pages, page, cx, cy, x: (vr.left + cx - cr.left - (cr.width - pages[page].width * scale) / 2) / scale,
      y: (vr.top + cy - cr.top - layout.offsets[page]) / scale };
    onScale(next);
  };
  useZoomInput(viewport, scale, change, enabled);
  const restore = () => {
    const a = anchor.current, v = viewport.current, content = host.current; anchor.current = null;
    if (!a || a.pages !== pages || !v || !content) return false;
    const vr = v.getBoundingClientRect(), cr = content.getBoundingClientRect(), layout = hwpViewportLayout(pages, scale);
    v.scrollLeft += cr.left + (cr.width - pages[a.page].width * scale) / 2 + a.x * scale - vr.left - a.cx;
    v.scrollTop += cr.top + layout.offsets[a.page] + a.y * scale - vr.top - a.cy;
    return true;
  };
  return restore;
}
