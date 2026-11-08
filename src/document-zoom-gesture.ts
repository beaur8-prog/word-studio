import { useEffect, useLayoutEffect, useRef, type RefObject } from 'react';
import { useZoomInput } from './zoom-input';

/** Display-only zoom. Never changes document measurements or text size. */
export function useDocumentZoom(viewport: RefObject<HTMLElement | null>, content: RefObject<HTMLElement | null>, scale: number, onScale: (scale: number) => void, enabled = true) {
  const latest = useRef({ scale, onScale }); latest.current = { scale, onScale };
  const touching = useRef(false);
  const anchor = useRef<{ x: number; y: number; cx: number; cy: number } | null>(null);
  const locate = (cx: number, cy: number) => {
    const v = viewport.current!, c = content.current!, vr = v.getBoundingClientRect(), cr = c.getBoundingClientRect();
    return { x: (vr.left + cx - cr.left) / latest.current.scale, y: (vr.top + cy - cr.top) / latest.current.scale, cx, cy };
  };
  const change = (next: number, cx?: number, cy?: number) => {
    const v = viewport.current; if (!v || !content.current) return;
    anchor.current = locate(cx ?? v.clientWidth / 2, cy ?? 0);
    latest.current.onScale(Math.min(4, Math.max(.1, next)));
  };
  useZoomInput(viewport, scale, change, enabled, { touchActive: () => touching.current });
  useLayoutEffect(() => {
    const a = anchor.current, v = viewport.current, c = content.current;
    if (!a || !v || !c) return;
    const vr = v.getBoundingClientRect(), cr = c.getBoundingClientRect();
    v.scrollLeft += cr.left + a.x * scale - vr.left - a.cx;
    v.scrollTop += cr.top + a.y * scale - vr.top - a.cy;
    anchor.current = null;
  }, [scale]);
  useEffect(() => {
    const v = viewport.current; if (!v || !enabled) return;
    let pinch: { distance: number; scale: number; point: { x: number; y: number } } | null = null, held = false;
    const center = (e: TouchEvent) => { const r = v.getBoundingClientRect(); return { x: (e.touches[0].clientX + e.touches[1].clientX) / 2 - r.left, y: (e.touches[0].clientY + e.touches[1].clientY) / 2 - r.top }; };
    const distance = (e: TouchEvent) => Math.hypot(e.touches[0].clientX - e.touches[1].clientX, e.touches[0].clientY - e.touches[1].clientY);
    const start = (e: TouchEvent) => {
      if (e.touches.length !== 2 || !content.current) return;
      e.preventDefault(); const c = center(e), point = locate(c.x, c.y);
      pinch = { distance: Math.max(1, distance(e)), scale: latest.current.scale, point }; held = true; touching.current = true;
    };
    const move = (e: TouchEvent) => {
      if (!held) return; e.preventDefault();
      if (!pinch || e.touches.length !== 2) return;
      const c = center(e); anchor.current = { ...pinch.point, cx: c.x, cy: c.y };
      latest.current.onScale(Math.min(4, Math.max(.1, pinch.scale * distance(e) / pinch.distance)));
    };
    const end = (e: TouchEvent) => { if (held) e.preventDefault(); pinch = null; if (!e.touches.length) { held = false; touching.current = false; } };
    v.addEventListener('touchstart', start, { passive: false }); v.addEventListener('touchmove', move, { passive: false });
    v.addEventListener('touchend', end, { passive: false }); v.addEventListener('touchcancel', end, { passive: false });
    return () => { touching.current = false; v.removeEventListener('touchstart', start); v.removeEventListener('touchmove', move); v.removeEventListener('touchend', end); v.removeEventListener('touchcancel', end); };
  }, [enabled, viewport, content]);
  return change;
}
