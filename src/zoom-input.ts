import { useEffect, useRef, type RefObject } from 'react';

type ZoomAt = (scale: number, x: number, y: number) => void;
type Options = { readScale(): number; onZoom: ZoomAt; min?: number; max?: number; touchActive?(): boolean };
type Gesture = Event & { scale?: number; clientX?: number; clientY?: number };

/** Trackpad pinch in Chromium is Ctrl+wheel; WebKit also sends gesture events.
 * Unmodified wheel/trackpad scrolling keeps its normal browser behavior. */
export function bindZoomInput(viewport: HTMLElement, { readScale, onZoom, min = .1, max = 4, touchActive }: Options) {
  const maxStep = Math.log(1.25);
  let observed = readScale(), pending = observed;
  let gesture: { base: number; x: number; y: number } | null = null;
  const current = () => { const value = readScale(); if (value !== observed) pending = observed = value; return pending; };
  const point = (event: Gesture) => {
    const r = viewport.getBoundingClientRect();
    return { x: Number.isFinite(event.clientX) ? event.clientX! - r.left : viewport.clientWidth / 2,
      y: Number.isFinite(event.clientY) ? event.clientY! - r.top : viewport.clientHeight / 2 };
  };
  const apply = (value: number, x: number, y: number) => {
    if (!Number.isFinite(value) || value <= 0) return;
    const next = Math.min(max, Math.max(min, value));
    if (next === current()) return;
    pending = next; onZoom(next, x, y);
  };
  const wheel = (event: WheelEvent) => {
    if (!event.ctrlKey && !event.metaKey) return;
    event.preventDefault(); event.stopPropagation();
    if (gesture) return;
    const delta = event.deltaY * (event.deltaMode === 1 ? 16 : event.deltaMode === 2 ? viewport.clientHeight : 1);
    if (!Number.isFinite(delta) || !delta) return;
    const p = point(event);
    apply(current() * Math.exp(Math.max(-maxStep, Math.min(maxStep, -delta * .01))), p.x, p.y);
  };
  const start = (event: Gesture) => {
    event.preventDefault(); event.stopPropagation();
    if (touchActive?.()) return;
    const p = point(event), factor = event.scale && event.scale > 0 ? event.scale : 1;
    gesture = { base: current() / factor, ...p };
  };
  const move = (event: Gesture) => {
    event.preventDefault(); event.stopPropagation();
    if (!gesture || touchActive?.()) return;
    const p = event.clientX === undefined ? gesture : point(event);
    apply(gesture.base * (event.scale ?? 1), p.x, p.y);
  };
  const end = (event: Gesture) => { if (gesture) { event.preventDefault(); event.stopPropagation(); if (event.scale !== undefined) move(event); } gesture = null; };
  const cancel = () => { gesture = null; };
  viewport.addEventListener('wheel', wheel, { passive: false });
  viewport.addEventListener('gesturestart', start, { passive: false });
  viewport.addEventListener('gesturechange', move, { passive: false });
  viewport.addEventListener('gestureend', end, { passive: false });
  window.addEventListener('blur', cancel);
  return () => {
    viewport.removeEventListener('wheel', wheel); viewport.removeEventListener('gesturestart', start);
    viewport.removeEventListener('gesturechange', move); viewport.removeEventListener('gestureend', end);
    window.removeEventListener('blur', cancel);
  };
}

export function useZoomInput(viewport: RefObject<HTMLElement | null>, scale: number, onZoom: ZoomAt, enabled = true, options: { min?: number; max?: number; touchActive?(): boolean } = {}) {
  const latest = useRef({ scale, onZoom, touchActive: options.touchActive }); latest.current = { scale, onZoom, touchActive: options.touchActive };
  const { min = .1, max = 4 } = options;
  useEffect(() => {
    const element = viewport.current; if (!element || !enabled) return;
    return bindZoomInput(element, { readScale: () => latest.current.scale, onZoom: (next, x, y) => latest.current.onZoom(next, x, y),
      min, max, touchActive: () => latest.current.touchActive?.() ?? false });
  }, [viewport, enabled, min, max]);
}
