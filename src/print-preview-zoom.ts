import { useEffect, useRef, useState } from 'react';
import type { PrintLayout } from './print-layout';
import { bindZoomInput } from './zoom-input';

/** Preview magnification is screen CSS only; print media uses original pages. */
export function usePrintPreviewZoom(layout: PrintLayout | null) {
  const [scale, setScale] = useState(1), reset = useRef<(() => void) | null>(null);
  useEffect(() => {
    setScale(1); if (!layout) return;
    const viewport = layout.element, pages = [...viewport.querySelectorAll<HTMLElement>(':scope > .print-page')];
    let current = 1;
    viewport.dataset.displayScale = '1';
    const change = (next: number, cx: number, cy: number) => {
      const vr = viewport.getBoundingClientRect(), x = vr.left + cx, y = vr.top + cy;
      const page = pages.find(page => page.getBoundingClientRect().bottom >= y) ?? pages.at(-1); if (!page) return;
      const before = page.getBoundingClientRect(), px = (x - before.left) / current, py = (y - before.top) / current;
      viewport.style.setProperty('--print-preview-scale', String(next));
      const after = page.getBoundingClientRect();
      viewport.scrollLeft += after.left + px * next - x; viewport.scrollTop += after.top + py * next - y;
      current = next; viewport.dataset.displayScale = String(next); setScale(next);
    };
    reset.current = () => change(1, viewport.clientWidth / 2, 76);
    const dispose = bindZoomInput(viewport, { readScale: () => current, onZoom: change });
    return () => { dispose(); reset.current = null; viewport.style.removeProperty('--print-preview-scale'); delete viewport.dataset.displayScale; };
  }, [layout]);
  return { scale, reset: () => reset.current?.() };
}
