import { useEffect, useLayoutEffect, useRef, useState } from 'react';
import { ModeSwitch } from './ModeSwitch';
import { Printer } from 'lucide-react';

type Size = { width: number; height: number };
type Drag = { x: number; y: number; left: number; top: number; moved: boolean; multi: boolean; leftEdge: boolean; rightEdge: boolean };
function Sheet({ number, size, scale, top, left, current, loadPage }: { number: number; size: Size; scale: number; top: number; left: number; current: boolean; loadPage(index: number): Promise<HTMLElement> }) {
  const host = useRef<HTMLDivElement>(null), loader = useRef(loadPage); loader.current = loadPage;
  const [error, setError] = useState(''), [ready, setReady] = useState(false), [retry, setRetry] = useState(0);
  useEffect(() => {
    let active = true; const root = host.current!; setReady(false); setError('');
    void loader.current(number).then(page => { if (!active) return; page.style.transform = 'none'; page.style.margin = '0'; root.replaceChildren(page); setReady(true); })
      .catch(e => { if (active) setError(e instanceof Error ? e.message : String(e)); });
    return () => { active = false; root.replaceChildren(); };
  }, [number, retry]);
  return <section className="android-presentation-slot" aria-label={`${number + 1}쪽`} aria-current={current ? 'page' : undefined} data-document-page={number + 1}
    style={{ position: 'absolute', top, left, width: size.width * scale, height: size.height * scale }}>
    <div ref={host} className={`android-presentation-sheet${current ? ' android-presentation-page' : ''}`} style={{ width: size.width, height: size.height, transform: `scale(${scale})` }} />
    {!ready && <div className="android-page-loading" role={error ? 'alert' : 'status'}>{error ? <>{error}<button onClick={() => setRetry(x => x + 1)}>다시 읽기</button></> : `${number + 1}쪽을 읽는 중…`}</div>}
  </section>;
}
/** A continuous virtual page stack: scale belongs to the document, never to a page. */
export function PagePresentation({ source, sizes, index, loadPage, onPage, onExit, onEdit, onPrint, notice = '' }: {
  source: object; sizes: Size[]; index: number; loadPage(index: number): Promise<HTMLElement>;
  onPage(index: number): void; onExit(): void; onEdit?(): void; onPrint?(): void; notice?: string;
}) {
  const root = useRef<HTMLElement>(null), viewport = useRef<HTMLDivElement>(null);
  const [box, setBox] = useState({ width: 1, height: 1 }), [manualScale, setManualScale] = useState<number | null>(null);
  const [controls, setControls] = useState(true), [range, setRange] = useState({ first: 0, last: 1 });
  const callbacks = useRef({ onPage, onExit, onEdit }); callbacks.current = { onPage, onExit, onEdit };
  const timer = useRef<ReturnType<typeof setTimeout> | undefined>(undefined), reported = useRef<number | null>(null);
  const widest = Math.max(1, ...sizes.map(s => s.width)), scale = manualScale ?? box.width / widest;
  const offsets: number[] = []; let logicalHeight = 0;
  for (const size of sizes) { offsets.push(logicalHeight); logicalHeight += size.height + 24; }
  const stageWidth = Math.max(box.width, widest * scale), left = (stageWidth - widest * scale) / 2;
  const size = sizes[index] ?? sizes[0] ?? { width: 794, height: 1123 };
  const stageHeight = Math.max(box.height, (logicalHeight - 24) * scale + Math.max(0, box.height - (sizes.at(-1)?.height ?? 0) * scale));
  const anchor = useRef<{ x: number; y: number; cx: number; cy: number } | null>(null);
  const points = useRef(new Map<number, { x: number; y: number }>()), drag = useRef<Drag | null>(null), pinch = useRef<{ distance: number; scale: number; x: number; y: number } | null>(null);
  const tap = useRef({ at: 0, x: 0, y: 0 });
  const show = () => { setControls(true); clearTimeout(timer.current); timer.current = setTimeout(() => setControls(false), 2200); };
  useEffect(() => { show(); return () => clearTimeout(timer.current); }, []);
  useLayoutEffect(() => {
    const node = root.current!, exit = () => callbacks.current.onExit(), edit = () => (callbacks.current.onEdit ?? callbacks.current.onExit)();
    node.addEventListener('studio-presentation-close', exit); node.addEventListener('studio-presentation-escape', exit); node.addEventListener('studio-presentation-edit', edit);
    return () => { node.removeEventListener('studio-presentation-close', exit); node.removeEventListener('studio-presentation-escape', exit); node.removeEventListener('studio-presentation-edit', edit); };
  }, []);
  useLayoutEffect(() => { const v = viewport.current!, resize = () => setBox({ width: v.clientWidth, height: v.clientHeight }); const observer = new ResizeObserver(resize); observer.observe(v); resize(); return () => observer.disconnect(); }, []);
  useEffect(() => { setManualScale(null); reported.current = null; }, [source]);
  const currentAt = (top: number) => { let n = 0; while (n + 1 < sizes.length && offsets[n + 1] * scale <= top + 1) n++; return n; };
  useLayoutEffect(() => {
    const v = viewport.current!, a = anchor.current;
    if (a) { v.scrollLeft = a.x * scale + left - a.cx; v.scrollTop = a.y * scale - a.cy; anchor.current = null; }
  }, [scale, left]);
  useLayoutEffect(() => {
    if (!box.width || box.width === 1) return;
    const v = viewport.current!;
    if (reported.current !== index) v.scrollTop = (offsets[index] ?? 0) * scale;
    reported.current = index;
  }, [index, source, box.width]);
  useEffect(() => {
    const v = viewport.current!; let frame = 0;
    const refresh = () => {
      if (box.width <= 1) return;
      const first = currentAt(v.scrollTop), last = currentAt(v.scrollTop + v.clientHeight);
      setRange(old => old.first === Math.max(0, first - 1) && old.last === Math.min(sizes.length - 1, last + 1) ? old : { first: Math.max(0, first - 1), last: Math.min(sizes.length - 1, last + 1) });
      if (reported.current !== first) { reported.current = first; callbacks.current.onPage(first); }
    };
    const schedule = () => { cancelAnimationFrame(frame); frame = requestAnimationFrame(refresh); };
    v.addEventListener('scroll', schedule, { passive: true }); refresh();
    return () => { cancelAnimationFrame(frame); v.removeEventListener('scroll', schedule); };
  }, [scale, sizes, box.height]);
  const zoom = (next: number, cx = box.width / 2, cy = 0, fixed?: { x: number; y: number }) => {
    const v = viewport.current!; anchor.current = { x: fixed?.x ?? (v.scrollLeft + cx - left) / scale, y: fixed?.y ?? (v.scrollTop + cy) / scale, cx, cy };
    setManualScale(Math.min(8, Math.max(.1, next)));
  };
  const go = (next: number) => { if (next < 0 || next >= sizes.length) return; reported.current = next; viewport.current!.scrollTop = offsets[next] * scale; callbacks.current.onPage(next); show(); };
  const local = (x: number, y: number) => { const r = viewport.current!.getBoundingClientRect(); return { x: x - r.left, y: y - r.top }; };
  useEffect(() => {
    const key = (e: KeyboardEvent) => {
      if (e.key === 'Escape') { e.preventDefault(); e.stopImmediatePropagation(); callbacks.current.onExit(); return; }
      if (e.target instanceof Element && e.target.closest('input,textarea,select')) return;
      const actions: Record<string, () => void> = { ArrowRight: () => go(index + 1), ArrowDown: () => go(index + 1), PageDown: () => go(index + 1), ArrowLeft: () => go(index - 1), ArrowUp: () => go(index - 1), PageUp: () => go(index - 1), Home: () => go(0), End: () => go(sizes.length - 1), '+': () => zoom(scale * 1.25), '-': () => zoom(scale / 1.25) };
      if (actions[e.key]) { e.preventDefault(); e.stopPropagation(); actions[e.key](); }
    };
    window.addEventListener('keydown', key, true); return () => window.removeEventListener('keydown', key, true);
  }, [index, scale, box, sizes]);
  return <section ref={root} className="android-presentation" aria-label="전체 화면 문서 뷰어" data-page={index + 1} data-controls={controls} data-scale={scale}>
    <div ref={viewport} className="android-presentation-viewport" tabIndex={0} aria-label="연속 문서 뷰어" onPointerDown={e => {
      if (e.pointerType === 'mouse') return;
      e.currentTarget.setPointerCapture(e.pointerId); points.current.set(e.pointerId, { x: e.clientX, y: e.clientY });
      const v = viewport.current!;
      if (points.current.size === 1) drag.current = { x: e.clientX, y: e.clientY, left: v.scrollLeft, top: v.scrollTop, moved: false, multi: false, leftEdge: v.scrollLeft <= 2, rightEdge: v.scrollLeft >= v.scrollWidth - v.clientWidth - 2 };
      else { const [a, b] = [...points.current.values()], c = local((a.x + b.x) / 2, (a.y + b.y) / 2); pinch.current = { distance: Math.max(1, Math.hypot(a.x - b.x, a.y - b.y)), scale, x: (v.scrollLeft + c.x - left) / scale, y: (v.scrollTop + c.y) / scale }; if (drag.current) drag.current.multi = true; }
    }} onPointerMove={e => {
      if (!points.current.has(e.pointerId)) return; points.current.set(e.pointerId, { x: e.clientX, y: e.clientY });
      const [a, b] = [...points.current.values()], p = pinch.current, d = drag.current;
      if (a && b && p) { const c = local((a.x + b.x) / 2, (a.y + b.y) / 2); zoom(p.scale * Math.hypot(a.x - b.x, a.y - b.y) / p.distance, c.x, c.y, p); }
      else if (d && !d.multi) { const dx = e.clientX - d.x, dy = e.clientY - d.y; if (Math.hypot(dx, dy) > 10) d.moved = true; viewport.current!.scrollLeft = d.left - dx; viewport.current!.scrollTop = d.top - dy; }
    }} onPointerUp={e => {
      const active = points.current.has(e.pointerId), d = drag.current; points.current.delete(e.pointerId); if (!active || !d || points.current.size) return;
      if (!d.multi) {
        const dx = e.clientX - d.x, dy = e.clientY - d.y;
        if (d.moved) {
          if (Math.abs(dx) > Math.abs(dy) * 1.2 && Math.abs(dx) >= 56 && (dx < 0 ? d.rightEdge : d.leftEdge)) go(index + (dx < 0 ? 1 : -1));
          else if (Math.abs(dy) > Math.abs(dx) * 1.2 && Math.abs(dy) >= 56) {
            const start = currentAt(d.top), finish = currentAt(viewport.current!.scrollTop);
            // A deliberate drag also turns a fitted page; a larger page keeps panning.
            if (sizes[start].height * scale <= box.height + 2) go(dy < 0 ? Math.max(start + 1, finish) : Math.min(start - 1, finish));
          }
        } else if (!d.moved) { const now = performance.now(); if (now - tap.current.at < 320 && Math.hypot(tap.current.x - e.clientX, tap.current.y - e.clientY) < 28) { const c = local(e.clientX, e.clientY); zoom(scale > box.width / widest * 1.1 ? box.width / widest : box.width / widest * 2, c.x, c.y); tap.current.at = 0; show(); } else { tap.current = { at: now, x: e.clientX, y: e.clientY }; if (controls) { setControls(false); clearTimeout(timer.current); } else show(); } }
      }
      drag.current = null; pinch.current = null;
    }} onPointerCancel={() => { points.current.clear(); drag.current = null; pinch.current = null; }}>
      <div className="print-layout android-presentation-canvas" style={{ width: stageWidth, height: stageHeight }}>
        {sizes.map((s, n) => n >= range.first && n <= range.last ? <Sheet key={n} number={n} size={s} scale={scale} top={offsets[n] * scale} left={(stageWidth - s.width * scale) / 2} current={n === index} loadPage={loadPage} /> : null)}
      </div>
    </div>
    <button className="android-presentation-exit" aria-label="뷰어 닫기" onClick={onExit}>‹</button>
    <div className="android-presentation-modebar"><span className="android-presentation-page-number" role="status">{index + 1} / {sizes.length}</span><>{onPrint && <button aria-label="인쇄·PDF 저장" onClick={onPrint}><Printer size={20} aria-hidden="true" /></button>}</><ModeSwitch editing={false} onEdit={onEdit} onView={show} /></div>
    <div className="android-presentation-controls" hidden={!controls}>
      <footer><button aria-label="이전 페이지" disabled={!index} onClick={() => go(index - 1)}>‹</button><button aria-label="뷰어 축소" onClick={() => { zoom(scale / 1.25); show(); }}>−</button><button aria-label="뷰어 폭 맞춤" onClick={() => { zoom(box.width / widest); show(); }}>{Math.round(scale * 100)}%</button><button aria-label="뷰어 확대" onClick={() => { zoom(scale * 1.25); show(); }}>＋</button><button aria-label="다음 페이지" disabled={index + 1 === sizes.length} onClick={() => go(index + 1)}>›</button></footer>
      <div className="android-presentation-options"><button onClick={() => { zoom(Math.min(box.width / widest, box.height / size.height)); show(); }}>쪽 전체</button><span>좌우로 넘기기 · 상하로 스크롤</span></div>{notice && <p role="status">{notice}</p>}
    </div>
  </section>;
}
