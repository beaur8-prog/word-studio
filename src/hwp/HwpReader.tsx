import { useEffect, useRef, useState } from 'react';
import { HwpReaderSession, safeHwpSvg, type HwpInfo, type HwpPage } from './reader';
import '@fontsource/noto-sans-kr/400.css';
import '@fontsource/noto-sans-kr/700.css';
import '@fontsource/noto-serif-kr/400.css';
import '@fontsource/noto-serif-kr/700.css';
import './reader.css';
import { HWP_CONVERSION_NOTICE, type EditableHwp } from './conversion';
import { sanitizeImportedHtml } from '../document-html';
import { previewEditableHwp } from './editing';

export function HwpReader({ file, onClose, onOpen, onEdit }: { file: File; onClose(): void; onOpen(): void; onEdit(source: EditableHwp): Promise<void> }) {
  const dialog = useRef<HTMLDialogElement>(null), session = useRef<HwpReaderSession | null>(null);
  const currentRequest = useRef(0), cache = useRef(new Map<number, HwpPage>());
  const [info, setInfo] = useState<HwpInfo>(), [page, setPage] = useState<HwpPage>();
  const [index, setIndex] = useState(0), [zoom, setZoom] = useState(85);
  const [loading, setLoading] = useState(true), [exporting, setExporting] = useState(false);
  const [error, setError] = useState(''), [status, setStatus] = useState('');
  const [converting, setConverting] = useState(false), [editable, setEditable] = useState<EditableHwp>();
  const [query, setQuery] = useState(''), [searching, setSearching] = useState(false);
  const searchInput = useRef<HTMLInputElement>(null);
  const viewport = useRef<HTMLElement>(null);
  const exportingRef = useRef(false);
  const getPage = async (number: number) => {
    if (cache.current.has(number)) return cache.current.get(number)!;
    if (!session.current) throw new Error('읽기 화면을 다시 열어 주세요.');
    const result = await session.current.page(number);
    if (cache.current.size >= 8) cache.current.delete(cache.current.keys().next().value!);
    cache.current.set(number, result); return result;
  };
  useEffect(() => {
    const previous = document.activeElement as HTMLElement | null;
    const previousTitle = document.title;
    document.title = `${file.name} — Word Studio (읽기)`;
    dialog.current?.showModal();
    cache.current.clear(); setInfo(undefined); setPage(undefined); setIndex(0); setError(''); setStatus(''); setLoading(true); setSearching(false);
    const reader = new HwpReaderSession(); session.current = reader;
    let active = true;
    void reader.open(file).then(value => { if (active) setInfo(value); }).catch(reason => {
      if (active) { setError(String(reason.message)); setLoading(false); }
    });
    return () => { active = false; ++currentRequest.current; reader.close(); session.current = null; document.title = previousTitle; if (previous?.isConnected) previous.focus(); };
  }, [file]);
  useEffect(() => {
    if (!info) return;
    const request = ++currentRequest.current;
    setLoading(true); setError('');
    viewport.current?.scrollTo({ top: 0, left: 0 });
    void getPage(index).then(value => { if (request === currentRequest.current) { setPage(value); setLoading(false); } })
      .catch(reason => { if (request === currentRequest.current) { setError(reason.message); setLoading(false); } });
  }, [info, index]);
  const findNext = async () => {
    if (!info || !query.trim() || searching) return;
    const request = currentRequest.current, value = query.replace(/\s/g, '').toLocaleLowerCase();
    setSearching(true); setStatus('찾는 중…');
    try {
      for (let offset = 1; offset <= info.count; offset++) {
        const number = (index + offset) % info.count;
        const next = await getPage(number);
        if (request !== currentRequest.current) return;
        if (next.text.replace(/\s/g, '').toLocaleLowerCase().includes(value)) { setIndex(number); setStatus(`${number + 1}쪽에서 찾았습니다.`); return; }
      }
      setStatus('찾는 내용이 없습니다.');
    } catch (reason) { setStatus(String(reason)); }
    finally { setSearching(false); }
  };
  const copyText = async () => {
    try { if (session.current) {
      const text = await session.current.text();
      if (window.wordDesktop?.copyText) await window.wordDesktop.copyText(text); else await navigator.clipboard.writeText(text);
      setStatus('문서 텍스트를 복사했습니다.');
    } }
    catch { setStatus('텍스트를 선택한 뒤 ⌘C / Ctrl+C로 복사하세요.'); }
  };
  const prepareEditing = async () => {
    const reader = session.current;
    if (!reader || converting) return;
    setConverting(true); setStatus('Word 편집본을 준비하는 중…');
    try {
      const source = await reader.editable();
      if (session.current !== reader) return;
      const clean = sanitizeImportedHtml(source.html);
      const preview = await previewEditableHwp({ ...source, html: clean.html, warnings: [...source.warnings, ...clean.warnings] }, file.name);
      if (session.current !== reader) return;
      setEditable(preview);
      setStatus('변환 안내를 확인한 뒤 편집본을 만드세요.');
    } catch (reason) { if (session.current === reader) setStatus(`편집 변환 실패: ${String(reason)}`); }
    finally { if (session.current === reader) setConverting(false); }
  };
  const print = async (pdf: boolean) => {
    if (!info || exportingRef.current) return;
    exportingRef.current = true; setExporting(true); setStatus('출력할 쪽을 준비하는 중…');
    const root = document.createElement('div'), style = document.createElement('style');
    root.className = 'print-layout hwp-print-layout';
    try {
      let size = 0;
      for (let number = 0; number < info.count; number++) {
        const item = await getPage(number); size += item.svg.length;
        if (size > 100 * 1024 * 1024) throw new Error('출력 데이터가 100MB를 넘습니다.');
        const sheet = document.createElement('div'); sheet.className = 'print-page hwp-print-page';
        sheet.style.cssText = `width:${item.width}px;height:${item.height}px;page:hwp-${number}`;
        sheet.innerHTML = safeHwpSvg(item.svg).svg; root.append(sheet);
        style.textContent += `@page hwp-${number}{size:${item.width}px ${item.height}px;margin:0;}`;
      }
      root.append(style); document.body.append(root); document.body.classList.add('studio-printing');
      await document.fonts.ready;
      if (window.wordDesktop) {
        const saved = pdf ? await window.wordDesktop.exportPdf({ name: file.name.replace(/\.(hwp|hwpx)$/i, '') + '.pdf', paginated: true }) : await window.wordDesktop.print();
        setStatus(saved ? (pdf ? 'PDF를 저장했습니다.' : '인쇄를 요청했습니다.') : '출력을 취소했습니다.');
      } else { window.print(); setStatus('인쇄 창에서 PDF로 저장할 수 있습니다.'); }
    } catch (reason) { setStatus(`출력 실패: ${String(reason)}`); }
    finally { root.remove(); document.body.classList.remove('studio-printing'); exportingRef.current = false; setExporting(false); }
  };
  useEffect(() => window.wordDesktop?.onMenu(action => {
    if (exportingRef.current) return;
    if (action === 'open') onOpen();
    else if (action === 'print') void print(false);
    else if (action === 'export-pdf') void print(true);
    else setStatus('HWP는 읽기 전용입니다. 편집 문서로 돌아가려면 읽기 화면을 닫으세요.');
  }), [info, file]);
  return <dialog ref={dialog} className="hwp-reader" aria-label="HWP 읽기" onCancel={event => { event.preventDefault(); if (!exporting) onClose(); }}
    onKeyDown={event => {
      if (!(event.ctrlKey || event.metaKey) || event.nativeEvent.isComposing) return;
      const key = event.key.toLowerCase();
      if (['s', 'o', 'n', 'p', 'f'].includes(key)) {
        event.preventDefault(); event.stopPropagation();
        if (exporting) return;
        if (key === 'f') searchInput.current?.focus();
        else if (key === 'o') onOpen();
        else if (key === 'p') void print(false);
        else setStatus('HWP는 읽기 전용입니다. 원본 파일은 변경되지 않습니다.');
      }
    }}>
    <header className="hwp-toolbar"><div><strong>{file.name}</strong><span>HWP/HWPX 원본 보기 · Word Studio</span></div>
      <button onClick={() => void prepareEditing()} disabled={!info || converting || exporting}>Word 화면에서 편집</button>
      <button onClick={onOpen} disabled={exporting}>다른 파일 열기</button>
      <button onClick={copyText} disabled={!info || exporting}>텍스트 복사</button>
      <button onClick={() => print(true)} disabled={!info || loading || exporting}>PDF 저장</button>
      <button onClick={() => print(false)} disabled={!info || loading || exporting}>인쇄</button>
      <button onClick={onClose} disabled={exporting} aria-label="HWP 읽기 닫기">닫기</button>
    </header>
    <div className="hwp-controls">
      <button aria-label="이전 쪽" onClick={() => setIndex(index - 1)} disabled={!info || index <= 0 || exporting || searching}>‹</button>
      <label>쪽 <input type="number" aria-label="HWP 쪽 번호" min="1" max={info?.count ?? 1} value={index + 1} disabled={!info || exporting || searching}
        onChange={event => { const number = Number(event.target.value); if (info && Number.isInteger(number) && number >= 1 && number <= info.count) setIndex(number - 1); }} /></label>
      <span data-testid="hwp-page-count">/ {info?.count ?? '…'}</span>
      <button aria-label="다음 쪽" onClick={() => setIndex(index + 1)} disabled={!info || index + 1 >= info.count || exporting || searching}>›</button>
      <label>확대 <select aria-label="HWP 확대" value={zoom} onChange={event => setZoom(Number(event.target.value))}>
        {[50, 65, 75, 85, 100, 125, 150, 200].map(value => <option key={value} value={value}>{value}%</option>)}
      </select></label>
      <form onSubmit={event => { event.preventDefault(); void findNext(); }}><input ref={searchInput} aria-label="HWP 문서에서 찾기" value={query} onChange={event => setQuery(event.target.value)} placeholder="문서에서 찾기" />
        <button disabled={!info || exporting || searching || !query.trim()}>다음 찾기</button></form>
    </div>
    <p className="hwp-notice">현재는 원본 보기입니다. ‘Word 화면에서 편집’을 누르면 본문·표·그림을 기존 편집 화면으로 가져옵니다. 설치 글꼴과 지원 범위에 따라 원본 뷰어와 배치가 다를 수 있습니다.
      {info?.warnings ? ` 원본 구조 주의 항목 ${info.warnings}개.` : ''}</p>
    {editable && <section className="hwp-conversion-panel" aria-label="한글 편집 변환 안내">
      <h2>Word 편집본 만들기</h2><p>{HWP_CONVERSION_NOTICE}</p>
      <p><strong>쪽 수 비교: 원본 보기 {info?.count}쪽 → Word 편집본 {editable.editorPages}쪽</strong>{info?.count !== editable.editorPages ? ' · 쪽 배치가 달라집니다. 정확한 쪽 참조가 필요한 문서는 원본을 함께 확인하세요.' : ' · 쪽 수가 같아도 세부 배치는 다를 수 있습니다.'}</p>
      <ul>{editable.warnings.map((warning, i) => <li key={i}>{warning}</li>)}</ul>
      <button onClick={() => void onEdit(editable).catch(reason => setStatus(`편집 변환 실패: ${String(reason)}`))}>편집본 만들기</button>
      <button onClick={() => { setEditable(undefined); setStatus('원본 보기를 계속합니다.'); }}>원본 보기로 돌아가기</button>
    </section>}
    <main ref={viewport} className="hwp-viewport" aria-busy={loading || exporting}>
      {error ? <div className="hwp-error" role="alert"><strong>HWP 파일을 열지 못했습니다</strong><p>{error}</p><button onClick={onClose}>편집 문서로 돌아가기</button></div>
        : page && <div className="hwp-sheet-space" style={{ width: page.width * zoom / 100, height: page.height * zoom / 100 }}>
          <div className="hwp-sheet" data-testid="hwp-page" aria-label={`${index + 1}쪽`} style={{ width: page.width, height: page.height, transform: `scale(${zoom / 100})` }} dangerouslySetInnerHTML={{ __html: page.svg }} />
        </div>}
      {loading && <p className="hwp-loading" role="status">HWP 문서를 읽는 중…</p>}
    </main>
    <footer className="hwp-status" role="status">{exporting ? '출력할 쪽을 준비하고 있습니다…' : status || (info ? `${info.count}쪽 · HWP ${info.version} · 글꼴 ${info.fonts.length}종` : '원본을 읽고 있습니다.')}</footer>
  </dialog>;
}
