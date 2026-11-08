import { useEffect, useRef, useState, type CSSProperties } from 'react';
import { EditorContent, useEditor } from '@tiptap/react';
import { generateHTML, generateJSON } from '@tiptap/core';
import { EditorState } from '@tiptap/pm/state';
import { version } from '../../package.json';
import { editorExtensions, parseDocumentHtml } from '../extensions';
import { DocumentHistory } from '../document-history';
import { importFile, exportDocx, validateDocument, sanitizeImportedHtml, documentFileLimit } from '../document-io';
import { defaultPage, pageMargins, paperDimensions, type StudioDocument } from '../types';
import { createPrintLayout, type PrintLayoutProgress } from '../print-layout';
import { withoutEditorReflow } from '../editor-layout-task';
import { readHistory, saveRecovery } from '../recovery';
import { importEditableHwp } from '../hwp/editing';
import { HwpReader } from '../hwp/HwpReader';
import type { EditableHwp } from '../hwp/conversion';
import { Sheet } from './Sheet';
import { TableTools } from './TableTools';
import { PaperViewer, paperDocument, type PaperDocument } from './PaperViewer';
import { androidServices, saveAndroidFile, shareAndroidFile, type AndroidFile } from './platform';
import './android.css';

const blank = (): StudioDocument => ({ id: crypto.randomUUID(), format: 'word-studio', version: 1, title: '문서1',
  page: { ...defaultPage }, comments: [], content: generateJSON('<p></p>', editorExtensions), updatedAt: new Date().toISOString() });
type Panel = 'menu' | 'find' | 'pages' | 'view' | 'share' | 'table' | 'recent' | 'leave' | null;

export default function AndroidApp() {
  const [documentState, setDocumentState] = useState(blank);
  const [opened, setOpened] = useState(false), [editing, setEditing] = useState(false), [dirty, setDirty] = useState(false);
  const [original, setOriginal] = useState<File | null>(null), [hwp, setHwp] = useState<File | null>(null);
  const [papers, setPapers] = useState<PaperDocument | null>(null), [panel, setPanel] = useState<Panel>(null);
  const [zoom, setZoom] = useState(100), [fit, setFit] = useState(true), [width, setWidth] = useState(0);
  const [page, setPage] = useState(0), [navigation, setNavigation] = useState({ page: 0, id: 0 });
  const [query, setQuery] = useState(''), [match, setMatch] = useState(0);
  const [message, setMessage] = useState(''), [status, setStatus] = useState(''), [revision, setRevision] = useState(0), [contentRevision, setContentRevision] = useState(0);
  const [operation, setOperation] = useState<{ label: string; controller: AbortController; cancellable: boolean } | null>(null);
  const pickerPending = useRef(false);
  const pendingIncoming = useRef<AndroidFile | null>(null);
  const [sharing, setSharing] = useState(false);
  const sharePending = useRef(false);
  const [composing, setComposing] = useState(false);
  const input = useRef<HTMLInputElement>(null), editorScroll = useRef<HTMLDivElement>(null);
  const running = useRef<AbortController | null>(null), layout = useRef<PaperDocument | null>(null);
  const state = useRef({ documentState, dirty, opened, original, editing, panel, hwp }); state.current = { documentState, dirty, opened, original, editing, panel, hwp };
  const continuation = useRef<(() => void) | null>(null);
  const installed = useRef(false), recoveryEpoch = useRef(0);
  const editor = useEditor({ extensions: [...editorExtensions, DocumentHistory], content: documentState.content, editable: false,
    editorProps: { attributes: { 'aria-label': 'Android 문서 편집 영역', spellcheck: 'false' },
      transformPastedHTML: html => sanitizeImportedHtml(html).html },
    onUpdate({ transaction }) { if (installed.current && state.current.editing && transaction.docChanged) { setDirty(true); setContentRevision(value => value + 1); setRevision(value => value + 1); } },
    onSelectionUpdate() { setRevision(value => value + 1); },
  });
  const snapshot = () => validateDocument({ ...state.current.documentState, content: editor?.getJSON() ?? state.current.documentState.content, updatedAt: new Date().toISOString() });
  useEffect(() => { editor?.setEditable(editing && !operation, false); }, [editor, editing, operation]);
  useEffect(() => {
    document.documentElement.dataset.android = 'true';
    const viewport = window.visualViewport;
    const resize = () => { document.documentElement.style.setProperty('--android-height', `${viewport && viewport.scale === 1 ? viewport.height : window.innerHeight}px`); };
    resize(); viewport?.addEventListener('resize', resize);

    return () => { delete document.documentElement.dataset.android; document.documentElement.style.removeProperty('--android-height'); viewport?.removeEventListener('resize', resize); layout.current?.layout.dispose(); running.current?.abort(); };
  }, []);
  useEffect(() => {
    const picker = input.current;
    const cancelPicker = () => { pickerPending.current = false; setMessage('파일 선택을 취소했습니다'); };
    picker?.addEventListener('cancel', cancelPicker);
    return () => picker?.removeEventListener('cancel', cancelPicker);
  }, [editor]);
  useEffect(() => { document.title = `${opened ? documentState.title : 'Word Studio'} — ${editing ? '편집' : '읽기'}`; }, [opened, documentState.title, editing]);
  useEffect(() => {
    if (!editor || !dirty || !opened) return;
    const epoch = ++recoveryEpoch.current;
    const timer = setTimeout(() => {
      void saveRecovery(snapshot()).then(() => { if (epoch === recoveryEpoch.current) setStatus('기기에 복구본 보관됨 · 원본 파일은 변경되지 않았습니다'); })
        .catch(() => { if (epoch === recoveryEpoch.current) setStatus('복구 저장 실패 · 사본을 저장하세요'); });
    }, 600);
    return () => clearTimeout(timer);
  }, [revision, dirty, opened, documentState, editor]);
  useEffect(() => {
    const before = (event: BeforeUnloadEvent) => { if (state.current.dirty) { event.preventDefault(); event.returnValue = ''; void saveRecovery(snapshot()).catch(() => {}); } };
    window.addEventListener('beforeunload', before);
    return () => window.removeEventListener('beforeunload', before);
  }, [editor]);
  const run = async (label: string, work: (signal: AbortSignal) => Promise<void>, cancellable = true) => {
    if (running.current || sharePending.current) return;
    const controller = new AbortController(); running.current = controller; setOperation({ label, controller, cancellable }); setMessage('');
    try { await new Promise(resolve => setTimeout(resolve, 0)); controller.signal.throwIfAborted(); await work(controller.signal); }
    catch (error) { if (!controller.signal.aborted) setMessage(error instanceof Error ? error.message : String(error)); }
    finally { if (running.current === controller) { running.current = null; setOperation(null);
      const incoming = pendingIncoming.current; pendingIncoming.current = null;
      if (incoming) setTimeout(() => callbacks.current.incoming(incoming), 0);
    } }
  };
  const cancel = () => { running.current?.abort(); running.current = null; setOperation(null); setMessage('작업을 취소했습니다');
    const incoming = pendingIncoming.current; pendingIncoming.current = null; if (incoming) setTimeout(() => callbacks.current.incoming(incoming), 0);
  };
  const protect = (next: () => void) => {
    if (running.current || sharePending.current || composing) return;
    if (state.current.dirty) { continuation.current = next; setPanel('leave'); }
    else { setPanel(null); next(); }
  };
  const reportLayout = (signal: AbortSignal) => (progress: PrintLayoutProgress) => {
    if (!signal.aborted) setOperation(current => current?.controller.signal === signal ? { ...current, label: progress.label + (progress.total ? ` ${progress.completed ?? 0}/${progress.total}` : '') } : current);
  };
  const install = async (candidate: StudioDocument, file: File | null, signal: AbortSignal, edit = false) => {
    if (!editor) return;
    const next = validateDocument(candidate), node = editor.schema.nodeFromJSON(next.content); node.check();
    const prepared = edit ? null : await withoutEditorReflow(() => createPrintLayout(generateHTML(next.content, editorExtensions), next.page, reportLayout(signal), signal));
    if (signal.aborted) { prepared?.dispose(); signal.throwIfAborted(); }
    installed.current = false;
    editor.view.updateState(EditorState.create({ schema: editor.schema, doc: node, plugins: editor.state.plugins }));
    installed.current = true; ++recoveryEpoch.current;
    layout.current?.layout.dispose(); layout.current = prepared ? paperDocument(prepared) : null;
    setPapers(layout.current); setDocumentState(next); setOriginal(file); setHwp(null); setOpened(true); setEditing(edit); setDirty(false);
    setPage(0); setNavigation(value => ({ page: 0, id: value.id + 1 })); setFit(true); setPanel(null); setQuery(''); setStatus(file ? '原本'.replace('原本', '原본 파일을 읽고 있습니다') : '새 문서');
  };
  const readFile = (file: File) => run('문서를 읽고 쪽 배치를 준비하는 중…', async signal => {
    if (/\.(hwp|hwpx)$/i.test(file.name)) { signal.throwIfAborted(); setHwp(file); setPanel(null); return; }
    const result = await importFile(file); signal.throwIfAborted();
    await install({ ...result.document, content: result.html === undefined ? result.document.content : parseDocumentHtml(result.html) }, file, signal);
    if (result.warnings.length) setMessage(result.warnings.join(' '));
  });
  const openFile = () => protect(() => {
    if (pickerPending.current) return;
    if (androidServices()?.openFile) void run('파일 선택창을 기다리는 중…', async signal => {
      const file = await androidServices()!.openFile!(); signal.throwIfAborted();
      if (!file) return;
      setOperation(current => current ? { ...current, cancellable: true, label: '문서를 읽는 중…' } : current);
      // Import within this same guarded request, never recursively retry it.
      const inputFile = new File([new Uint8Array(file.data)], file.name, { type: file.type });
      if (/\.(hwp|hwpx)$/i.test(file.name)) { setHwp(inputFile); return; }
      const result = await importFile(inputFile); signal.throwIfAborted();
      await install({ ...result.document, content: result.html === undefined ? result.document.content : parseDocumentHtml(result.html) }, inputFile, signal);
      if (result.warnings.length) setMessage(result.warnings.join(' '));
    }, false);
    else { pickerPending.current = true; input.current?.click(); }
  });
  const showReading = (nextPanel: Panel = null) => run('읽기 화면의 쪽 배치를 준비하는 중…', async signal => {
    const prepared = await withoutEditorReflow(() => createPrintLayout(editor!.getHTML(), snapshot().page, reportLayout(signal), signal));
    if (signal.aborted) { prepared.dispose(); return; }
    layout.current?.layout.dispose(); layout.current = paperDocument(prepared); setPapers(layout.current); setEditing(false); setPanel(nextPanel);
    setNavigation(value => ({ page: Math.min(page, prepared.pages - 1), id: value.id + 1 }));
  });
  const editHwp = async (source: EditableHwp, signal?: AbortSignal) => {
    const file = state.current.hwp; if (!file || !editor) return;
    const ownSignal = signal ?? new AbortController().signal;
    if (state.current.dirty) await saveRecovery(snapshot());
    ownSignal.throwIfAborted();
    await install(validateDocument(importEditableHwp(source, file.name)), file, ownSignal, true);
    setMessage('Word 편집본입니다. 사본으로 저장하며 한글 원본은 유지됩니다.');
  };
  const outputFile = async () => {
    const blob = await exportDocx(snapshot());
    if (blob.size > documentFileLimit('docx')) throw new Error('저장할 DOCX가 지원하는 크기를 넘습니다.');
    return new File([blob], `${state.current.documentState.title.replace(/[<>:"/\\|?*]/g, '-') || '문서'}.docx`, { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  };
  const save = () => run('Word 문서 사본을 준비하는 중…', async signal => {
    const file = await outputFile(); signal.throwIfAborted();
    // Native writes cannot be cancelled by merely abandoning their promise.
    // Keep the guard until the OS host acknowledges completion/cancellation.
    if (androidServices()?.saveFile) setOperation(current => current ? { ...current, cancellable: false, label: '시스템의 사본 저장 작업을 기다리는 중…' } : current);
    const result = await saveAndroidFile(file); signal.throwIfAborted();
    if (result !== 'cancelled') {
      try { await saveRecovery(snapshot()); } catch { setMessage('파일 작업 후 복구본을 보관하지 못했습니다. 저장된 사본을 확인하세요.'); }
    }
    if (result === 'completed') { setDirty(false); setStatus('Word 문서 사본 저장 완료'); }
    else if (result === 'download-requested') setMessage('DOCX 다운로드를 요청했습니다. 다운로드 완료 여부를 확인하세요.');
    else setMessage('저장을 취소했습니다');
  });
  const share = (source: boolean) => {
    // Invoke native/browser share in the user gesture when the original is
    // available; do not hide it behind unrelated asynchronous conversion.
    if (source && original) {
      void runShare(original); return;
    }
    void run('공유할 Word 편집본을 준비하는 중…', async signal => {
      const file = await outputFile(); signal.throwIfAborted();
      // Browser activation can expire while exporting: leave a visible second
      // tap on the prepared file rather than claiming the share completed.
      setPreparedShare(file); setMessage('사본이 준비됐습니다. 준비된 사본 공유를 누르세요.');
    });
  };
  const [preparedShare, setPreparedShare] = useState<File | null>(null);
  const runShare = async (file: File) => {
    if (running.current || sharePending.current) return;
    sharePending.current = true; setSharing(true);
    try { const result = await shareAndroidFile(file); setMessage(result === 'completed' ? '공유창에 파일을 전달했습니다' : '공유를 취소했습니다'); }
    catch (error) { setMessage(error instanceof Error ? error.message : String(error)); }
    finally { sharePending.current = false; setSharing(false);
      const incoming = pendingIncoming.current; pendingIncoming.current = null; if (incoming) setTimeout(() => callbacks.current.incoming(incoming), 0);
    }
  };
  const back = () => {
    if (running.current || sharePending.current) return;
    const sheet = document.querySelector<HTMLDialogElement>('.android-sheet[open]');
    if (sheet) { sheet.dispatchEvent(new Event('cancel', { cancelable: true })); return; }
    if (state.current.panel) { setPanel(null); return; }
    if (state.current.hwp) { setHwp(null); return; }
    if (state.current.editing) { void showReading(); return; }
    if (!state.current.opened) { androidServices()?.close?.(); return; }
    protect(() => { setOpened(false); setEditing(false); setOriginal(null); layout.current?.layout.dispose(); layout.current = null; setPapers(null); setMessage(''); });
  };
  const callbacks = useRef({ back, incoming: (file: AndroidFile) => protect(() => void readFile(new File([new Uint8Array(file.data)], file.name, { type: file.type }))) });
  callbacks.current = { back, incoming: file => {
    if (running.current || sharePending.current) {
      if (pendingIncoming.current) setMessage('이미 열기 대기 중인 파일이 있습니다. 다음 파일은 메뉴에서 다시 열어 주세요.');
      else pendingIncoming.current = file;
      return;
    }
    protect(() => void readFile(new File([new Uint8Array(file.data)], file.name, { type: file.type })));
  } };
  useEffect(() => { const services = androidServices(); const offOpen = services?.onOpenFile?.(file => callbacks.current.incoming(file)); const offBack = services?.onBack?.(() => callbacks.current.back()); return () => { offOpen?.(); offBack?.(); }; }, []);
  useEffect(() => {
    if (editing && editorScroll.current) {
      const dims = paperDimensions(documentState.page);
      const height = documentState.page.orientation === 'portrait' ? dims[1] : dims[0];
      editorScroll.current.scrollTop = page * (height + 24);
    }
  }, [editing]);
  const shortcuts = useRef<(event: KeyboardEvent) => void>(() => {});
  shortcuts.current = event => {
    if (!(event.ctrlKey || event.metaKey) || event.isComposing || composing || document.querySelector('dialog[open]')) return;
    const key = event.key.toLowerCase();
    if (!['s', 'o', 'n', 'f', 'p'].includes(key)) return;
    event.preventDefault();
    if (running.current || sharePending.current) return;
    if (key === 's' && opened) void save();
    else if (key === 'o') openFile();
    else if (key === 'n') protect(() => void run('새 문서를 준비하는 중…', signal => install(blank(), null, signal, true)));
    else if (key === 'f' && opened) { if (editing) void showReading('find'); else setPanel('find'); }
    else if (key === 'p') setMessage('PDF·인쇄는 Android 출력 연결 후 제공됩니다.');
  };
  useEffect(() => { const key = (event: KeyboardEvent) => shortcuts.current(event); window.addEventListener('keydown', key); return () => window.removeEventListener('keydown', key); }, []);
  const sizes = papers?.sizes ?? [];
  const widest = Math.max(1, ...sizes.map(size => size.width));
  useEffect(() => { if (fit && width && sizes.length) setZoom(Math.max(10, Math.min(200, Math.floor((width - 24) / widest * 100)))); }, [fit, width, widest, papers]);
  useEffect(() => { setPreparedShare(null); }, [contentRevision, documentState, original]);
  const results = papers && query.trim() ? papers.text.flatMap((text, i) => text.toLocaleLowerCase().includes(query.trim().toLocaleLowerCase()) ? [i] : []) : [];
  const go = (number: number) => { if (!Number.isInteger(number) || !papers || number < 0 || number >= papers.sheets.length) return; setPage(number); setNavigation(value => ({ page: number, id: value.id + 1 })); };
  const command = (action: () => void) => { if (!composing && !running.current) { action(); setRevision(value => value + 1); } };
  if (!editor) return <div className="android-start" role="status">문서 기능을 준비하는 중…</div>;
  const pageSettings = documentState.page, margins = pageMargins(pageSettings), dimensions = paperDimensions(pageSettings);
  const paperWidth = pageSettings.orientation === 'portrait' ? dimensions[0] : dimensions[1];
  const paperHeight = pageSettings.orientation === 'portrait' ? dimensions[1] : dimensions[0];
  const button = (label: string, action: () => void, active = false) => <button disabled={!!operation || composing} aria-pressed={active} onMouseDown={event => event.preventDefault()} onClick={() => command(action)}>{label}</button>;
  return <div className="android-shell" data-mode={editing ? 'edit' : 'read'}>
    <input ref={input} type="file" hidden accept=".docx,.hwp,.hwpx,.wordstudio,.html,.htm,.txt"
      onChange={event => { pickerPending.current = false; const file = event.target.files?.[0]; event.target.value = ''; if (file) protect(() => void readFile(file)); }} />
    <header className="android-topbar"><button aria-label="뒤로" disabled={!!operation || !opened} onClick={back}>‹</button>
      <div><strong>{opened ? documentState.title : 'Word Studio'}</strong><small>{opened ? editing ? 'Word 편집본' : dirty ? '편집본 읽기' : '읽기' : '문서 열기 · 원본 보기'}</small></div>
      <button aria-label="문서 메뉴" disabled={!!operation} onClick={() => setPanel('menu')}>더보기</button></header>
    {!opened && <main className="android-start"><h1>문서를 열어보세요</h1><p>DOCX · HWP · HWPX</p><button onClick={openFile}>파일 열기</button>
      <button onClick={() => protect(() => void run('새 문서를 준비하는 중…', signal => install(blank(), null, signal, true)))}>새 문서</button>
      <button onClick={() => setPanel('recent')}>최근 문서·복구본</button></main>}
    {opened && !editing && papers && <PaperViewer document={papers} highlight={query} zoom={zoom} navigation={navigation} onPage={setPage} onWidth={setWidth} />}
    {opened && editing && <>
      <div className="android-edit-tools" role="toolbar" aria-label="기본 편집 도구">
        {button('실행 취소', () => editor.chain().focus().undo().run())}{button('다시 실행', () => editor.chain().focus().redo().run())}
        {button('굵게', () => editor.chain().focus().toggleBold().run(), editor.isActive('bold'))}
        {button('기울임', () => editor.chain().focus().toggleItalic().run(), editor.isActive('italic'))}
        {button('밑줄', () => editor.chain().focus().toggleUnderline().run(), editor.isActive('underline'))}
        {button('왼쪽 정렬', () => editor.chain().focus().setTextAlign('left').run())}{button('가운데 정렬', () => editor.chain().focus().setTextAlign('center').run())}
        {button('표 도구', () => setPanel('table'))}
      </div>
      <div className="android-editor-scroll" ref={editorScroll} onCompositionStart={() => setComposing(true)} onCompositionEnd={() => setComposing(false)}>
        <div className="document-stack" data-page-height={paperHeight} data-margin={margins.top * 96 / 25.4} data-bottom-margin={margins.bottom * 96 / 25.4}
          style={{ width: paperWidth, minHeight: paperHeight, '--page-width': `${paperWidth}px`, '--page-height': `${paperHeight}px`, '--page-margin': `${margins.top * 96 / 25.4}px`,
            '--page-padding': `${margins.top * 96 / 25.4}px ${margins.right * 96 / 25.4}px ${margins.bottom * 96 / 25.4}px ${margins.left * 96 / 25.4}px` } as CSSProperties}>
          <EditorContent editor={editor} />
        </div>
      </div>
    </>}
    {opened && <footer className="android-reader-tools" aria-label="문서 도구">
      <button disabled={!!operation} onClick={() => { if (editing) void showReading('find'); else setPanel('find'); }}>찾기</button>
      <button disabled={!!operation || editing} onClick={() => setPanel('pages')}>{editing ? '편집 중' : `쪽 ${page + 1}/${papers?.sheets.length ?? '…'}`}</button>
      <button disabled={!!operation} onClick={() => editing ? void showReading() : setPanel('view')}>{editing ? '읽기' : '보기'}</button>
      <button disabled={!!operation} onClick={() => setPanel('share')}>공유·저장</button>
      <button disabled={!!operation || composing} onClick={() => { if (editing) void showReading(); else { setEditing(true); setPanel(null); } }}>{editing ? '편집 완료' : '편집'}</button>
    </footer>}
    {status && opened && <div className="android-status">{status}</div>}
    {message && !panel && <div className="android-message" role="status"><span>{message}</span><button aria-label="알림 닫기" onClick={() => setMessage('')}>닫기</button></div>}
    {operation && <div className="android-operation" role="status"><progress aria-label={operation.label} /><span>{operation.label}</span>{operation.cancellable ? <button onClick={cancel}>취소</button> : <p>시스템 선택창에서 작업을 완료하거나 취소해 주세요.</p>}</div>}
    {panel && <Sheet title={panel === 'menu' ? '문서 메뉴' : panel === 'find' ? '문서에서 찾기' : panel === 'pages' ? '쪽 이동' : panel === 'view' ? '보기 설정' : panel === 'share' ? '공유·사본 저장' : panel === 'table' ? '표 도구' : panel === 'recent' ? '최근 문서·복구본' : '변경 내용 보관'} onClose={() => setPanel(null)}>
      {message && <p role="status">{message}</p>}
      {panel === 'menu' && <><button onClick={openFile}>다른 파일 열기</button><button onClick={() => protect(() => void run('새 문서를 준비하는 중…', signal => install(blank(), null, signal, true)))}>새 문서</button>
        <button onClick={() => setPanel('recent')}>최근 문서·복구본</button>
        {original && /\.(hwp|hwpx)$/i.test(original.name) && <button onClick={() => { setPanel(null); setHwp(original); }}>한글 원본 보기</button>}{opened && <button onClick={() => setPanel('share')}>공유·사본 저장</button>}
        {opened && <label>문서 제목<input aria-label="문서 제목" value={documentState.title} onChange={event => { setDocumentState(doc => ({ ...doc, title: event.target.value })); setDirty(true); }} /></label>}
        <p>Word Studio {version} · Android</p><p>이 버전은 문서를 기기에서 처리합니다. 공유는 직접 선택한 파일만 전달합니다.</p></>}
      {panel === 'view' && <><button onClick={() => { setFit(true); setZoom(Math.max(10, Math.min(200, Math.floor((width - 24) / widest * 100)))); }}>폭에 맞추기</button>
        <label>문서 확대 <input aria-label="문서 확대" type="range" min="10" max="300" step="1" value={zoom} onChange={event => { setFit(false); setZoom(Number(event.target.value)); }} /><output>{zoom}%</output></label>
        <button onClick={() => { setFit(false); setZoom(value => Math.max(10, value - 10)); }}>축소 −</button><button onClick={() => { setFit(false); setZoom(value => Math.min(300, value + 10)); }}>확대 +</button>
        <p>용지·표·그림 배치는 그대로 두고 표시 배율만 바꿉니다. 두 손가락 확대는 기기의 기본 확대 동작을 사용합니다.</p></>}
      {panel === 'pages' && <><label>쪽 번호<input aria-label="쪽 번호" type="number" min="1" max={papers?.sheets.length} value={page + 1} onChange={event => go(Number(event.target.value) - 1)} /></label>
        <div className="android-action-row"><button disabled={page <= 0} onClick={() => go(page - 1)}>이전 쪽</button><button disabled={page + 1 >= (papers?.sheets.length ?? 0)} onClick={() => go(page + 1)}>다음 쪽</button></div><button onClick={() => setPanel(null)}>문서로 돌아가기</button></>}
      {panel === 'find' && <><label>찾을 내용<input autoFocus aria-label="찾을 내용" value={query} onChange={event => { setQuery(event.target.value); setMatch(0); }} /></label>
        <p role="status">{query.trim() ? `${results.length}개 쪽에서 찾았습니다` : '문서 전체에서 찾습니다'}</p>
        <button disabled={!results.length} onClick={() => { go(results[match % results.length]); setMatch(value => value + 1); setPanel(null); setMessage(`${results[match % results.length] + 1}쪽에서 찾았습니다`); }}>다음 결과로 이동</button></>}
      {panel === 'share' && <><button disabled={!original || !!operation || sharing} onClick={() => share(true)}>원본 파일 공유</button>
        <button disabled={!!operation} onClick={() => share(false)}>Word 편집본 사본 준비</button>{preparedShare && <button disabled={sharing} onClick={() => void runShare(preparedShare)}>준비된 사본 공유</button>}
        <button disabled={!!operation} onClick={() => void save()}>Word 문서 사본 저장 (.docx)</button>
        <p>원본은 덮어쓰지 않습니다. PDF·인쇄는 Android 출력 연결 후 제공됩니다.</p></>}
      {panel === 'recent' && <>{readHistory().length ? readHistory().map(doc => <button key={doc.id} onClick={() => protect(() => void run('복구본을 읽는 중…', signal => install(doc, null, signal)))}>{doc.title} · {new Date(doc.updatedAt).toLocaleString()}</button>) : <p>보관된 복구본이 없습니다.</p>}</>}
      {panel === 'leave' && <><p>변경한 내용을 파일 사본이나 기기의 복구본에 보관한 뒤 이동할 수 있습니다.</p><button onClick={() => { continuation.current = null; setPanel(null); }}>계속 편집</button>
        <button onClick={() => void save()}>Word 문서 사본 저장</button><button onClick={() => void run('복구본을 보관하는 중…', async signal => {
          await saveRecovery(snapshot()); signal.throwIfAborted(); setDirty(false); setPanel(null);
          const next = continuation.current; continuation.current = null;
          // The current guarded operation must finish before the next starts.
          setTimeout(() => next?.(), 0);
        })}>복구본 보관 후 계속</button></>}
      {panel === 'table' && <><TableTools editor={editor} disabled={!!operation || composing} onChange={() => setRevision(value => value + 1)} /><button onClick={() => setPanel(null)}>편집으로 돌아가기</button></>}
    </Sheet>}
    {hwp && <HwpReader key={`${hwp.name}:${hwp.lastModified}:${hwp.size}`} file={hwp} android onClose={() => setHwp(null)} onOpen={openFile} onEdit={editHwp} />}
  </div>;
}
