import { useCallback, useEffect, useRef, useState } from 'react';
import { Download, LoaderCircle, RefreshCw, X } from 'lucide-react';
import type { StudioUpdateAction, StudioUpdateState } from './desktop';

const releasePage = 'https://github.com/beaur8-prog/word-studio/releases/latest';
const working = (state: StudioUpdateState | null) => !!state && ['checking', 'downloading', 'verifying', 'installing'].includes(state.phase);
const size = (bytes: number) => `${(bytes / 1024 / 1024).toFixed(1)} MB`;
export function updateLabel(state: StudioUpdateState | null) {
  if (!state) return '업데이트 정보를 불러오는 중…';
  switch (state.phase) {
    case 'checking': return '새 버전을 확인하는 중…';
    case 'available': return `${state.version} 버전으로 업데이트할 수 있습니다.`;
    case 'downloading': return `업데이트 다운로드 중 · ${state.percent}%`;
    case 'verifying': return '다운로드 파일을 검증하는 중…';
    case 'ready': return `${state.version} 버전 설치 준비가 완료되었습니다.`;
    case 'installing': return '문서 저장·종료 후 업데이트를 설치합니다…';
    case 'error': return '업데이트를 완료하지 못했습니다.';
    default: return state.result === 'latest' ? '최신 버전을 사용하고 있습니다.' : '업데이트를 확인해 주세요.';
  }
}

export function useStudioUpdates() {
  const [state, setState] = useState<StudioUpdateState | null>(null);
  const [open, setOpen] = useState(false);
  const [pending, setPending] = useState<StudioUpdateAction | null>(null);
  const [error, setError] = useState('');
  useEffect(() => {
    const api = window.wordDesktop;
    if (!api?.getUpdateState) return;
    let active = true, received = false;
    const off = api.onUpdateState?.(next => { received = true; if (active) setState(next); });
    void api.getUpdateState().then(next => { if (active && !received) setState(next); })
      .catch(() => { if (active) setError('업데이트 정보를 읽지 못했습니다. 업데이트 확인을 눌러 다시 시도해 주세요.'); });
    return () => { active = false; off?.(); };
  }, []);
  const run = useCallback(async (action: StudioUpdateAction, value?: boolean) => {
    const api = window.wordDesktop;
    if (!api?.updateAction) return;
    setPending(action); setError('');
    if (action === 'automatic') setState(current => current ? { ...current, automatic: !!value } : current);
    try { setState(await api.updateAction(action, value)); }
    catch {
      setError('업데이트 요청을 처리하지 못했습니다. 잠시 후 다시 시도해 주세요.');
      if (action === 'automatic') { try { if (api.getUpdateState) setState(await api.getUpdateState()); } catch { /* Keep the actionable error visible. */ } }
    }
    finally { setPending(current => current === action ? null : current); }
  }, []);
  const show = useCallback((check = false) => {
    setOpen(true);
    if (check) void run('check');
  }, [run]);
  return { state, open, setOpen, pending, error, run, show };
}

type Updates = ReturnType<typeof useStudioUpdates>;
export function UpdateStatus({ updates }: { updates: Updates }) {
  const state = updates.state;
  if (!state || state.phase === 'idle') return null;
  return <button className={`update-status update-${state.phase}`} onClick={() => updates.show()} title="업데이트 진행 상황 보기">
    {working(state) ? <LoaderCircle className="update-spinner" size={14} aria-hidden="true" /> : <Download size={14} aria-hidden="true" />}
    <span role="status">{updateLabel(state)}</span>
  </button>;
}

export function UpdateCenter({ updates, currentVersion }: { updates: Updates; currentVersion: string }) {
  const ref = useRef<HTMLDialogElement>(null);
  const { state, open, pending, error, run } = updates;
  useEffect(() => {
    if (!open) return;
    const node = ref.current;
    node?.showModal();
    return () => { if (node?.open) node.close(); };
  }, [open]);
  if (!open) return null;
  const automaticSupported = !!state?.packaged && state.supported;
  const canUseUpdater = !!window.wordDesktop?.updateAction && (!state || automaticSupported);
  const active = working(state);
  const disabled = active || !!pending || !canUseUpdater;
  const manualHelp = !state ? '다운로드 페이지에서 사용 중인 운영체제에 맞는 설치 파일을 선택해 주세요.' : state.platform === 'darwin'
    ? 'Mac: 다운로드 파일을 열고 문서를 저장한 뒤 앱을 종료하세요. ZIP 압축을 풀어 응용 프로그램 폴더의 Word Studio를 교체하면 됩니다.'
    : state?.platform === 'linux'
      ? 'Linux: 다운로드 페이지에서 이 기기에 맞는 최신 패키지를 받아 설치해 주세요.'
      : 'Windows: 설치를 선택하면 문서 저장·종료를 확인한 뒤 Setup 설치 화면이 열립니다. 설치 과정을 확인할 수 있고, 완료되면 앱을 다시 엽니다. 종료를 취소하면 설치도 취소됩니다.';
  return <dialog ref={ref} className="dialog update-dialog" aria-labelledby="update-title"
    onCancel={event => { event.preventDefault(); updates.setOpen(false); }}>
    <div className="dialog-heading">
      <h2 id="update-title">Word Studio 업데이트</h2>
      <button className="icon-button" aria-label="업데이트 창 닫기" onClick={() => updates.setOpen(false)}><X size={20} /></button>
    </div>
    <dl className="update-versions"><div><dt>현재 버전</dt><dd>{state?.currentVersion || currentVersion}</dd></div>
      {state?.version && <div><dt>새 버전</dt><dd>{state.version}</dd></div>}</dl>
    <section className="update-progress" aria-label="업데이트 진행 상황" aria-busy={active}>
      <strong role="status">{active && <LoaderCircle className="update-spinner" size={18} aria-hidden="true" />}{!window.wordDesktop?.getUpdateState ? '수동 설치로 업데이트할 수 있습니다.' : error && !state ? '업데이트 정보를 불러오지 못했습니다.' : updateLabel(state)}</strong>
      {state?.phase === 'downloading' && <>
        <progress aria-label="업데이트 다운로드" max={100} value={state.percent} />
        <small>{size(state.transferred)} / {size(state.total)} · {state.percent}%</small>
      </>}
      {state && ['checking', 'verifying', 'installing'].includes(state.phase) && <progress aria-label={updateLabel(state)} />}
      {state?.phase === 'verifying' && <small>다운로드 파일이 정상인지 확인하고 있습니다.</small>}
      {state?.checkedAt && <small>마지막 확인: {new Date(state.checkedAt).toLocaleString('ko-KR')}</small>}
      {(state?.error || error) && <p className="update-error" role="alert">{error || state?.error}</p>}
    </section>
    <label className="update-automatic"><input type="checkbox" checked={state?.automatic ?? false}
      disabled={!automaticSupported || pending === 'automatic' || state?.phase === 'installing'}
      onChange={event => { void run('automatic', event.target.checked); }} />자동으로 업데이트 확인·다운로드</label>
    <p className="update-help">자동 업데이트를 끄면 아래 버튼으로 직접 확인하고 다운로드할 수 있습니다. 설치는 선택한 뒤에 진행됩니다.</p>
    {!canUseUpdater && <p className="update-help">{!state?.packaged
      ? '설치된 배포 앱에서 업데이트 확인·다운로드를 사용할 수 있습니다. 이 실행본은 다운로드 페이지에서 수동으로 업데이트해 주세요.'
      : '이 환경에서는 다운로드 페이지에서 수동으로 업데이트해 주세요.'}</p>}
    <p className="update-help">{manualHelp}</p>
    {state?.platform === 'darwin' && <p className="update-help">다운로드가 끝나도 Mac 앱이 자동으로 교체되지는 않습니다.</p>}
    <div className="update-actions">
      <button disabled={disabled} onClick={() => { void run('check'); }}><RefreshCw size={15} />업데이트 확인</button>
      {(state?.phase === 'available' || state?.phase === 'error' && state.version) &&
        <button className="primary" disabled={disabled} onClick={() => { void run('download'); }}>다운로드{state?.total ? ` (${size(state.total)})` : ''}</button>}
      {state?.phase === 'ready' && <button className="primary" disabled={disabled} onClick={() => { void run('install'); }}>
        {state.platform === 'darwin' ? '다운로드 파일·교체 안내' : '종료하고 설치'}</button>}
      {active && state?.phase !== 'installing' && <button onClick={() => { void run('cancel'); }}>취소</button>}
    </div>
    <div className="update-manual">
      <strong>수동 설치</strong><p>최신 설치 파일을 직접 받아 설치할 수도 있습니다.</p>
      {window.wordDesktop?.updateAction
        ? <button onClick={() => { void run('release-page'); }}>다운로드 페이지 열기</button>
        : <a href={releasePage} target="_blank" rel="noreferrer">다운로드 페이지 열기</a>}
    </div>
    <p className="update-help">창을 닫아도 다운로드는 계속되며, 하단 상태 표시줄에서 진행 상황을 다시 볼 수 있습니다.</p>
  </dialog>;
}
