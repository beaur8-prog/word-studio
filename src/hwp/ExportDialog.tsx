import { useEffect, useRef, useState } from 'react';
import { HWP_CONVERSION_NOTICE } from './conversion';

export interface HwpExportFile { blob: Blob; name: string; extension: 'hwp' | 'hwpx'; warnings: string[] }
export function HwpExportDialog({ file, onClose, onSave }: { file: HwpExportFile; onClose(): void; onSave(): Promise<boolean> }) {
  const ref = useRef<HTMLDialogElement>(null);
  const [saving, setSaving] = useState(false), [error, setError] = useState('');
  useEffect(() => { ref.current?.showModal(); }, []);
  return <dialog ref={ref} className="hwp-conversion-dialog" aria-label="한글 문서로 내보내기" onCancel={event => { event.preventDefault(); if (!saving) onClose(); }}>
    <h2>{file.extension.toUpperCase()}로 저장</h2><p>{HWP_CONVERSION_NOTICE}</p>
    {!!file.warnings.length && <ul>{file.warnings.map((warning, i) => <li key={i}>{warning}</li>)}</ul>}
    <p>저장할 파일: <strong>{file.name}</strong> · 변환된 파일을 다시 읽는 검사를 통과했습니다.</p>
    {error && <p role="alert">{error}</p>}
    <button disabled={saving} onClick={() => void (async () => {
      setSaving(true); setError('');
      try { if (await onSave()) onClose(); }
      catch (reason) { setError(`저장 실패: ${reason instanceof Error ? reason.message : String(reason)}`); }
      finally { setSaving(false); }
    })()}>{saving ? '저장 중…' : `${file.extension.toUpperCase()} 파일 저장`}</button>
    <button disabled={saving} onClick={onClose}>취소</button>
  </dialog>;
}
