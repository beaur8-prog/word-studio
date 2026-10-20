import { useEffect, useState } from 'react';
import { Hourglass } from 'lucide-react';
import type { HwpOperation } from './progress';

export function HwpProgressPanel({ operation }: { operation: HwpOperation }) {
  const { progress, startedAt } = operation;
  const [now, setNow] = useState(Date.now);
  useEffect(() => {
    setNow(Date.now());
    const timer = setInterval(() => setNow(Date.now()), 1000);
    return () => clearInterval(timer);
  }, [startedAt]);
  const seconds = Math.max(0, Math.floor((now - startedAt) / 1000));
  const elapsed = seconds >= 60 ? `${Math.floor(seconds / 60)}분 ${seconds % 60}초` : `${seconds}초`;
  const known = Number.isFinite(progress.completed) && Number.isFinite(progress.total) && progress.total! > 0;
  const completed = known ? Math.max(0, Math.min(progress.completed!, progress.total!)) : 0;
  return <section className="hwp-progress-panel" role="region" aria-label="HWP 처리 진행" aria-busy="true">
    <div className="hwp-progress-heading">
      {!known && <Hourglass className="hwp-hourglass" data-testid="hwp-hourglass" size={20} aria-hidden="true" />}
      <strong role="status">{progress.label}</strong>
      <small aria-label="경과 시간" aria-live="off">경과 {elapsed}</small>
    </div>
    {known ? <>
      <progress aria-label={progress.label} max={progress.total} value={completed} />
      <span className="hwp-progress-detail">{completed.toLocaleString()} / {progress.total!.toLocaleString()}{progress.unit ?? '개'} · {Math.floor(completed / progress.total! * 100)}% <small>(현재 단계)</small></span>
    </> : <span className="hwp-progress-detail">진행률을 계산할 수 없는 단계입니다. 잠시 기다려 주세요.</span>}
  </section>;
}
