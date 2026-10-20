/** Counts describe the current stage, never an estimate of the whole conversion. */
export interface HwpProgress { label: string; completed?: number; total?: number; unit?: string }
export type HwpProgressReporter = (progress: HwpProgress) => void;
export interface HwpOperation { progress: HwpProgress; startedAt: number }

// Paint the new stage before synchronous DOM/engine preparation starts. A timer
// also works when a native window is hidden and requestAnimationFrame is paused.
export const paintHwpProgress = () => new Promise<void>(resolve => setTimeout(resolve, 24));
