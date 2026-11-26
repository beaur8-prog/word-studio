/** An output layout owns its detached paper while preparing or displaying
 * it. Live editor reflow can otherwise compete for the same renderer. */
export const editorLayoutResume = 'studio-editor-layout-resume';
let measurements = 0;
export const editorLayoutSuspended = () => measurements > 0;

export function suspendEditorReflow() {
  measurements++;
  let active = true;
  return () => {
    if (!active) return;
    active = false;
    if (--measurements === 0) document.dispatchEvent(new Event(editorLayoutResume));
  };
}

export async function withoutEditorReflow<T>(measure: () => Promise<T>): Promise<T> {
  const resume = suspendEditorReflow();
  try { return await measure(); }
  finally { resume(); }
}

export type LayoutCheckpoint = () => Promise<void> | undefined;

/** Yield between complete operations in a detached layout tree. */
export function layoutCheckpoint(assertCurrent: () => void): LayoutCheckpoint {
  let started = performance.now();
  return () => {
    assertCurrent();
    if (performance.now() - started < 32) return;
    return new Promise<void>(resolve => setTimeout(resolve, 0)).then(() => {
      assertCurrent();
      started = performance.now();
    });
  };
}
