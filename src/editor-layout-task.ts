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
