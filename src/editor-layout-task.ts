/** Export measurements use their own detached document. Reflowing the live
 * editor between every yielded print block can starve that measurement,
 * particularly when floating artwork and table continuations trigger each
 * other. Resume both editor planners after the last measurement finishes. */
export const editorLayoutResume = 'studio-editor-layout-resume';
let measurements = 0;
export const editorLayoutSuspended = () => measurements > 0;
export async function withoutEditorReflow<T>(measure: () => Promise<T>): Promise<T> {
  measurements++;
  try { return await measure(); }
  finally {
    if (--measurements === 0) document.dispatchEvent(new Event(editorLayoutResume));
  }
}
