import { Extension, type Editor } from '@tiptap/core';
import { Plugin, PluginKey, type Transaction } from '@tiptap/pm/state';
import { closeHistory } from '@tiptap/pm/history';
import { Step, StepResult } from '@tiptap/pm/transform';
import type { Node, Schema } from '@tiptap/pm/model';
import type { StudioDocument } from './types';

type Metadata = Pick<StudioDocument, 'page' | 'comments'>;
export const documentMetadata = new PluginKey<Metadata | undefined>('wordStudioMetadata');

// Page settings and comment records live outside the body schema. A reversible
// step puts them in the SAME history as typing/marks, without duplicating body
// text or image data in a second undo stack or changing the saved file format.
class MetadataStep extends Step {
  constructor(readonly before: Metadata, readonly after: Metadata) { super(); }
  apply(doc: Node) { return StepResult.ok(doc); }
  invert() { return new MetadataStep(this.after, this.before); }
  map() { return this; } // Metadata has no document positions to map.
  toJSON() { return { stepType: 'wordStudioMetadata', before: this.before, after: this.after }; }
  static fromJSON(_schema: Schema, json: { before: Metadata; after: Metadata }) {
    return new MetadataStep(json.before, json.after);
  }
}
Step.jsonID('wordStudioMetadata', MetadataStep);

export const DocumentHistory = Extension.create({
  name: 'documentHistory',
  addProseMirrorPlugins: () => [new Plugin<Metadata | undefined>({
    key: documentMetadata,
    state: {
      init: () => undefined,
      apply(tr, previous) {
        let value = (tr.getMeta(documentMetadata) as Metadata | undefined) ?? previous;
        for (const step of tr.steps) if (step instanceof MetadataStep) value = step.after;
        return value;
      },
    },
  })],
});

/** A discrete UI command must not merge with typing immediately before/after it. */
export function dispatchHistoryEvent(editor: Editor, tr: Transaction) {
  editor.view.dispatch(closeHistory(tr));
  editor.view.dispatch(closeHistory(editor.state.tr));
}

export function changeDocumentMetadata(editor: Editor, before: Metadata, after: Metadata, tr = editor.state.tr) {
  if (JSON.stringify(before) === JSON.stringify(after) && !tr.docChanged) return;
  dispatchHistoryEvent(editor, tr.step(new MetadataStep(before, after)));
}
