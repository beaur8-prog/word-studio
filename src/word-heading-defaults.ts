import { Extension } from '@tiptap/core';
import { Plugin, PluginKey } from '@tiptap/pm/state';

const names = ['bold', 'italic', 'underline', 'strike'] as const;
type HeadingMarks = Record<typeof names[number], boolean>;
const key = new PluginKey('wordHeadingDefaults');
export function wordHeadingMarks(value: unknown): HeadingMarks | null {
  if (typeof value !== 'string' || value.length > 160) return null;
  try {
    const parsed = JSON.parse(value);
    return parsed && typeof parsed === 'object' && !Array.isArray(parsed) &&
      Object.keys(parsed).length === names.length && names.every(name => typeof parsed[name] === 'boolean') ? parsed : null;
  } catch { return null; }
}

/** Preserve typing defaults separately from existing run marks. Applying a bold
 * default as parent CSS would also make explicitly normal child runs bold. */
export const WordHeadingDefaults = Extension.create({
  name: 'wordHeadingDefaults',
  addGlobalAttributes: () => [{ types: ['heading'], attributes: {
    wordHeadingMarks: {
      default: null,
      parseHTML: el => {
        const value = el.getAttribute('data-word-heading-marks');
        return el.hasAttribute('data-word-line-rule') && wordHeadingMarks(value) ? value : null;
      },
      renderHTML: attrs => wordHeadingMarks(attrs.wordHeadingMarks) ? { 'data-word-heading-marks': attrs.wordHeadingMarks } : {},
    },
  } }],
  addProseMirrorPlugins: () => [new Plugin({
    key,
    appendTransaction(transactions, _old, state) {
      const { selection, storedMarks } = state, parent = selection.$from.parent;
      if (!selection.empty || parent.type.name !== 'heading' || parent.content.size) return null;
      const defaults = wordHeadingMarks(parent.attrs.wordHeadingMarks);
      if (!defaults) return null;
      if (storedMarks !== null && transactions.some(tr => tr.storedMarksSet && !tr.docChanged && !tr.getMeta(key))) {
        const chosen = { ...defaults };
        for (const name of names) chosen[name] = storedMarks.some(mark => mark.type.name === name);
        if (!names.some(name => chosen[name] !== defaults[name])) return null;
        // A user's format choice on the empty heading changes its paragraph
        // mark too, so it survives leaving the paragraph and saving the file.
        return state.tr.setNodeMarkup(selection.$from.before(), undefined,
          { ...parent.attrs, wordHeadingMarks: JSON.stringify(chosen) }).setStoredMarks(storedMarks).setMeta(key, true);
      }
      if (storedMarks !== null || !transactions.some(tr => tr.selectionSet || tr.docChanged)) return null;
      const marks = names.flatMap(name => defaults[name] && state.schema.marks[name] ? [state.schema.marks[name].create()] : []);
      // An explicit toolbar/keyboard choice sets storedMarks itself, including
      // an empty array. Do not turn a user's disabled bold/italic back on.
      return state.tr.setStoredMarks(marks).setMeta(key, true).setMeta('addToHistory', false);
    },
  })],
});
