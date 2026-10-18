import Image from '@tiptap/extension-image';
import { mergeAttributes } from '@tiptap/core';
import { imageFloating, floatingCss } from './image-floating';

// Retain the block node for existing .wordstudio files and inserted images.
export const BlockImage = Image.extend({
  parseHTML: () => [{ tag: 'img[src]:not([data-inline-image])' }],
});
export const InlineImage = Image.extend({
  name: 'inlineImage',
  addCommands: () => ({}),
  addInputRules: () => [],
  addAttributes() { return { ...this.parent?.(), floating: { default: null, parseHTML: el => imageFloating(el.getAttribute('data-image-floating')) ? el.getAttribute('data-image-floating') : null,
    renderHTML: attrs => attrs.floating ? { 'data-image-floating': attrs.floating } : {} } }; },
  parseHTML: () => [{ tag: 'img[data-inline-image][src]' }],
  renderHTML({ HTMLAttributes }) {
    const width = Number(HTMLAttributes.width), height = Number(HTMLAttributes.height);
    return ['img', mergeAttributes(HTMLAttributes, { 'data-inline-image': 'true',
      ...(width > 0 && height > 0 ? { style: `width:${width}px;aspect-ratio:${width}/${height};${floatingCss(HTMLAttributes['data-image-floating'], width)}` } : {}),
    })];
  },
}).configure({ inline: true, allowBase64: true });
