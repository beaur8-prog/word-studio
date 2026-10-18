import { Node } from '@tiptap/core';
import maps from './symbol-font-maps.json';

const mappings = maps as Record<string, Record<string, string>>;
export function symbolFont(style: string) {
  const css = document.createElement('span').style; css.cssText = style;
  const family = css.fontFamily.split(',')[0].replace(/["']/g, '').trim();
  return mappings[family.toLowerCase()] ? family : '';
}
export function symbolElement(font: string, code: number) {
  const span = document.createElement('span');
  span.dataset.symbolFont = font; span.dataset.symbolCode = String(code);
  const mapped = mappings[font.toLowerCase()]?.[String(code & 255)];
  // Keep the original character even if the legacy font had no Unicode
  // equivalent. An explicit label is preferable to an indistinguishable box.
  span.dataset.symbolGlyph = mapped || '�';
  span.textContent = span.dataset.symbolGlyph; return span;
}
export const DocumentSymbol = Node.create({
  name: 'documentSymbol', group: 'inline', inline: true, atom: true,
  addAttributes: () => Object.fromEntries(['font','code','glyph'].map(name => [name, { default: '',
    parseHTML: (el: HTMLElement) => el.getAttribute(`data-symbol-${name}`),
    renderHTML: (attrs: Record<string, unknown>) => ({ [`data-symbol-${name}`]: attrs[name] }),
  }])),
  parseHTML: () => [{ tag: 'span[data-symbol-font][data-symbol-code]' }],
  renderHTML: ({ node, HTMLAttributes }) => ['span', { ...HTMLAttributes, class: 'document-symbol',
    title: `${node.attrs.font} U+${Number(node.attrs.code).toString(16).toUpperCase()}` }, node.attrs.glyph],
  renderText: ({ node }) => node.attrs.glyph,
});
