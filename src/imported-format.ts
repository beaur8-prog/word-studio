import { Extension } from '@tiptap/core';
import { TableView } from '@tiptap/extension-table';
import type { Node } from '@tiptap/pm/model';
import { tableAnchor } from './table-anchor';

const properties = {
  spaceBefore: 'margin-top', spaceAfter: 'margin-bottom', rightIndent: 'margin-right', firstLineIndent: 'text-indent',
  paragraphFont: 'font-family', paragraphSize: 'font-size', paragraphColor: 'color',
  wordBreak: 'word-break',
} as const;
function css(property: string, value: unknown) {
  if (typeof value !== 'string' || /[;{}\\]|url\s*\(|var\s*\(/i.test(value)) return '';
  const style = document.createElement('span').style; style.setProperty(property, value);
  return style.getPropertyValue(property);
}
const attribute = (property: string) => ({
  default: null,
  parseHTML: (el: HTMLElement) => el.style.getPropertyValue(property) || null,
  renderHTML: (attrs: Record<string, unknown>) => {
    const key = Object.keys(properties).find(key => properties[key as keyof typeof properties] === property)!;
    const value = css(property, attrs[key]); return value ? { style: `${property}:${value}` } : {};
  },
});
export const ImportedFormat = Extension.create({
  name: 'importedFormat',
  addGlobalAttributes() {
    return [
      { types: ['listItem'], attributes: Object.fromEntries(['id','level','format','pattern','label'].map(key => [`wordList${key[0].toUpperCase()}${key.slice(1)}`, {
        default: null, parseHTML: (el: HTMLElement) => el.getAttribute(`data-list-${key}`),
        renderHTML: (a: Record<string, unknown>) => a[`wordList${key[0].toUpperCase()}${key.slice(1)}`] != null ? { [`data-list-${key}`]: a[`wordList${key[0].toUpperCase()}${key.slice(1)}`] } : {},
      }])) },
      { types: ['paragraph', 'heading'], attributes: {
        ...Object.fromEntries(Object.entries(properties).map(([key, property]) => [key, attribute(property)])),
        wordWidowControl: {
          default: null,
          parseHTML: (el: HTMLElement) => {
            const count = Number(el.style.widows || el.style.orphans);
            return count > 0 ? count > 1 : null;
          },
          renderHTML: (a: Record<string, unknown>) => typeof a.wordWidowControl === 'boolean'
            ? { style: `widows:${a.wordWidowControl ? 2 : 1};orphans:${a.wordWidowControl ? 2 : 1}` } : {},
        },
      } },
      { types: ['textStyle'], attributes: { letterSpacing: {
        default: null, parseHTML: (el: HTMLElement) => el.style.letterSpacing || null,
        renderHTML: (a: Record<string, unknown>) => css('letter-spacing', a.letterSpacing) ? { style: `letter-spacing:${css('letter-spacing', a.letterSpacing)}` } : {},
      },
      smallCaps: { default: null, parseHTML: (el: HTMLElement) => el.style.fontVariantCaps === 'small-caps' ? true : null,
        renderHTML: (a: Record<string, unknown>) => a.smallCaps ? { style: 'font-variant-caps:small-caps' } : {} },
      allCaps: { default: null, parseHTML: (el: HTMLElement) => el.style.textTransform === 'uppercase' ? true : null,
        renderHTML: (a: Record<string, unknown>) => a.allCaps ? { style: 'text-transform:uppercase' } : {} },
      decorationStyle: { default: null, parseHTML: (el: HTMLElement) => el.style.textDecorationStyle || null,
        renderHTML: (a: Record<string, unknown>) => ['solid','double','dotted','dashed','wavy'].includes(String(a.decorationStyle)) ? { style: `text-decoration-style:${a.decorationStyle}` } : {} },
      } },
      { types: ['table'], attributes: {
        tableAnchor: { default: null, parseHTML: (el: HTMLElement) => { const a = tableAnchor(el.dataset.tableAnchor); return a ? JSON.stringify(a) : null; },
          renderHTML: (a: Record<string, unknown>) => tableAnchor(a.tableAnchor) ? { 'data-table-anchor': a.tableAnchor } : {} },
        tableAlignment: { default: null, parseHTML: (el: HTMLElement) => el.style.marginLeft === 'auto' ? el.style.marginRight === 'auto' ? 'center' : 'right' : null,
          renderHTML: (a: Record<string, unknown>) => !tableAnchor(a.tableAnchor) && ['center', 'right'].includes(String(a.tableAlignment)) ? { style: `margin-left:auto;margin-right:${a.tableAlignment === 'center' ? 'auto' : '0'}` } : {} },
        tableIndent: { default: null, parseHTML: (el: HTMLElement) => el.style.marginLeft !== 'auto' ? el.style.marginLeft || null : null,
          renderHTML: (a: Record<string, unknown>) => a.tableIndent ? { style: `margin-left:${css('margin-left', a.tableIndent)}` } : {} },
        tableSpacing: { default: null, parseHTML: (el: HTMLElement) => el.style.marginBottom || null,
          renderHTML: (a: Record<string, unknown>) => a.tableSpacing ? { style: `margin-bottom:${css('margin-bottom', a.tableSpacing)}` } : {} },
      } },
    ];
  },
});

/** The resize plugin constructs its own NodeView without rendered attributes. */
export class StudioTableView extends TableView {
  constructor(...args: ConstructorParameters<typeof TableView>) { super(...args); this.applyFormat(); }
  applyFormat() {
    const anchor = tableAnchor(this.node.attrs.tableAnchor);
    if (anchor) this.table.dataset.tableAnchor = JSON.stringify(anchor); else { delete this.table.dataset.tableAnchor; this.dom.style.removeProperty('margin-top'); }
    this.dom.style.marginBottom = css('margin-bottom', this.node.attrs.tableSpacing);
    this.dom.style.marginLeft = css('margin-left', this.node.attrs.tableIndent);
    this.table.style.marginLeft = !anchor && ['center', 'right'].includes(this.node.attrs.tableAlignment) ? 'auto' : '0';
    this.table.style.marginRight = !anchor && this.node.attrs.tableAlignment === 'center' ? 'auto' : '0';
  }
  update(node: Node) { const updated = super.update(node); if (updated) this.applyFormat(); return updated; }
  ignoreMutation(mutation: Parameters<TableView['ignoreMutation']>[0]) {
    return mutation.type === 'attributes' && (mutation.target === this.dom || mutation.target === this.table) || super.ignoreMutation(mutation);
  }
}
