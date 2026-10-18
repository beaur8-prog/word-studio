import { Extension } from '@tiptap/core';
import { TableView } from '@tiptap/extension-table';
import type { Node } from '@tiptap/pm/model';

const properties = {
  spaceBefore: 'margin-top', spaceAfter: 'margin-bottom', rightIndent: 'margin-right', firstLineIndent: 'text-indent',
  paragraphFont: 'font-family', paragraphSize: 'font-size', paragraphColor: 'color',
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
      { types: ['paragraph', 'heading'], attributes: Object.fromEntries(Object.entries(properties).map(([key, property]) => [key, attribute(property)])) },
      { types: ['textStyle'], attributes: { letterSpacing: {
        default: null, parseHTML: (el: HTMLElement) => el.style.letterSpacing || null,
        renderHTML: (a: Record<string, unknown>) => css('letter-spacing', a.letterSpacing) ? { style: `letter-spacing:${css('letter-spacing', a.letterSpacing)}` } : {},
      } } },
      { types: ['table'], attributes: {
        tableIndent: { default: null, parseHTML: (el: HTMLElement) => el.style.marginLeft || null,
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
    this.dom.style.marginBottom = css('margin-bottom', this.node.attrs.tableSpacing);
    this.dom.style.marginLeft = css('margin-left', this.node.attrs.tableIndent);
  }
  update(node: Node) { const updated = super.update(node); if (updated) this.applyFormat(); return updated; }
}
