import { hwpRelatedTableLayoutAttribute, parseHwpRelatedTableLayout } from './hwp-related-table-layout';
import { hwpNestedTableLayoutAttribute, parseHwpNestedTableLayout } from './hwp-nested-table-layout';
import { parseHwpTableContinuation } from './hwp-table-continuation';
import { Extension } from '@tiptap/core';
import { TableView } from '@tiptap/extension-table';
import type { Node } from '@tiptap/pm/model';
import { cellTablePosition } from './cell-table-position';
import { tableAnchor } from './table-anchor';
import { hwpInlineTableTop, hwpInlineTableTopAttribute, hwpKeptTablePage, hwpTableTopGap } from './table-anchor-data';
import { paragraphTextAlignment } from './paragraph-text-alignment';
import { parseWordParagraphBottomBorder } from './word-paragraph-border';
import { paragraphBorderBoxAttribute, parseParagraphBorderBox } from './paragraph-border-box';
import { parseWordReflection, wordParagraphReflectionAttribute } from './word-reflection';
import { hwpSourceSpacesAttribute, hwpSourceWrapAttribute, parseHwpSourceWrap } from './hwp-text-advances';
import { characterScaleProperty, parseCharacterScale } from './character-scale';

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
    const value = css(property, attrs[key]);
    if (!value) return {};
    // The native HWP paragraph strut can carry its Korean and Latin faces
    // together. A missing pair still needs the established Korean serif
    // fallback without changing the names of either authored script.
    const nativeSerifPair = key === 'paragraphFont' && attrs.wordLineRule == null &&
      typeof attrs.lineSpacing === 'string' && attrs.lineSpacing.endsWith('px') &&
      /^(?:휴먼명조|"휴먼명조"),\s*(?:Palatino Linotype|"Palatino Linotype")$/i.test(value);
    return { style: `${property}:${value}`, ...(nativeSerifPair && { 'data-hwp-korean-serif-strut': 'true' }) };
  },
});
export const ImportedFormat = Extension.create({
  name: 'importedFormat',
  addGlobalAttributes() {
    return [
      { types: ['listItem'], attributes: Object.fromEntries(['id','level','format','pattern','label','suffix','font','start'].map(key => [`wordList${key[0].toUpperCase()}${key.slice(1)}`, {
        default: null, parseHTML: (el: HTMLElement) => el.getAttribute(`data-list-${key}`),
        renderHTML: (a: Record<string, unknown>) => a[`wordList${key[0].toUpperCase()}${key.slice(1)}`] != null ? { [`data-list-${key}`]: a[`wordList${key[0].toUpperCase()}${key.slice(1)}`] } : {},
      }])) },
      { types: ['paragraph', 'heading'], attributes: {
        ...Object.fromEntries(Object.entries(properties).map(([key, property]) => [key, attribute(property)])),
        ...Object.fromEntries([['wordAutoSpaceDE', 'data-word-auto-space-de'], ['wordAutoSpaceDN', 'data-word-auto-space-dn'], ['wordContextualSpacing', 'data-word-contextual-spacing'], ['wordKeepNext', 'data-word-keep-next'], ['wordKeepLines', 'data-word-keep-lines']].map(([key, name]) => [key, {
          default: null,
          parseHTML: (el: HTMLElement) => { const value = el.getAttribute(name); return value === 'true' ? true : value === 'false' ? false : null; },
          renderHTML: (a: Record<string, unknown>) => typeof a[key] === 'boolean' ? { [name]: String(a[key]) } : {},
        }])),
        hwpSourceSpaces: { default: null,
          parseHTML: (el: HTMLElement) => el.getAttribute(hwpSourceSpacesAttribute) === 'true' || null,
          renderHTML: (a: Record<string, unknown>) => a.hwpSourceSpaces === true ? { [hwpSourceSpacesAttribute]: 'true' } : {},
        },
        hwpSourceWrap: { default: null,
          parseHTML: (el: HTMLElement) => {
            const wrap = parseHwpSourceWrap(el.getAttribute(hwpSourceWrapAttribute));
            return wrap ? JSON.stringify(wrap) : null;
          },
          renderHTML: (a: Record<string, unknown>) => parseHwpSourceWrap(a.hwpSourceWrap)
            ? { [hwpSourceWrapAttribute]: a.hwpSourceWrap } : {},
        },
        wordTextAlignment: {
          default: null, parseHTML: (el: HTMLElement) => paragraphTextAlignment(el.dataset.wordTextAlignment),
          renderHTML: (a: Record<string, unknown>) => paragraphTextAlignment(a.wordTextAlignment)
            ? { 'data-word-text-alignment': a.wordTextAlignment } : {},
        },
        // An outline level supplies TOC/navigation metadata. It is separate
        // from the visible heading tag and never implies keep-with-next.
        wordOutlineLevel: {
          default: null,
          parseHTML: (el: HTMLElement) => /^[0-9]$/.test(el.dataset.wordOutlineLevel ?? '') ? Number(el.dataset.wordOutlineLevel) : null,
          renderHTML: (a: Record<string, unknown>) => Number.isInteger(a.wordOutlineLevel) && Number(a.wordOutlineLevel) >= 0 && Number(a.wordOutlineLevel) <= 9
            ? { 'data-word-outline-level': String(a.wordOutlineLevel) } : {},
        },
        wordParagraphStyle: {
          default: null, parseHTML: (el: HTMLElement) => el.getAttribute('data-word-paragraph-style'),
          renderHTML: (a: Record<string, unknown>) => typeof a.wordParagraphStyle === 'string' && a.wordParagraphStyle
            ? { 'data-word-paragraph-style': a.wordParagraphStyle } : {},
        },
        wordParagraphReflection: {
          default: null,
          parseHTML: (element: HTMLElement) => {
            const effect = parseWordReflection(element.getAttribute(wordParagraphReflectionAttribute));
            return effect ? JSON.stringify(effect) : null;
          },
          renderHTML: (attrs: Record<string, unknown>) => {
            const effect = parseWordReflection(attrs.wordParagraphReflection);
            return effect ? { [wordParagraphReflectionAttribute]: JSON.stringify(effect) } : {};
          },
        },
        paragraphBorderBox: {
          default: null,
          parseHTML: (element: HTMLElement) => {
            const box = parseParagraphBorderBox(element.getAttribute(paragraphBorderBoxAttribute));
            return box ? JSON.stringify(box) : null;
          },
          renderHTML: (attrs: Record<string, unknown>) => {
            const box = parseParagraphBorderBox(attrs.paragraphBorderBox);
            return box ? { [paragraphBorderBoxAttribute]: JSON.stringify(box) } : {};
          },
        },
        wordParagraphBottomBorder: {
          default: null,
          parseHTML: (el: HTMLElement) => {
            const border = parseWordParagraphBottomBorder(el.dataset.wordParagraphBottomBorder);
            return border ? JSON.stringify(border) : null;
          },
          renderHTML: (attrs: Record<string, unknown>) => {
            const border = parseWordParagraphBottomBorder(attrs.wordParagraphBottomBorder);
            return border ? { 'data-word-paragraph-bottom-border': JSON.stringify(border) } : {};
          },
        },
        wordParagraphStyleName: {
          default: null, parseHTML: (el: HTMLElement) => el.getAttribute('data-word-paragraph-style-name'),
          renderHTML: (a: Record<string, unknown>) => typeof a.wordParagraphStyleName === 'string' && a.wordParagraphStyleName
            ? { 'data-word-paragraph-style-name': a.wordParagraphStyleName } : {},
        },
        // Internal typing context, derived from the actual story tag. One
        // Word style (Title) can be a body heading and an ordinary box paragraph.
        wordParagraphStyleKind: {
          default: null, rendered: false,
          parseHTML: (el: HTMLElement) => el.hasAttribute('data-word-paragraph-style')
            ? /^H[1-6]$/.test(el.tagName) ? `heading${el.tagName.slice(1)}` : 'paragraph' : null,
        },
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
      characterScale: { default: null,
        parseHTML: (el: HTMLElement) => parseCharacterScale(el.style.getPropertyValue(characterScaleProperty)),
        renderHTML: (attrs: Record<string, unknown>) => {
          const scale = parseCharacterScale(attrs.characterScale);
          return scale && scale !== 100 ? { style: `${characterScaleProperty}:${scale}` } : {};
        },
      },
      smallCaps: { default: null, parseHTML: (el: HTMLElement) => el.style.fontVariantCaps === 'small-caps' ? true : null,
        renderHTML: (a: Record<string, unknown>) => a.smallCaps ? { style: 'font-variant-caps:small-caps' } : {} },
      allCaps: { default: null, parseHTML: (el: HTMLElement) => el.style.textTransform === 'uppercase' ? true : null,
        renderHTML: (a: Record<string, unknown>) => a.allCaps ? { style: 'text-transform:uppercase' } : {} },
      decorationStyle: { default: null, parseHTML: (el: HTMLElement) => el.style.textDecorationStyle || null,
        renderHTML: (a: Record<string, unknown>) => ['solid','double','dotted','dashed','wavy'].includes(String(a.decorationStyle)) ? { style: `text-decoration-style:${a.decorationStyle}` } : {} },
      } },
      { types: ['table'], attributes: {
        wordTableWidthAuto: { default: null,
          parseHTML: (el: HTMLElement) => el.dataset.wordTableWidthAuto === 'true' ? true : null,
          renderHTML: (a: Record<string, unknown>) => a.wordTableWidthAuto === true ? { 'data-word-table-width-auto': 'true' } : {} },
        wordTableLayout: { default: null,
          parseHTML: (el: HTMLElement) => ['fixed', 'autofit'].includes(el.dataset.wordTableLayout || '') ? el.dataset.wordTableLayout : null,
          renderHTML: (a: Record<string, unknown>) => ['fixed', 'autofit'].includes(String(a.wordTableLayout)) ? { 'data-word-table-layout': a.wordTableLayout } : {} },
        cellTablePosition: { default: null,
          parseHTML: (el: HTMLElement) => cellTablePosition(el.dataset.cellTablePosition) ? el.dataset.cellTablePosition : null,
          renderHTML: (a: Record<string, unknown>) => cellTablePosition(a.cellTablePosition) ? { 'data-cell-table-position': a.cellTablePosition } : {} },
        tableAnchor: { default: null, parseHTML: (el: HTMLElement) => { const a = tableAnchor(el.dataset.tableAnchor); return a ? JSON.stringify(a) : null; },
          renderHTML: (a: Record<string, unknown>) => tableAnchor(a.tableAnchor) ? { 'data-table-anchor': a.tableAnchor } : {} },
        hwpKeptTablePage: { default: null,
          parseHTML: (el: HTMLElement) => hwpKeptTablePage(el.dataset.hwpKeptTablePage) ?? null,
          renderHTML: (a: Record<string, unknown>) => hwpKeptTablePage(a.hwpKeptTablePage) ? { 'data-hwp-kept-table-page': a.hwpKeptTablePage } : {} },
        hwpInlineTableTop: { default: null,
          parseHTML: (el: HTMLElement) => hwpInlineTableTop(el.getAttribute(hwpInlineTableTopAttribute)) ?? null,
          renderHTML: (a: Record<string, unknown>) => {
            const top = hwpInlineTableTop(a.hwpInlineTableTop);
            return top === undefined ? {} : { [hwpInlineTableTopAttribute]: String(top) };
          } },
        hwpNestedTableLayout: { default: null,
          parseHTML: (el: HTMLElement) => {
            const value = el.getAttribute(hwpNestedTableLayoutAttribute);
            return parseHwpNestedTableLayout(value) ? value : null;
          },
          renderHTML: (a: Record<string, unknown>) => parseHwpNestedTableLayout(a.hwpNestedTableLayout)
            ? { [hwpNestedTableLayoutAttribute]: a.hwpNestedTableLayout } : {} },
        hwpRelatedTableLayout: { default: null,
          parseHTML: (el: HTMLElement) => {
            const value = el.getAttribute(hwpRelatedTableLayoutAttribute);
            return parseHwpRelatedTableLayout(value) ? value : null;
          },
          renderHTML: (a: Record<string, unknown>) => parseHwpRelatedTableLayout(a.hwpRelatedTableLayout)
            ? { [hwpRelatedTableLayoutAttribute]: a.hwpRelatedTableLayout } : {} },
        hwpTableContinuation: { default: null,
          parseHTML: (el: HTMLElement) => parseHwpTableContinuation(el.dataset.hwpTableContinuation) ? el.dataset.hwpTableContinuation : null,
          renderHTML: (a: Record<string, unknown>) => parseHwpTableContinuation(a.hwpTableContinuation)
            ? { 'data-hwp-table-continuation': a.hwpTableContinuation } : {} },
        hwpCenteredBorders: { default: null,
          parseHTML: (el: HTMLElement) => el.dataset.hwpCenteredBorders === 'inline' || el.dataset.hwpCenteredBorders === 'continuation'
            ? el.dataset.hwpCenteredBorders : el.dataset.hwpCenteredBorders === 'true' || null,
          renderHTML: (a: Record<string, unknown>) => a.hwpCenteredBorders === true || a.hwpCenteredBorders === 'inline' || a.hwpCenteredBorders === 'continuation'
            ? { 'data-hwp-centered-borders': String(a.hwpCenteredBorders) } : {} },
        hwpTableTopGap: { default: null,
          parseHTML: (el: HTMLElement) => hwpTableTopGap(el.dataset.hwpTableTopGap) ?? null,
          renderHTML: (a: Record<string, unknown>) => {
            const gap = !tableAnchor(a.tableAnchor) && hwpTableTopGap(a.hwpTableTopGap);
            return gap ? { 'data-hwp-table-top-gap': String(gap), style: `margin-top:${gap}px` } : {};
          },
        },
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
    if (this.node.attrs.wordTableWidthAuto === true) this.table.dataset.wordTableWidthAuto = 'true'; else delete this.table.dataset.wordTableWidthAuto;
    const layout = this.node.attrs.wordTableLayout;
    if (layout === 'fixed' || layout === 'autofit') this.table.dataset.wordTableLayout = layout; else delete this.table.dataset.wordTableLayout;
    const anchor = tableAnchor(this.node.attrs.tableAnchor);
    if (cellTablePosition(this.node.attrs.cellTablePosition)) this.table.dataset.cellTablePosition = this.node.attrs.cellTablePosition;
    else delete this.table.dataset.cellTablePosition;
    const firstRow = this.node.firstChild;
    const repeated = firstRow?.childCount && firstRow.content.content.every(cell => cell.type.name === 'tableHeader');
    const gap = repeated && !anchor && hwpTableTopGap(this.node.attrs.hwpTableTopGap);
    if (gap) this.table.dataset.hwpTableTopGap = String(gap); else delete this.table.dataset.hwpTableTopGap;
    const centered = this.node.attrs.hwpCenteredBorders;
    const inlineGrid = centered === 'inline' && !anchor;
    const repeatedGrid = repeated && (!anchor || anchor.reference === 'paragraph' && !anchor.wrap);
    if (inlineGrid || repeatedGrid && (centered === true || centered === 'inline' || centered === 'continuation'))
      this.table.dataset.hwpCenteredBorders = String(centered);
    else { delete this.table.dataset.hwpCenteredBorders; delete this.table.dataset.hwpCenteredBordersActive; }
    if (anchor) this.table.dataset.tableAnchor = JSON.stringify(anchor);
    else { delete this.table.dataset.tableAnchor; this.dom.style.marginTop = gap ? `${gap}px` : ''; }
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
