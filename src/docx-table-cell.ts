import { Table, TableCell, type ITableOptions, type ITableCellOptions } from 'docx';
import { paragraphTableNamespace, type TableAnchor } from './table-anchor';

/** Inline tables can continue across readers without clipping floating merged
 * cells. Ignorable per-table metadata preserves the editor's live reference. */
export class PreservedTable extends Table {
  constructor(options: ITableOptions, private readonly anchor?: TableAnchor,
    private readonly overlayPosition?: { x: number; y: number; width: number }) { super(options); }
  prepForXml(context: Parameters<Table['prepForXml']>[0]) {
    const result = super.prepForXml(context);
    const children = result?.['w:tbl'];
    // docx creates a bare <w:p/> in each generated vMerge continuation.
    // Word can give that required placeholder Normal's body leading and space
    // after, expanding short source rows and pushing a following tall row
    // onto another page. Only compact the library's empty continuation;
    // authored empty paragraphs and the merged cell's real content remain.
    if (Array.isArray(children)) for (const row of children) {
      if (!Array.isArray(row?.['w:tr'])) continue;
      for (const cell of row['w:tr']) {
        const parts = cell?.['w:tc'];
        if (!Array.isArray(parts)) continue;
        const properties = parts.find((part: any) => Array.isArray(part?.['w:tcPr']))?.['w:tcPr'];
        const merge = properties?.find((part: any) => part && 'w:vMerge' in part)?.['w:vMerge'];
        if (merge?._attr?.['w:val'] !== 'continue') continue;
        const paragraphs = parts.filter((part: any) => part && 'w:p' in part);
        if (paragraphs.length !== 1 || !paragraphs[0]['w:p'] || Object.keys(paragraphs[0]['w:p']).length) continue;
        // End-of-cell markers otherwise retain the default font's minimum
        // even when the required paragraph has exact one-twip leading.
        if (Array.isArray(properties) && !properties.some((part: any) => part && 'w:hideMark' in part))
          properties.push({ 'w:hideMark': {} });
        paragraphs[0]['w:p'] = [{ 'w:pPr': [{ 'w:spacing': [{ _attr: {
          'w:before': '0', 'w:after': '0', 'w:line': '1', 'w:lineRule': 'exact',
        } }] }] }];
      }
    }
    if (this.anchor && Array.isArray(children)) {
      let attributes = children.find((entry: any) => entry && '_attr' in entry)?.['_attr'];
      if (!attributes) { attributes = {}; children.unshift({ _attr: attributes }); }
      Object.assign(attributes, { 'xmlns:ws': paragraphTableNamespace,
        'xmlns:mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
        'mc:Ignorable': 'ws', 'mc:PreserveAttributes': this.overlayPosition ? 'ws:anchor ws:overlayPosition' : 'ws:anchor',
        'ws:anchor': JSON.stringify(this.anchor), ...(this.overlayPosition ? { 'ws:overlayPosition': JSON.stringify(this.overlayPosition) } : {}) });
    }
    return result;
  }
}

/** docx 9.x has no noWrap option. Extend its public XML serializer rather than
 * silently passing an ignored option or changing the table's column widths. */
export class PreservedTableCell extends TableCell {
  constructor(options: ITableCellOptions, private readonly noWrap: boolean) { super(options); }
  prepForXml(context: Parameters<TableCell['prepForXml']>[0]) {
    const result = super.prepForXml(context);
    if (this.noWrap && result) {
      const children = result['w:tc'];
      if (Array.isArray(children)) {
        const props = children.find((entry: any) => entry && 'w:tcPr' in entry)?.['w:tcPr'];
        if (Array.isArray(props) && !props.some((entry: any) => entry && 'w:noWrap' in entry)) {
          const index = props.findIndex((entry: any) => ['w:tcMar', 'w:textDirection', 'w:vAlign'].some(key => entry && key in entry));
          props.splice(index < 0 ? props.length : index, 0, { 'w:noWrap': {} });
        }
      }
    }
    return result;
  }
}
