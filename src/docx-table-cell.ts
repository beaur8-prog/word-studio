import { Table, TableCell, type ITableOptions, type ITableCellOptions } from 'docx';
import { paragraphTableNamespace, type TableAnchor } from './table-anchor';

/** Inline tables can continue across readers without clipping floating merged
 * cells. Ignorable per-table metadata preserves the editor's live reference. */
export class PreservedTable extends Table {
  constructor(options: ITableOptions, private readonly anchor?: TableAnchor) { super(options); }
  prepForXml(context: Parameters<Table['prepForXml']>[0]) {
    const result = super.prepForXml(context);
    const children = result?.['w:tbl'];
    if (this.anchor && Array.isArray(children)) {
      let attributes = children.find((entry: any) => entry && '_attr' in entry)?.['_attr'];
      if (!attributes) { attributes = {}; children.unshift({ _attr: attributes }); }
      Object.assign(attributes, { 'xmlns:ws': paragraphTableNamespace,
        'xmlns:mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
        'mc:Ignorable': 'ws', 'mc:PreserveAttributes': 'ws:anchor', 'ws:anchor': JSON.stringify(this.anchor) });
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
