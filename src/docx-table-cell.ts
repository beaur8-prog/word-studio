import { TableCell, type ITableCellOptions } from 'docx';

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
