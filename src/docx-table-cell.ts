import { hwpCenteredCellLabelsNamespace } from './hwp-centered-cell-labels';
import { Table, TableCell, TableRow, type ITableOptions, type ITableCellOptions, type ITableRowOptions } from 'docx';
import { paragraphTableNamespace, type TableAnchor } from './table-anchor';

/** Inline tables can continue across readers without clipping floating merged
 * cells. Ignorable per-table metadata preserves the editor's live reference. */
export class PreservedTable extends Table {
  constructor(options: ITableOptions, private readonly anchor?: TableAnchor,
    private readonly overlayPosition?: { x: number; y: number; width: number },
    private readonly topGap?: number, private readonly centeredBorders: boolean | 'inline' | 'continuation' = false, private readonly keptTablePage?: number,
    private readonly continuation?: string, private readonly nestedLayout?: string, private readonly relatedLayout?: string) { super(options); }
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
    if ((this.anchor || this.topGap || this.centeredBorders || this.keptTablePage || this.continuation || this.nestedLayout || this.relatedLayout) && Array.isArray(children)) {
      let attributes = children.find((entry: any) => entry && '_attr' in entry)?.['_attr'];
      if (!attributes) { attributes = {}; children.unshift({ _attr: attributes }); }
      const layout = { ...(this.anchor ? { 'ws:anchor': JSON.stringify(this.anchor) } : {}),
        ...(this.overlayPosition ? { 'ws:overlayPosition': JSON.stringify(this.overlayPosition) } : {}),
        ...(this.topGap ? { 'ws:topGap': String(this.topGap) } : {}),
        ...(this.centeredBorders ? { 'ws:centeredBorders': this.centeredBorders === true ? '1' : this.centeredBorders } : {}),
        ...(this.keptTablePage ? { 'ws:keptTablePage': String(this.keptTablePage) } : {}),
        ...(this.continuation ? { 'ws:continuation': this.continuation } : {}),
        ...(this.nestedLayout ? { 'ws:nestedLayout': this.nestedLayout } : {}),
        ...(this.relatedLayout ? { 'ws:relatedLayout': this.relatedLayout } : {}) };
      Object.assign(attributes, { 'xmlns:ws': paragraphTableNamespace,
        'xmlns:mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
        'mc:Ignorable': 'ws', 'mc:PreserveAttributes': Object.keys(layout).join(' '), ...layout });
    }
    return result;
  }
}

export const hwpGridRowHeightNamespace = 'urn:wordstudio:hwp-grid-row-height:1';

/** A native grid segment can be shorter than the surrounding merged cells'
 * insets. Retain that minimum separately when Word's positive row value
 * cannot encode it. The standard row and cell properties still stand alone. */
export class PreservedTableRow extends TableRow {
  constructor(options: ITableRowOptions, private readonly gridHeight?: number, private readonly continued = false,
    private readonly reserve?: string, private readonly intact = false, private readonly intactTopRow?: string) { super(options); }
  prepForXml(context: Parameters<TableRow['prepForXml']>[0]) {
    const result = super.prepForXml(context), children = result?.['w:tr'];
    if ((this.gridHeight === undefined && !this.continued && !this.intact && !this.reserve && !this.intactTopRow) || !Array.isArray(children)) return result;
    let attrs = children.find((entry: any) => entry && '_attr' in entry)?._attr;
    if (!attrs) { attrs = {}; children.unshift({ _attr: attrs }); }
    const flags: Record<string, string> = {};
    if (this.gridHeight !== undefined) flags['wsr:height'] = String(this.gridHeight);
    if (this.continued) flags['wsr:continued'] = '1';
    if (this.reserve) flags['wsr:reserve'] = this.reserve;
    if (this.intact) flags['wsr:intact'] = '1';
    if (this.intactTopRow) flags['wsr:intactTop'] = this.intactTopRow;
    Object.assign(attrs, {
      'xmlns:mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
      'xmlns:wsr': hwpGridRowHeightNamespace,
      'mc:Ignorable': 'wsr',
      'mc:PreserveAttributes': Object.keys(flags).join(' '),
      ...flags,
    });
    return result;
  }
}

/** docx 9.x has no noWrap option. Extend its public XML serializer rather than
 * silently passing an ignored option or changing the table's column widths. */
export class PreservedTableCell extends TableCell {
  constructor(options: ITableCellOptions, private readonly noWrap: boolean, private readonly centeredLabels?: string) { super(options); }
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
    const children = result?.['w:tc'];
    if (this.centeredLabels && Array.isArray(children)) {
      let attrs = children.find((entry: any) => entry && '_attr' in entry)?._attr;
      if (!attrs) { attrs = {}; children.unshift({ _attr: attrs }); }
      Object.assign(attrs, {
        'xmlns:mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
        'xmlns:wcl': hwpCenteredCellLabelsNamespace,
        'mc:Ignorable': 'wcl',
        'mc:PreserveAttributes': 'wcl:leading',
        'wcl:leading': this.centeredLabels,
      });
    }
    return result;
  }
}
