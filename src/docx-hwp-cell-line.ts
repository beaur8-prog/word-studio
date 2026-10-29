import { Paragraph, type IParagraphOptions } from 'docx';
import { hwpCellLineNamespace } from './hwp-cell-line';
import { hwpBodyLineNamespace } from './hwp-body-line';

/** Standard spacing carries the current measured single/multi-line geometry.
 * Ignorable metadata retains the original live HWP spacing policy locally. */
export class HwpCellParagraph extends Paragraph {
  constructor(options: IParagraphOptions, private readonly policy?: string, private readonly bodyPolicy?: string) { super(options); }
  prepForXml(context: Parameters<Paragraph['prepForXml']>[0]) {
    const result = super.prepForXml(context), children = result?.['w:p'];
    // docx 9.x emits wordWrap before tabs. CT_PPr requires tabs first;
    // readers accepting the wrong order do not make the package valid OOXML.
    const properties = Array.isArray(children) && children.find((entry: any) => Array.isArray(entry?.['w:pPr']))?.['w:pPr'];
    if (Array.isArray(properties)) {
      const wrap = properties.findIndex((entry: any) => entry && 'w:wordWrap' in entry);
      const tabs = properties.findIndex((entry: any) => entry && 'w:tabs' in entry);
      if (wrap >= 0 && tabs > wrap) properties.splice(wrap, 0, properties.splice(tabs, 1)[0]);
    }
    if ((this.policy || this.bodyPolicy) && Array.isArray(children)) {
      let attributes = children.find((entry: any) => entry && '_attr' in entry)?.['_attr'];
      if (!attributes) { attributes = {}; children.unshift({ _attr: attributes }); }
      Object.assign(attributes, {
        'xmlns:mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
        'mc:Ignorable': [this.policy && 'wsh', this.bodyPolicy && 'wsb'].filter(Boolean).join(' '),
        'mc:PreserveAttributes': [this.policy && 'wsh:cellLine', this.bodyPolicy && 'wsb:bodyLine'].filter(Boolean).join(' '),
        ...(this.policy ? { 'xmlns:wsh': hwpCellLineNamespace, 'wsh:cellLine': this.policy } : {}),
        ...(this.bodyPolicy ? { 'xmlns:wsb': hwpBodyLineNamespace, 'wsb:bodyLine': this.bodyPolicy } : {}),
      });
    }
    return result;
  }
}
