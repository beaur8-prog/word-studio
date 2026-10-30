import { Paragraph, type IParagraphOptions } from 'docx';
import { hwpCellLineNamespace } from './hwp-cell-line';
import { hwpBodyLineNamespace } from './hwp-body-line';

/** Standard spacing carries the current measured single/multi-line geometry.
 * Ignorable metadata retains the original live HWP spacing policy locally. */
export class HwpCellParagraph extends Paragraph {
  constructor(options: IParagraphOptions, private readonly policy?: string, private readonly bodyPolicy?: string,
    private readonly autoSpacing?: { de?: unknown; dn?: unknown }) { super(options); }
  prepForXml(context: Parameters<Paragraph['prepForXml']>[0]) {
    const result = super.prepForXml(context), children = result?.['w:p'];
    // docx 9.x emits wordWrap before tabs. CT_PPr requires tabs first;
    // readers accepting the wrong order do not make the package valid OOXML.
    let properties = Array.isArray(children) && children.find((entry: any) => Array.isArray(entry?.['w:pPr']))?.['w:pPr'];
    if (Array.isArray(properties)) {
      const wrap = properties.findIndex((entry: any) => entry && 'w:wordWrap' in entry);
      const tabs = properties.findIndex((entry: any) => entry && 'w:tabs' in entry);
      if (wrap >= 0 && tabs > wrap) properties.splice(wrap, 0, properties.splice(tabs, 1)[0]);
    }
    const automatic = [['w:autoSpaceDE', this.autoSpacing?.de], ['w:autoSpaceDN', this.autoSpacing?.dn]]
      .filter(([, value]) => typeof value === 'boolean');
    if (automatic.length && Array.isArray(children)) {
      if (!Array.isArray(properties)) { properties = []; children.splice(children[0]?._attr ? 1 : 0, 0, { 'w:pPr': properties }); }
      for (const [name] of automatic) { const at = properties.findIndex((p: any) => p && (name as string) in p); if (at >= 0) properties.splice(at, 1); }
      // CT_PPr places these after wordWrap and before bidi/spacing/indent.
      const later = new Set(['bidi','adjustRightInd','snapToGrid','spacing','ind','contextualSpacing','mirrorIndents','suppressOverlap','jc','textDirection','textAlignment','textboxTightWrap','outlineLvl','divId','cnfStyle','rPr','sectPr','pPrChange'].map(n => 'w:' + n));
      const at = properties.findIndex((p: any) => p && Object.keys(p).some(n => later.has(n)));
      properties.splice(at < 0 ? properties.length : at, 0, ...automatic.map(([name, value]) => ({ [name as string]: [{ _attr: { 'w:val': String(value) } }] })));
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
