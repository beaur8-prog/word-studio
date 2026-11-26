import { TextRun, type IRunOptions } from 'docx';
import { hwpBoxedDigitBorder, hwpBoxedDigitNamespace } from './hwp-boxed-digit';

/** Preserve the digit as standard bordered text. The optional metadata keeps
 * the original square advance when reopening in Studio. */
export class HwpBoxedDigitRun extends TextRun {
  constructor(private readonly box: { size: number; digit: string }, options: IRunOptions) {
    const size = Math.round(box.size * .7);
    super({ ...options, text: box.digit, size, sizeComplexScript: size, border: hwpBoxedDigitBorder(box.size) });
  }

  prepForXml(context: Parameters<TextRun['prepForXml']>[0]) {
    const result = super.prepForXml(context);
    const children = result?.['w:r'];
    if (!Array.isArray(children)) return result;
    children.unshift({ _attr: {
      'xmlns:wsdigit': hwpBoxedDigitNamespace,
      'xmlns:mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
      'mc:Ignorable': 'wsdigit',
      'mc:PreserveAttributes': 'wsdigit:box',
      'wsdigit:box': String(this.box.size) + ':' + this.box.digit,
    } });
    return result;
  }
}
