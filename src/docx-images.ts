import { W, all } from './docx-format';

const WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const V = 'urn:schemas-microsoft-com:vml';

/** Mammoth lifts VML pictures out of their paragraph and drops their geometry.
 * Normalize image-only VML into an inline DrawingML picture before conversion;
 * relationship IDs, order and accessible descriptions stay intact. */
export function normalizeVmlImages(body: Document) {
  for (const pict of all(body, 'pict')) {
    const shapes = [...pict.getElementsByTagNameNS(V, 'shape')].filter(s => s.getElementsByTagNameNS(V, 'imagedata').length);
    if (!shapes.length) continue;
    const replacement = body.createDocumentFragment();
    for (const shape of shapes) {
      const image = shape.getElementsByTagNameNS(V, 'imagedata')[0];
      const id = image.getAttributeNS(R, 'id'); if (!id) continue;
      const length = (property: string) => {
        const match = shape.getAttribute('style')?.match(new RegExp(`(?:^|;)\\s*${property}\\s*:\\s*([\\d.]+)(pt|px|in|mm|cm)`, 'i'));
        return match ? Number(match[1]) * ({ pt: 4 / 3, px: 1, in: 96, mm: 96 / 25.4, cm: 96 / 2.54 }[match[2].toLowerCase()] || 1) : 0;
      };
      const width = Math.min(10000, length('width')), height = Math.min(10000, length('height'));
      const parsed = new DOMParser().parseFromString(`<w:drawing xmlns:w="${W}" xmlns:wp="${WP}" xmlns:a="${A}" xmlns:pic="${PIC}" xmlns:r="${R}"><wp:inline><wp:extent/><wp:docPr id="1" name="Picture"/><a:graphic><a:graphicData uri="${PIC}"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="Picture"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr/></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>`, 'application/xml');
      const drawing = body.importNode(parsed.documentElement, true);
      const wrap = shape.getElementsByTagNameNS('urn:schemas-microsoft-com:office:word', 'wrap')[0]?.getAttribute('type');
      const coordinates = (shape.getAttribute('wrapcoords') || '').trim().split(/[ ,]+/).map(Number);
      const polygon = coordinates.length >= 6 && coordinates.length % 2 === 0 ? Array.from({ length: coordinates.length / 2 }, (_, i) => [coordinates[i * 2] / 21600, coordinates[i * 2 + 1] / 21600]) : undefined;
      if (/position\s*:\s*absolute/i.test(shape.getAttribute('style') || '') && wrap) drawing.setAttribute('data-ws-floating', JSON.stringify({
        horizontalPosition: { relative: 'margin', offset: Math.round(length('margin-left') * 9525) },
        verticalPosition: { relative: 'paragraph', offset: Math.round(length('margin-top') * 9525) },
        wrap: { type: wrap === 'tight' ? 2 : wrap === 'through' ? 4 : wrap === 'topAndBottom' ? 3 : 1 },
        ...(polygon ? { polygon } : {}),
      }));
      const extent = drawing.getElementsByTagNameNS(WP, 'extent')[0];
      extent.setAttribute('cx', String(Math.round(width * 9525))); extent.setAttribute('cy', String(Math.round(height * 9525)));
      drawing.getElementsByTagNameNS(A, 'blip')[0].setAttributeNS(R, 'r:embed', id);
      drawing.getElementsByTagNameNS(WP, 'docPr')[0].setAttribute('descr', image.getAttributeNS('urn:schemas-microsoft-com:office:office', 'title') || '');
      replacement.append(drawing);
    }
    if (replacement.childNodes.length) pict.replaceWith(replacement);
  }
}
