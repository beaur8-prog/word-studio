import type { HwpDocument } from '@rhwp/core';
import { hwpWrapFingerprint, normalizeHwpWrapText, type HwpSourceWrap } from '../hwp-text-advances';

export function sourceHwpPageHead(doc: HwpDocument, section: number, paragraph: number, page: number,
  y: number, nextY: number, head: string): HwpSourceWrap | undefined {
  const text = normalizeHwpWrapText(head);
  if (text.length < 8 || text.length > 92 || /[\u0000-\u001f\ufffc]/u.test(head) || doc.getParagraphLength(section, paragraph + 1)) return;
  const body = doc.getTextRange(section, paragraph + 2, 0, 0xffffffff);
  if (!normalizeHwpWrapText(body) || body.length > 200 || /[\u0000-\u001f\ufffc]/u.test(body)) return;
  const bounds = JSON.parse(doc.getPageInfo(page));
  if (bounds.columns.length !== 1 || Math.abs(bounds.marginTop + bounds.marginHeader - y) > .16 ||
      !(nextY - y >= 16 && nextY - y <= 96)) return;
  const metrics = [paragraph, paragraph + 1, paragraph + 2].map(index => JSON.parse(doc.getParagraphLineMetrics(section, index, '[]')));
  if (metrics.some(entry => !entry.stored || entry.controlCount || entry.lines.length !== 1) ||
      [paragraph, paragraph + 1, paragraph + 2].some(index => {
        const style = JSON.parse(doc.getParaPropertiesAt(section, index));
        return style.pageBreakBefore || style.keepLines || style.keepWithNext;
      })) return;
  const { fontSize: spacer } = JSON.parse(doc.getCharPropertiesAt(section, paragraph + 1, 0));
  if (!Number.isInteger(spacer) || !(spacer >= 200 && spacer <= 800)) return;
  return { fingerprint: hwpWrapFingerprint(head), ends: [text.length],
    authoredPageHead: { page: page + 1, y, nextY, spacer, next: hwpWrapFingerprint(body) } };
}
