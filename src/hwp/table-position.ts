import type { HwpDocument } from '@rhwp/core';
import { tableAnchor } from '../table-anchor-data';

export function hwpMarginCenteredTableAnchor(doc: HwpDocument, properties: Record<string, any>,
  position: { page: number; x: number; y: number; w: number; h: number }) {
  if (properties.horzRelTo !== 'Page' || properties.horzAlign !== 'Center' || properties.horzOffset !== 0 ||
      properties.vertRelTo !== 'Paper' || properties.vertAlign !== 'Top' || properties.textWrap !== 'TopAndBottom' ||
      properties.hasCaption || properties.treatAsChar || properties.allowOverlap ||
      ![position.x, position.y, position.w, position.h].every(Number.isFinite) || position.w <= 0 || position.h <= 0) return;
  const page = JSON.parse(doc.getPageInfo(position.page));
  const y = (Number(properties.vertOffset) + Number(properties.outerTop)) / 75;
  const x = page.bodyLeft + Math.max(0, (page.bodyRight - page.bodyLeft - position.w) / 2);
  // Native page coordinates are rounded to a tenth of a pixel. Verify that
  // the authored alignment accounts for the observed object position.
  if (![x, y, page.width, page.height].every(Number.isFinite) ||
      Math.abs(position.x - x) > .15 || Math.abs(position.y - y) > .15 ||
      position.x + position.w > page.width + .15 || position.y + position.h > page.height + .15) return;
  return tableAnchor({ x: 0, y, bottom: Number(properties.outerBottom) / 75, horizontal: 'margin', horizontalAlign: 'center' });
}
