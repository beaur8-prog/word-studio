import type { IFloating } from 'docx';

export function imageFloating(value: unknown): IFloating | undefined {
  if (typeof value !== 'string' || value.length > 2000) return;
  try {
    const data = JSON.parse(value);
    const position = (p: any, vertical: boolean) => p && typeof p === 'object'
      && (p.offset === undefined || typeof p.offset === 'number' && Number.isFinite(p.offset) && Math.abs(p.offset) <= 100000000)
      && (p.align === undefined || (vertical ? ['top','center','bottom','inside','outside'] : ['left','center','right','inside','outside']).includes(p.align))
      && (p.relative === undefined || (vertical ? ['page','margin','paragraph','line','topMargin','bottomMargin','insideMargin','outsideMargin'] : ['page','margin','column','character','leftMargin','rightMargin','insideMargin','outsideMargin']).includes(p.relative));
    if (!position(data.horizontalPosition, false) || !position(data.verticalPosition, true) || ![0,1,2,3,4].includes(data.wrap?.type)) return;
    return { horizontalPosition: data.horizontalPosition, verticalPosition: data.verticalPosition, wrap: { type: data.wrap.type },
      ...(data.margins && typeof data.margins === 'object' ? { margins: Object.fromEntries(['top','right','bottom','left'].map(k => [k, Math.max(0, Math.min(10000000, Number(data.margins[k]) || 0))])) } : {}),
      behindDocument: data.behindDocument === true, allowOverlap: true, layoutInCell: true };
  } catch { return; }
}

export function floatingCss(value: unknown, width: number) {
  const data = imageFloating(value); if (!data) return '';
  const horizontal = data.horizontalPosition, vertical = data.verticalPosition;
  if (data.wrap?.type === 3) return `display:block;clear:both;margin-left:${horizontal.align === 'right' ? 'auto' : horizontal.align === 'center' ? 'auto' : '0'};margin-right:${horizontal.align === 'right' ? '0' : 'auto'}`;
  const x = Math.max(0, (horizontal.offset || 0) / 9525), y = vertical.relative === 'paragraph' ? Math.max(0, (vertical.offset || 0) / 9525) : 0;
  const side = horizontal.align === 'right' || x > width / 4 ? 'right' : 'left';
  const distance = Math.max(0, (side === 'right' ? data.margins?.left : data.margins?.right) || 0) / 9525;
  return `float:${side};margin-top:${y}px;margin-bottom:4px;${side === 'right' ? `margin-left:${distance}px;${x ? `margin-right:max(0px,calc(100% - ${x + width}px));` : ''}` : `margin-right:${distance}px;`}shape-outside:margin-box`;
}

export function drawingFloating(drawing: Element) {
  const normalized = drawing.getAttribute('data-ws-floating'); if (normalized) return imageFloating(normalized);
  const anchor = [...drawing.children].find(el => el.localName === 'anchor'); if (!anchor) return;
  const read = (kind: string) => {
    const p = [...anchor.children].find(el => el.localName === kind);
    const align = [...(p?.children || [])].find(el => el.localName === 'align')?.textContent;
    const offset = Number([...(p?.children || [])].find(el => el.localName === 'posOffset')?.textContent || 0);
    return { relative: p?.getAttribute('relativeFrom') || (kind === 'positionH' ? 'column' : 'paragraph'), ...(align ? { align } : { offset }) };
  };
  const type = [...anchor.children].find(el => /^wrap/.test(el.localName))?.localName;
  return imageFloating(JSON.stringify({ horizontalPosition: read('positionH'), verticalPosition: read('positionV'),
    wrap: { type: ({wrapNone:0,wrapSquare:1,wrapTight:2,wrapTopAndBottom:3,wrapThrough:4} as Record<string, number>)[type || ''] ?? 1 },
    margins: Object.fromEntries(['top','right','bottom','left'].map(k=>[k, Number(anchor.getAttribute(`dist${k[0].toUpperCase()}`)) || 0])), behindDocument: anchor.getAttribute('behindDoc') === '1' }));
}
