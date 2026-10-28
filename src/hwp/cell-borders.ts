const widths = [.1, .12, .15, .2, .25, .3, .4, .5, .6, .7, 1, 1.5, 2, 3, 4, 5];

/** HWP stores a width index, not a CSS pixel width, at every nesting depth. */
export function hwpCellBorders(cell: Record<string, any>): string {
  return ['Left', 'Right', 'Top', 'Bottom'].map(side => {
    const border = cell['border' + side];
    const kind = !border?.type ? 'none' : border.type >= 8 && border.type <= 11 ? 'double'
      : border.type === 3 || border.type === 7 ? 'dotted' : border.type >= 2 && border.type <= 6 ? 'dashed' : 'solid';
    return `border-${side.toLowerCase()}:${widths[border?.width] ?? .1}mm ${kind} ${/^#[a-f0-9]{6}$/i.test(border?.color) ? border.color : '#000000'};`;
  }).join('');
}
