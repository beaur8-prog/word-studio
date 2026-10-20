export const HWP_PAGE_GAP = 32;
export const HWP_VIEW_PADDING = 22;
export interface HwpViewportLayout { offsets: number[]; heights: number[]; width: number; height: number }

export function hwpViewportLayout(pages: readonly { width: number; height: number }[], scale: number): HwpViewportLayout {
  const offsets: number[] = [], heights = pages.map(page => page.height * scale);
  let height = 0, width = 0;
  pages.forEach((page, i) => { offsets.push(height); height += heights[i] + HWP_PAGE_GAP; width = Math.max(width, page.width * scale); });
  return { offsets, heights, width, height: Math.max(0, height - HWP_PAGE_GAP) };
}

export function hwpPageAtOffset(layout: HwpViewportLayout, position: number): number {
  let low = 0, high = layout.offsets.length - 1;
  while (low < high) {
    const middle = Math.ceil((low + high) / 2);
    if (layout.offsets[middle] <= position) low = middle; else high = middle - 1;
  }
  return Math.max(0, low);
}

export function hwpVisiblePages(layout: HwpViewportLayout, top: number, height: number) {
  const first = hwpPageAtOffset(layout, top), last = hwpPageAtOffset(layout, top + height);
  let current = first, mostVisible = -1;
  for (let page = first; page <= last; page++) {
    const visible = Math.max(0, Math.min(top + height, layout.offsets[page] + layout.heights[page]) - Math.max(top, layout.offsets[page]));
    if (visible > mostVisible + 1) { current = page; mostVisible = visible; }
  }
  return { first, last, current };
}
