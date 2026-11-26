import documentLimits from '../electron/document-limits.json';

export const hwpCellRasterFillAttribute = 'data-hwp-cell-raster-fill';
export const hwpCellRasterFillNamespace = 'urn:wordstudio:hwp-cell-raster-fill:1';
export type HwpCellRasterPicture = { source: string; brightness: number };

export function isHwpCellRasterBrightness(value: unknown): value is number {
  return typeof value === 'number' && Number.isInteger(value) && Math.abs(value) <= 100;
}

const maxSourceLength = Math.ceil(documentLimits.embeddedImageBytes / 3) * 4 + 64;

export function hwpCellRasterSource(value: unknown): value is string {
  if (typeof value !== 'string' || value.length > maxSourceLength) return false;
  const match = /^data:image\/(jpeg|png);base64,([A-Za-z0-9+/]+={0,2})$/.exec(value);
  if (!match || match[2].length % 4 !== 0) return false;
  return match[1] === 'jpeg' ? match[2].startsWith('/9j/') : match[2].startsWith('iVBORw0KGgo');
}

export function parseHwpCellRasterFill(value: unknown) {
  if (typeof value !== 'string' || value.length > maxSourceLength + 40) return;
  const parts = value.split(' ', 5);
  if (parts.length !== 3 && parts.length !== 4) return;
  const [width, height] = parts, source = parts.at(-1);
  const w = Number(width), h = Number(height), brightness = parts.length === 4 ? Number(parts[2]) : 0;
  if (!/^\d{1,4}(?:\.\d{1,4})?$/.test(width) || !/^\d{1,4}(?:\.\d{1,4})?$/.test(height) ||
      w < 1 || h < 1 || !isHwpCellRasterBrightness(brightness) ||
      parts.length === 4 && String(brightness) !== parts[2] || !hwpCellRasterSource(source)) return;
  return { width: w, height: h, source, brightness };
}

export function encodeHwpCellRasterFill(width: number, height: number, source: string, brightness = 0) {
  const effect = brightness ? `${brightness} ` : '';
  return `${+width.toFixed(4)} ${+height.toFixed(4)} ${effect}${source}`;
}

function brightnessFilter(brightness: number) {
  const channel = brightness / 100;
  const transfer = [...'RGB'].map(key => `<feFunc${key} type="linear" slope="1" intercept="${channel}"/>`).join('');
  return `<filter id="bright" color-interpolation-filters="linearRGB"><feComponentTransfer>${transfer}</feComponentTransfer></filter>`;
}

/** Match the Reader's per-channel linear-light transfer. */
export function hwpCellRasterPaintSource({ source, brightness }: HwpCellRasterPicture) {
  if (!brightness) return source;
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100" preserveAspectRatio="none">${brightnessFilter(brightness)}<g filter="url(#bright)"><image width="100" height="100" preserveAspectRatio="none" href="${source}"/></g></svg>`;
  return `data:image/svg+xml,${encodeURIComponent(svg)}`;
}

export function hwpRasterPictureFilter(brightness: number) {
  if (!brightness || !isHwpCellRasterBrightness(brightness)) return '';
  const svg = `<svg xmlns="http://www.w3.org/2000/svg">${brightnessFilter(brightness)}</svg>`;
  return `url("data:image/svg+xml,${encodeURIComponent(svg)}#bright")`;
}
