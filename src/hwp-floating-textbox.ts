// Native HWP units keep the authored line boxes without float accumulation.
export type NativeBoxLine = [start:number, top:number, height:number, baseline:number, gap:number, width:number, flags:number];

export type HwpFloatingTextBox = {
  width: number;
  height: number;
  order: number;
  count: number;
  pitch: number;
  hanging: number;
  size: number;
  scale: number;
  alignment: 'left' | 'center' | 'right' | 'justify';
  faces: string[];
  baselines: number[];
  stored?: NativeBoxLine[];
  child?: { x: number; y: number; before: number; band: number; host?: NativeBoxLine };
};

const number = (value: unknown, lo: number, hi: number): value is number =>
  typeof value === 'number' && Number.isFinite(value) && value >= lo && value <= hi;

export function nativeBoxLine(value: unknown): value is NativeBoxLine {
  return Array.isArray(value) && value.length === 7 && value.every(Number.isInteger) &&
    number(value[0], 0, 1000) && number(value[1], 0, 120000) && number(value[2], 1, 20000) &&
    number(value[3], 1, value[2]) && number(value[4], 0, 15000) && number(value[5], 1, 150000) &&
    number(value[6], 0, 0x7fffffff);
}

/** Native geometry is optional. A malformed record must leave an ordinary
 * editable DOCX textbox with its standard runs and anchors. */
export function floatingTextBox(value: unknown, lines: number): HwpFloatingTextBox | undefined {
  if (!value || typeof value !== 'object' || lines < 1 || lines > 12) return;
  const { width, height, order, count, pitch, hanging, size, scale, alignment, faces, baselines, stored, child } = value as HwpFloatingTextBox;
  if (!number(width, 100, 2000) || !number(height, 36, 1600) || !Number.isInteger(count) ||
      !number(count, 3, 16) || !Number.isInteger(order) || !number(order, 0, count - 1) ||
      !number(size, 6, 72) || !number(pitch, size, size * 3) || !number(hanging, 0, width / 3) ||
      !number(scale, 70, 100) || !Number.isInteger(scale) ||
      !['left', 'center', 'right', 'justify'].includes(alignment) ||
      !Array.isArray(faces) || !faces.length || faces.length > 5 || new Set(faces).size !== faces.length ||
      faces.some(face => typeof face !== 'string' || !face.trim() || face.length > 80 || /[\u0000-\u001f]/.test(face)) ||
      !Array.isArray(baselines) || baselines.length !== lines || baselines.some((baseline, i) =>
        !number(baseline, size / 2, height - 1) || i && Math.abs(baseline - baselines[i - 1] - pitch) > .22)) return;
  if (stored !== undefined && (!Array.isArray(stored) || stored.length !== lines ||
      stored.some((line, i) => !nativeBoxLine(line) || !i && line[0] !== 0 ||
        i > 0 && (line[0] <= stored[i - 1][0] || line[1] <= stored[i - 1][1])))) return;
  if (child !== undefined && (!child || !number(child.x, 0, 2000) || !number(child.y, 0, 1600) ||
      !number(child.before, 0, 100) || !number(child.band, height, height + 100) || child.host !== undefined &&
      (!nativeBoxLine(child.host) || child.host[0] !== 0 || Math.abs(child.host[2] + child.host[4] - child.band * 75) > .5))) return;
  return { width, height, order, count, pitch, hanging, size, scale, alignment, faces, baselines,
    ...(stored ? { stored } : {}),
    ...(child ? { child: { x: child.x, y: child.y, before: child.before, band: child.band,
      ...(child.host ? { host: child.host } : {}) } } : {}) };
}
