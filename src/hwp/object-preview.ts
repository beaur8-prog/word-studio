// A deliberately small adapter for isolated RHWP shape layer subtrees. It
// consumes drawing data only; it never executes SVG or loads document URLs.
const number = (value: unknown): number => {
  if (typeof value !== 'number' || !Number.isFinite(value) || Math.abs(value) > 1_000_000) throw Error('Invalid drawing coordinate');
  return value;
};
const color = (value: unknown): string | undefined => {
  if (value === null || value === undefined) return;
  if (typeof value !== 'string' || !/^#[a-f\d]{6}(?:[a-f\d]{2})?$/i.test(value)) throw Error('Invalid drawing color');
  return value;
};
function box(value: any) {
  const [x, y, width, height] = ['x', 'y', 'width', 'height'].map(key => number(value?.[key]));
  if (width < 0 || height < 0) throw Error('Invalid drawing bounds');
  return { x, y, width, height };
}
function font(style: any): string {
  const size = number(style.fontSize);
  if (size <= 0 || size > 2000 || typeof style.fontFamily !== 'string' || style.fontFamily.length > 200) throw Error('Invalid object font');
  const serif = /명조|바탕|serif|batang|times|georgia/i.test(style.fontFamily);
  return `${style.italic ? 'italic ' : ''}${style.bold ? 'bold ' : ''}${size}px ${JSON.stringify(style.fontFamily)}, ${serif ? '"Noto Serif KR", "AppleMyungjo", serif' : '"Noto Sans KR", "AppleGothic", sans-serif'}`;
}
function dash(ctx: CanvasRenderingContext2D, value: string | undefined, width: number) {
  const unit = Math.max(.5, width);
  const patterns: Record<string, number[]> = { solid: [], dash: [4, 2], dashed: [4, 2], dot: [1, 2], dotted: [1, 2], dashDot: [4, 2, 1, 2], dashDotDot: [4, 2, 1, 2, 1, 2] };
  if (value !== undefined && !Object.hasOwn(patterns, value)) throw Error('Unsupported object dash');
  ctx.setLineDash((patterns[value ?? 'solid']).map(n => n * unit));
}

export async function rasterizeHwpObject(raw: string, width: number, height: number, signal?: AbortSignal): Promise<string> {
  if (raw.length > 4 * 1024 * 1024 || ![width, height].every(n => Number.isFinite(n) && n > 0 && n <= 10000)) throw Error('Object preview too large');
  const root = JSON.parse(raw), bounds = box(root.bounds), padding = number(root.previewPadding ?? 2);
  if (padding < 2 || padding > 1000 || Math.abs(width - bounds.width - padding * 2) > 1 || Math.abs(height - bounds.height - padding * 2) > 1) throw Error('Object size mismatch');
  const fonts = new Map<string, string>();
  let count = 0;
  const visit = (node: any, draw: (node: any) => void, depth = 0) => {
    if (!node || ++count > 10000 || depth > 40) throw Error('Object drawing too complex');
    draw(node);
    for (const child of node.children ?? []) visit(child, draw, depth + 1);
    if (node.child) visit(node.child, draw, depth + 1);
  };
  let characters = 0;
  visit(root, node => {
    for (const op of node.ops ?? []) {
      if (++count > 10000) throw Error('Object drawing too complex');
      if (op.type === 'textRun') {
        const text = op.displayText ?? op.text ?? '', family = font(op.paintStyle ?? op.style);
        characters += text.length;
        fonts.set(family, (fonts.get(family) ?? '') + text);
        if (characters > 100000 || fonts.size > 100) throw Error('Object text too large');
      }
    }
  });
  signal?.throwIfAborted();
  let abort: (() => void) | undefined;
  try {
    // CJK fonts are split by unicode-range. Loading the default space alone
    // would not wait for the Korean glyphs needed by the actual textbox.
    await Promise.race([Promise.all([...fonts].map(([value, text]) => document.fonts.load(value, text))),
      new Promise((_, reject) => { if (signal) { abort = () => reject(signal.reason); signal.addEventListener('abort', abort, { once: true }); } })]);
  } finally { if (abort) signal?.removeEventListener('abort', abort); }
  signal?.throwIfAborted();
  const scale = Math.min(2, Math.sqrt(4_000_000 / (width * height)));
  const canvas = document.createElement('canvas');
  canvas.width = Math.max(1, Math.ceil(width * scale)); canvas.height = Math.max(1, Math.ceil(height * scale));
  const ctx = canvas.getContext('2d')!;
  ctx.scale(scale, scale); ctx.translate(padding - bounds.x, padding - bounds.y);
  const paintText = (op: any) => {
    if (op.rotation || op.isVertical || op.charOverlap) throw Error('Unsupported object text orientation');
    const b = box(op.bbox), s = op.paintStyle ?? op.style;
    if (s.outlineType || s.emboss || s.engrave || s.shadowType || s.emphasisDot) throw Error('Unsupported object text effect');
    const text = op.displayText ?? op.text;
    if (typeof text !== 'string' || text.length > 32768) throw Error('Invalid object text');
    ctx.font = font(s); ctx.fillStyle = color(s.color) ?? '#000000'; ctx.textBaseline = 'alphabetic';
    const y = b.y + number(op.baseline), glyphs = Array.from(text), positions = op.displayPositions ?? op.positions;
    const draw = (value: string, x: number, advance: number) => {
      if (!value.trim()) return;
      const natural = ctx.measureText(value).width;
      if (!(natural > 0 && advance > 0)) throw Error('Invalid object text advance');
      ctx.save(); ctx.translate(x, y); ctx.scale(advance / natural, 1); ctx.fillText(value, 0, 0); ctx.restore();
    };
    if (Array.isArray(positions) && positions.length === glyphs.length + 1) {
      glyphs.forEach((glyph, i) => draw(glyph, b.x + number(positions[i]), number(positions[i + 1]) - number(positions[i])));
    } else draw(text, b.x, b.width);
    if (s.underline && s.underline !== 'none' || s.strikethrough) {
      ctx.strokeStyle = color(s.underline && s.underline !== 'none' ? s.underlineColor : s.strikeColor) ?? ctx.fillStyle;
      ctx.lineWidth = Math.max(.5, s.fontSize / 18); ctx.setLineDash([]);
      const dy = s.underline && s.underline !== 'none' ? s.fontSize * .12 : -s.fontSize * .3;
      ctx.beginPath(); ctx.moveTo(b.x, y + dy); ctx.lineTo(b.x + b.width, y + dy); ctx.stroke();
    }
  };
  const paint = (op: any) => {
    ctx.save();
    try {
      if (op.type === 'textRun') { paintText(op); return; }
      const b = box(op.bbox), s = op.style ?? {}, transform = op.transform;
      if (transform) {
        ctx.translate(b.x + b.width / 2, b.y + b.height / 2);
        ctx.rotate(number(transform.rotation ?? 0) * Math.PI / 180);
        ctx.scale(transform.horzFlip ? -1 : 1, transform.vertFlip ? -1 : 1);
        ctx.translate(-b.x - b.width / 2, -b.y - b.height / 2);
      }
      ctx.globalAlpha = Math.max(0, Math.min(1, number(s.opacity ?? 1)));
      ctx.beginPath();
      if (op.type === 'line') {
        if ((s.startArrow && s.startArrow !== 'none') || (s.endArrow && s.endArrow !== 'none') || (s.lineType && s.lineType !== 'single')) throw Error('Unsupported object line');
        ctx.moveTo(number(op.x1), number(op.y1)); ctx.lineTo(number(op.x2), number(op.y2));
        ctx.strokeStyle = color(s.color) ?? '#000000'; ctx.lineWidth = number(s.width ?? .5);
        dash(ctx, s.dash, ctx.lineWidth); ctx.stroke(); return;
      }
      if (op.type === 'ellipse') ctx.ellipse(b.x + b.width / 2, b.y + b.height / 2, b.width / 2, b.height / 2, 0, 0, Math.PI * 2);
      else if (op.type === 'rectangle') ctx.roundRect(b.x, b.y, b.width, b.height, Math.max(0, Math.min(number(op.cornerRadius ?? 0), b.width / 2, b.height / 2)));
      else if (op.type === 'path') {
        if (!Array.isArray(op.commands) || op.commands.length > 10000) throw Error('Invalid object path');
        for (const command of op.commands) {
          if (command.type === 'moveTo') ctx.moveTo(number(command.x), number(command.y));
          else if (command.type === 'lineTo') ctx.lineTo(number(command.x), number(command.y));
          else if (command.type === 'curveTo') ctx.bezierCurveTo(number(command.x1), number(command.y1), number(command.x2), number(command.y2), number(command.x), number(command.y));
          else if (command.type === 'closePath') ctx.closePath();
          else throw Error('Unsupported object path command');
        }
      } else if (op.type === 'polygon') {
        if (!Array.isArray(op.points) || op.points.length < 2 || op.points.length > 10000) throw Error('Invalid object polygon');
        op.points.forEach((p: any, i: number) => i ? ctx.lineTo(number(p.x), number(p.y)) : ctx.moveTo(number(p.x), number(p.y)));
        ctx.closePath();
      } else throw Error('Unsupported object operation');
      const fill = color(s.fillColor), stroke = color(s.strokeColor);
      if (fill) { ctx.fillStyle = fill; ctx.fill(); }
      if (stroke && number(s.strokeWidth ?? 0) > 0) { ctx.strokeStyle = stroke; ctx.lineWidth = s.strokeWidth; dash(ctx, s.strokeDash, ctx.lineWidth); ctx.stroke(); }
    } finally { ctx.restore(); }
  };
  count = 0;
  const render = (node: any, depth = 0) => {
    signal?.throwIfAborted();
    if (++count > 10000 || depth > 40) throw Error('Object drawing too complex');
    ctx.save();
    try {
      if (node.kind === 'group') for (const child of node.children ?? []) render(child, depth + 1);
      else if (node.kind === 'clipRect') { const b = box(node.clip); ctx.beginPath(); ctx.rect(b.x, b.y, b.width, b.height); ctx.clip(); render(node.child, depth + 1); }
      else if (node.kind === 'leaf') for (const op of node.ops ?? []) paint(op);
      else throw Error('Unsupported object node');
    } finally { ctx.restore(); }
  };
  try { render(root); return canvas.toDataURL('image/png'); }
  finally { canvas.width = canvas.height = 0; }
}
