export function cssPoints(value: unknown): number | undefined {
  const match = String(value).trim().match(/^((?:\d+(?:\.\d*)?|\.\d+))(pt|px|mm|cm|in|pc)?$/i);
  if (!match) return undefined;
  const scale = { pt: 1, px: .75, mm: 72 / 25.4, cm: 72 / 2.54, in: 72, pc: 12 };
  return Math.min(400, Number(match[1]) * scale[(match[2]?.toLowerCase() || 'pt') as keyof typeof scale]);
}
