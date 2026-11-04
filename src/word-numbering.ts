export function wordNumberText(value: number, format: string) {
  if (/^(lower|upper)Letter$/.test(format)) {
    let text = ''; for (let n = value; n > 0; n = Math.floor((n - 1) / 26)) text = String.fromCharCode(65 + (n - 1) % 26) + text;
    return format.startsWith('lower') ? text.toLowerCase() : text;
  }
  if (/^(lower|upper)Roman$/.test(format)) {
    let text = '', remaining = value;
    for (const [n, label] of [[1000,'M'],[900,'CM'],[500,'D'],[400,'CD'],[100,'C'],[90,'XC'],[50,'L'],[40,'XL'],[10,'X'],[9,'IX'],[5,'V'],[4,'IV'],[1,'I']] as const)
      while (remaining >= n && text.length < 1000) { text += label; remaining -= n; }
    return format.startsWith('lower') ? text.toLowerCase() : text;
  }
  return format === 'decimalZero' ? String(value).padStart(2, '0') : String(value);
}
