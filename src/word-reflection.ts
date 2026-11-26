export const wordReflectionNamespace = 'http://schemas.microsoft.com/office/word/2010/wordml';
export const wordReflectionAttribute = 'data-word-reflection';
export const wordParagraphReflectionAttribute = 'data-word-paragraph-reflection';

const intMax = 2147483647;
const integerRanges = {
  blurRad: [0, intMax], dist: [0, intMax],
  stA: [0, 100000], stPos: [0, 100000], endA: [0, 100000], endPos: [0, 100000],
  dir: [0, 21599999], fadeDir: [0, 21599999],
  sx: [-2147483648, intMax], sy: [-2147483648, intMax],
  kx: [-5399999, 5399999], ky: [-5399999, 5399999],
} as const;
const alignments = ['none', 'tl', 't', 'tr', 'l', 'ctr', 'r', 'bl', 'b', 'br'];
type NumericAttribute = keyof typeof integerRanges;
type Reflection = Partial<Record<NumericAttribute | 'algn', string>>;

function validReflection(input: unknown): input is Reflection {
  if (!input || typeof input !== 'object' || Array.isArray(input)) return false;
  for (const [name, value] of Object.entries(input)) {
    if (typeof value !== 'string') return false;
    if (name === 'algn') {
      if (!alignments.includes(value)) return false;
      continue;
    }
    if (!Object.hasOwn(integerRanges, name) || !/^[+-]?\d{1,17}$/.test(value)) return false;
    const [smallest, largest] = integerRanges[name as NumericAttribute];
    const number = Number(value);
    if (number < smallest || number > largest) return false;
  }
  return true;
}

export function parseWordReflection(serialized: unknown): Reflection | null {
  if (typeof serialized !== 'string' || serialized.length > 512) return null;
  try {
    const reflection: unknown = JSON.parse(serialized);
    return validReflection(reflection) ? reflection : null;
  } catch { return null; }
}

export function readWordReflection(properties?: Element | null): string | undefined {
  const matches = Array.from(properties?.children ?? []).filter(element =>
    element.namespaceURI === wordReflectionNamespace && element.localName === 'reflection');
  if (matches.length !== 1 || matches[0].children.length || matches[0].textContent?.trim()) return;
  const attributes = Array.from(matches[0].attributes).filter(attribute =>
    attribute.namespaceURI !== 'http://www.w3.org/2000/xmlns/');
  if (attributes.some(attribute => attribute.namespaceURI !== wordReflectionNamespace)) return;
  const reflection = Object.fromEntries(attributes.map(attribute => [attribute.localName, attribute.value]));
  if (validReflection(reflection)) return JSON.stringify(reflection);
}

export function wordReflectionXmlAttributes(serialized: unknown): Record<string, string> | null {
  const reflection = parseWordReflection(serialized);
  return reflection && Object.fromEntries(Object.entries(reflection).map(([name, value]) => ['w14:' + name, value]));
}
