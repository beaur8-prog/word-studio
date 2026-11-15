import type { WordBackground } from './types';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const V = 'urn:schemas-microsoft-com:vml', O = 'urn:schemas-microsoft-com:office:office';
const XMLNS = 'http://www.w3.org/2000/xmlns/';
export const MAX_BACKGROUND_XML = 32768;
const rgb = /^[0-9a-f]{6}$/i;

/** Keep a bounded, relationship-free VML fill, never an arbitrary shape tree.
 * Word's cached RGB must remain independent of the exporter's theme palette. */
export function normalizeBackgroundDrawing(raw: unknown): string | null {
  if (typeof raw !== 'string' || !raw || raw.length > MAX_BACKGROUND_XML || /<!DOCTYPE|<!ENTITY/i.test(raw)) return null;
  const doc = new DOMParser().parseFromString(raw, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length) return null;
  const root = doc.documentElement;
  if (root.namespaceURI !== V || root.localName !== 'background') return null;
  const allowed = new Map([
    [V + ':background', new Set(['id', 'fillcolor', 'filled', O + ':bwmode', O + ':targetscreensize'])],
    [V + ':fill', new Set(['color', 'color2', 'opacity', 'opacity2', 'focus', 'focusposition', 'focussize',
      'method', 'type', 'angle', 'aspect', 'colors', 'rotate', 'recolor'])],
    [O + ':fill', new Set([V + ':ext', 'type'])],
  ]);
  let nodes = 0;
  const check = (element: Element, depth: number): boolean => {
    if (++nodes > 128 || depth > 8 || element.attributes.length > 32) return false;
    const key = element.namespaceURI + ':' + element.localName, names = allowed.get(key);
    if (!names) return false;
    for (const attribute of [...element.attributes]) {
      if (attribute.namespaceURI === XMLNS) continue;
      const name = attribute.namespaceURI ? attribute.namespaceURI + ':' + attribute.localName : attribute.name;
      if (!names.has(name) || attribute.value.length > 1024 || !/^[\w\s#.%(),;+[\]-]*$/.test(attribute.value)) return false;
      // VML's bracketed theme index otherwise overrides the cached RGB under
      // the newly generated theme. Relative fill lighten/darken stays relative.
      if (['fillcolor', 'color', 'color2'].includes(attribute.localName))
        attribute.value = attribute.value.replace(/(#[0-9a-f]{6})\s*\[\d+\]/ig, '$1');
    }
    const next = key === V + ':background' ? V + ':fill' : key === V + ':fill' ? O + ':fill' : null;
    if (element.children.length > 1 || [...element.childNodes].some(n => n.nodeType !== Node.ELEMENT_NODE && n.textContent?.trim())) return false;
    return [...element.children].every(child => child.namespaceURI + ':' + child.localName === next && check(child, depth + 1));
  };
  return check(root, 0) ? new XMLSerializer().serializeToString(root) : null;
}

export function validWordBackground(value: unknown): value is WordBackground {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false;
  const bg = value as Record<string, unknown>;
  return Object.keys(bg).every(key => ['color', 'display', 'drawingXml'].includes(key)) &&
    typeof bg.color === 'string' && rgb.test(bg.color) && typeof bg.display === 'boolean' &&
    (bg.drawingXml === undefined || normalizeBackgroundDrawing(bg.drawingXml) !== null);
}

export function readWordBackground(document: Document, settings: Document | undefined, warnings: string[]): WordBackground | undefined {
  const source = [...document.documentElement.children].find(n => n.namespaceURI === W && n.localName === 'background');
  if (!source) return;
  const color = source.getAttributeNS(W, 'color') || '';
  if (!rgb.test(color)) { warnings.push('페이지 배경의 색상을 완전히 가져오지 못했습니다.'); return; }
  const display = settings?.getElementsByTagNameNS(W, 'displayBackgroundShape')[0];
  const result: WordBackground = { color: color.toUpperCase(),
    display: !!display && !/^(0|false|off)$/i.test(display.getAttributeNS(W, 'val') || 'true') };
  const drawing = [...source.children].find(n => n.namespaceURI === V && n.localName === 'background');
  if (drawing) {
    const normalized = normalizeBackgroundDrawing(new XMLSerializer().serializeToString(drawing));
    if (normalized) result.drawingXml = normalized;
    else warnings.push('지원하지 않는 페이지 배경 효과는 색상만 가져왔습니다.');
  }
  return result;
}

/** Serialization metadata only. Studio's editing/printing background policy
 * remains separate from Word's displayBackgroundShape document setting. */
export function writeWordBackground(document: Document, settings: Document, background: WordBackground) {
  const root = document.documentElement;
  for (const old of [...root.children].filter(n => n.namespaceURI === W && n.localName === 'background')) old.remove();
  const node = document.createElementNS(W, 'w:background');
  node.setAttributeNS(W, 'w:color', background.color.toUpperCase());
  if (background.drawingXml) {
    const normalized = normalizeBackgroundDrawing(background.drawingXml);
    if (!normalized) throw new Error('페이지 배경 효과가 올바르지 않습니다.');
    node.append(document.importNode(new DOMParser().parseFromString(normalized, 'application/xml').documentElement, true));
  }
  root.insertBefore(node, root.firstChild);
  let display = settings.getElementsByTagNameNS(W, 'displayBackgroundShape')[0];
  if (!display) { display = settings.createElementNS(W, 'w:displayBackgroundShape'); settings.documentElement.insertBefore(display, settings.documentElement.firstChild); }
  display.setAttributeNS(W, 'w:val', background.display ? 'true' : 'false');
}
