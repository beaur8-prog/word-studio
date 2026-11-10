import { sanitizeImportedHtml } from './document-html';
import type { PageSettings } from './types';

export type PrintFontRequest = { font: string; text: string };

function abortable<T>(pending: Promise<T>, signal?: AbortSignal): Promise<T> {
  if (!signal) return pending;
  signal.throwIfAborted();
  return new Promise((resolve, reject) => {
    const abort = () => {
      signal.removeEventListener('abort', abort);
      reject(signal.reason ?? new DOMException('Printing cancelled', 'AbortError'));
    };
    signal.addEventListener('abort', abort, { once: true });
    pending.then(value => {
      signal.removeEventListener('abort', abort); resolve(value);
    }, error => {
      signal.removeEventListener('abort', abort); reject(error);
    });
  });
}

/** fonts.ready can resolve before any source text has used its lazy faces.
 * Resolve the real CSS stacks with the document present, then load their
 * actual characters and styles before the first page or line is measured. */
export async function preparePrintFonts(host: HTMLElement, source: HTMLElement, notes: Iterable<HTMLElement>,
  initialPage: PageSettings, extra: readonly PrintFontRequest[], signal?: AbortSignal): Promise<number> {
  signal?.throwIfAborted();
  const probe = document.createElement('div');
  probe.className = 'print-font-probe tiptap'; probe.setAttribute('aria-hidden', 'true');
  probe.style.cssText = 'position:fixed;left:-100000px;top:0;width:700px;min-height:0;visibility:hidden;pointer-events:none';
  probe.append(...[...source.childNodes].map(node => node.cloneNode(true)));
  const noteProbe = document.createElement('div'); noteProbe.className = 'print-footnotes';
  noteProbe.append(...[...notes].map(note => note.cloneNode(true))); probe.append(noteProbe);

  const pages = [initialPage];
  for (const node of source.querySelectorAll<HTMLElement>('[data-section-page],[data-section-content]')) {
    try {
      const page = JSON.parse(node.dataset.sectionPage || node.dataset.sectionContent || '');
      if (page && typeof page === 'object') pages.push(page as PageSettings);
    }
    catch { /* The normal section parser still reports invalid page settings. */ }
  }
  const furniture = new Set<string>();
  for (const page of pages) for (const prefix of ['', 'first', 'even']) for (const part of ['header', 'footer']) {
    const property = (prefix ? prefix + part[0].toUpperCase() + part.slice(1) : part) as keyof PageSettings;
    const html = page[(property + 'Html') as keyof PageSettings], text = page[property];
    const key = typeof html === 'string' && html ? html : typeof text === 'string' ? text : '';
    if (!key || furniture.has(key)) continue;
    furniture.add(key);
    const container = document.createElement('div'); container.className = 'document-furniture';
    if (typeof html === 'string' && html) container.innerHTML = sanitizeImportedHtml(html).html;
    else container.textContent = key;
    probe.append(container);
  }
  // Pictures are decoded by the normal print path; this tree only needs text
  // styles and must not start a second set of picture decodes.
  probe.querySelectorAll('img,svg,video,audio').forEach(node => node.remove());
  host.append(probe);
  try {
    // Include implicit fallback, transformed text and marker requests too.
    probe.getBoundingClientRect();
    const groups = new Map<string, Set<string>>();
    for (const request of extra) {
      const chars = groups.get(request.font) ?? new Set<string>();
      for (const character of request.text) chars.add(character);
      groups.set(request.font, chars);
    }
    for (const element of [probe, ...probe.querySelectorAll<HTMLElement>(
      'p,span,strong,b,em,i,u,s,mark,a,li,h1,h2,h3,h4,h5,h6,pre,code,td,th,[data-page-number],[data-note-reference],[data-document-field]')]) {
      const css = getComputedStyle(element);
      const query = [css.fontStyle, css.fontWeight, css.fontSize, css.fontFamily].join(' ');
      const chars = groups.get(query) ?? new Set<string>();
      const sample = element.matches('[data-page-number],[data-note-reference],[data-document-field]')
        ? element.textContent || '0123456789' : element.textContent || ' ';
      for (const character of sample) chars.add(character);
      groups.set(query, chars);
    }
    const loaded = await abortable(Promise.allSettled([...groups].map(([font, chars]) =>
      document.fonts.load(font, [...chars].join('')))), signal);
    await abortable(document.fonts.ready, signal);
    return loaded.filter(result => result.status === 'rejected').length;
  } finally { probe.remove(); }
}
