import type JSZip from 'jszip';
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const X = 'http://schemas.microsoft.com/office/word/2010/wordml';
const child = (root: Element | undefined, name: string) => root && [...root.children].find(n => n.namespaceURI === W && n.localName === name);
const xvalue = (node: Element | undefined, key = 'val') => node?.getAttributeNS(X, key) || node?.getAttributeNS(W, key) || '';
const checked = (node: Element | undefined) => !!node && !/^(0|false|off)$/i.test(xvalue(node) || 'true');

/** Word checkbox controls are displayed as editable symbols in Studio. Keep
 * the selected state and custom Unicode/font without exposing an HTML input.
 * Unwrap content in place so bookmarks and original run formatting survive. */
export function flattenDocxCheckboxes(doc: Document) {
  let count = 0;
  for (const sdt of [...doc.getElementsByTagNameNS(W, 'sdt')]) {
    if (!doc.documentElement.contains(sdt)) continue;
    const props = child(sdt, 'sdtPr'), box = props?.getElementsByTagNameNS(X, 'checkbox')[0];
    if (!box) continue;
    const on = checked(box.getElementsByTagNameNS(X, 'checked')[0]);
    const state = box.getElementsByTagNameNS(X, on ? 'checkedState' : 'uncheckedState')[0];
    const raw = xvalue(state), value = /^[\da-f]{1,6}$/i.test(raw) ? parseInt(raw, 16) : -1;
    const scalar = value >= 32 && value <= 0x10ffff && !(value >= 0xd800 && value <= 0xdfff) && (value & 0xffff) < 0xfffe;
    const glyph = String.fromCodePoint(scalar ? value : on ? 0x2612 : 0x2610);
    const content = child(sdt, 'sdtContent') || doc.createElementNS(W, 'w:sdtContent');
    const text = [...content.getElementsByTagNameNS(W, 't')].find(t => {
      for (let p = t.parentElement; p && p !== content; p = p.parentElement) if (p.namespaceURI === W && p.localName === 'del') return false;
      return true;
    });
    let run: Element | undefined = text?.parentElement ?? undefined;
    if (!text || run?.namespaceURI !== W || run.localName !== 'r') {
      run = doc.createElementNS(W, 'w:r');
      const defaults = child(props, 'rPr'); if (defaults) run.append(defaults.cloneNode(true));
      const node = doc.createElementNS(W, 'w:t'); node.textContent = glyph; run.append(node);
      let host = content.getElementsByTagNameNS(W, 'p')[0] as Element | undefined;
      if (!host && ['body', 'tc', 'hdr', 'ftr'].includes(sdt.parentElement?.localName || '')) {
        host = doc.createElementNS(W, 'w:p'); content.append(host);
      }
      (host || content).append(run);
    } else {
      // Valid controls contain one character. Do not discard extra authored
      // text in a malformed control while repairing its first displayed box.
      text.textContent = glyph + ([...(text.textContent || '')].length > 1 ? text.textContent : '');
    }
    if (!run) continue;
    const defaults = child(props, 'rPr');
    if (defaults) {
      let pr = child(run, 'rPr');
      if (!pr) run.insertBefore(defaults.cloneNode(true), run.firstChild);
      else for (const property of [...defaults.children]) if (!child(pr, property.localName))
        pr.insertBefore(property.cloneNode(true), pr.firstChild);
    }
    const font = xvalue(state, 'font').trim();
    if (font && font.length <= 100 && !/[\u0000-\u001f]/.test(font)) {
      let pr = child(run, 'rPr'); if (!pr) { pr = doc.createElementNS(W, 'w:rPr'); run!.insertBefore(pr, run!.firstChild); }
      child(pr, 'rFonts')?.remove(); const fonts = doc.createElementNS(W, 'w:rFonts');
      for (const key of ['ascii', 'hAnsi', 'eastAsia', 'cs']) fonts.setAttributeNS(W, 'w:' + key, font);
      pr.insertBefore(fonts, pr.firstChild);
    }
    sdt.replaceWith(...content.childNodes); count++;
  }
  return count;
}

/** Mammoth represents legacy FORMCHECKBOX fields as HTML inputs. Convert only
 * that DOCX intermediate tree; the general HTML sanitizer keeps rejecting inputs. */
export function restoreDocxCheckboxInputs(root: DocumentFragment) {
  for (const input of root.querySelectorAll<HTMLInputElement>('input[type="checkbox"]'))
    input.replaceWith(document.createTextNode(input.checked ? '☒' : '☐'));
}

export async function countDocxCheckboxes(zip: JSZip, body: Document) {
  const docs = [body];
  const parts = Object.keys(zip.files).filter(p => /^word\/(?:header\d+|footer\d+|footnotes|endnotes|comments)\.xml$/.test(p));
  docs.push(...await Promise.all(parts.map(async p => new DOMParser().parseFromString(await zip.file(p)!.async('string'), 'application/xml'))));
  return docs.reduce((sum, doc) => sum + doc.getElementsByTagNameNS(X, 'checkbox').length +
    ([...doc.getElementsByTagNameNS(W, 'instrText')].map(n => n.textContent || '').join(' ').match(/\bFORMCHECKBOX\b/g) || []).length, 0);
}
