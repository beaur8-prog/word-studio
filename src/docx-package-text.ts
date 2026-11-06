import type JSZip from 'jszip';

/** Keep a surrogate pair together before JSZip splits strings into 16 KiB chunks. */
export async function encodeDocxXmlParts(zip: JSZip): Promise<void> {
  const encoder = new TextEncoder();
  for (const part of Object.values(zip.files)) {
    if (!part.dir && /\.(xml|rels)$/.test(part.name))
      zip.file(part.name, encoder.encode(await part.async('string')));
  }
}
