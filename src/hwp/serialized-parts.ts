import type { HwpDocument } from '@rhwp/core';
import { unzipSync } from 'fflate';

/** The engine's normalized HWPX is a common, read-only view of both formats.
 * Keep only metadata/embedded OLE parts, with an independent expansion budget. */
export function serializedParts(doc: HwpDocument) {
  let size = 0;
  return unzipSync(doc.exportHwpx(), { filter: file => {
    if (!/^Contents\/(?:section\d+\.xml|content\.hpf|header\.xml)$|^BinData\/[^/]+\.OLE$/i.test(file.name)) return false;
    if (file.originalSize > 32 * 1024 * 1024 || (size += file.originalSize) > 96 * 1024 * 1024) throw Error('HWP conversion metadata too large');
    return true;
  } });
}
