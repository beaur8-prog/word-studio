'use strict';

// The reader, OS file-open queue and installers must offer the same formats.
const documentTypes = [
  { name: 'Word Studio Editing Copy', ext: ['wordstudio'], mimeType: 'application/x-wordstudio' },
  { name: 'Word Document', ext: ['docx'], mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', uti: 'org.openxmlformats.wordprocessingml.document' },
  { name: 'Hangul Document', ext: ['hwp'], mimeType: 'application/x-hwp' },
  { name: 'Hangul XML Document', ext: ['hwpx'], mimeType: 'application/hwp+zip' },
  { name: 'HTML Document', ext: ['html', 'htm'], mimeType: 'text/html', uti: 'public.html' },
  { name: 'Text Document', ext: ['txt'], mimeType: 'text/plain', uti: 'public.plain-text' },
];
const OPEN_EXTENSIONS = new Set(documentTypes.flatMap(type => type.ext));
const fileAssociations = documentTypes.flatMap(({ name, ext, mimeType }) =>
  ext.map(extension => ({ name, ext: extension, mimeType, role: 'Editor', rank: 'Alternate' })));
const macDocumentTypes = documentTypes.map(({ name, ext, uti }) => ({
  CFBundleTypeName: name, CFBundleTypeExtensions: ext,
  CFBundleTypeRole: 'Editor', LSHandlerRank: 'Alternate',
  ...(uti ? { LSItemContentTypes: [uti] } : {}),
}));

module.exports = { OPEN_EXTENSIONS, fileAssociations, macDocumentTypes };
