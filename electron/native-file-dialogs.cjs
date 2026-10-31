'use strict';

const path = require('node:path');

// defaultPath alone does not stop AppKit from restoring an iCloud bookmark.
// Write only this application's preference, as a plain path, before creating
// either panel. Never read/delete the shared recent-folder or bookmark records.
function createNativeFileDialogs({ app, dialog, systemPreferences, platform = process.platform }) {
  function prepare(kind, options) {
    if (platform !== 'darwin') return;
    const selected = options.defaultPath;
    const directory = typeof selected === 'string' && path.isAbsolute(selected)
      ? (kind === 'save' ? path.dirname(selected) : selected) : app.getPath('documents');
    if (typeof directory !== 'string' || !path.isAbsolute(directory)) throw new Error('파일 선택창의 시작 폴더를 확인하지 못했습니다.');
    systemPreferences.setUserDefault('NSOSPLastRootDirectory', 'string', directory);
  }
  return {
    open(parent, options) { prepare('open', options); return dialog.showOpenDialog(parent, options); },
    save(parent, options) { prepare('save', options); return dialog.showSaveDialog(parent, options); },
  };
}
module.exports = { createNativeFileDialogs };
