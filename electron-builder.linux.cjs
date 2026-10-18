'use strict';

module.exports = {
  appId: 'local.wordstudio.app',
  productName: 'Word Studio',
  executableName: 'word-studio',
  files: ['dist/**/*', 'electron/**/*', 'package.json'],
  asar: true,
  publish: null,
  linux: {
    target: [{ target: 'deb', arch: ['x64'] }],
    category: 'Office;WordProcessor',
    maintainer: 'Word Studio contributors',
    vendor: 'Word Studio',
    syncDesktopName: true,
    desktop: { entry: { 'Name[ko]': 'Word Studio', 'Comment[ko]': '로컬 문서 편집기' } },
  },
  deb: {
    artifactName: 'Word-Studio-${version}-linux-x64.deb',
    packageName: 'word-studio',
    packageCategory: 'editors',
    recommends: [],
    depends: ['libc6 (>= 2.35)', 'libgtk-3-0', 'libnotify4', 'libnss3',
      'libxss1', 'libxtst6', 'xdg-utils', 'libatspi2.0-0', 'libuuid1',
      'libsecret-1-0', 'libasound2 | libasound2t64', 'libgbm1', 'libdrm2',
      'libx11-6', 'libxcomposite1', 'libxdamage1', 'libxfixes3', 'libxrandr2',
      'libxcb1', 'libxkbcommon0'],
  },
};
