'use strict';

// Package the already verified, branded Windows application without rebuilding
// its runtime or changing the executable's resources.
module.exports = {
  appId: 'local.wordstudio.app',
  productName: 'Word Studio',
  executableName: 'Word Studio',
  directories: { output: 'release/installer' },
  publish: null,
  win: {
    target: [{ target: 'nsis', arch: ['x64'] }],
    icon: 'electron/WordStudio.ico',
    signAndEditExecutable: false,
  },
  nsis: {
    artifactName: 'Word-Studio-${version}-windows-${arch}-setup.${ext}',
    oneClick: false,
    perMachine: false,
    allowToChangeInstallationDirectory: true,
    createDesktopShortcut: true,
    createStartMenuShortcut: true,
    shortcutName: 'Word Studio',
    uninstallDisplayName: 'Word Studio',
    installerIcon: 'electron/WordStudio.ico',
    uninstallerIcon: 'electron/WordStudio.ico',
    installerLanguages: ['ko_KR', 'en_US'],
    displayLanguageSelector: true,
    runAfterFinish: true,
    deleteAppDataOnUninstall: false,
  },
};
