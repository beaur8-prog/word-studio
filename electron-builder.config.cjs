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
    // Register supported handlers without replacing the user's default app.
    // electron-builder's APP_ASSOCIATE also overwrites the extension default.
    include: 'electron/windows-file-associations.nsh',
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
    // NSIS chooses the Windows UI language automatically. A forced picker runs
    // again when Setup restarts for elevation, including during updates.
    displayLanguageSelector: false,
    runAfterFinish: true,
    deleteAppDataOnUninstall: false,
  },
};
