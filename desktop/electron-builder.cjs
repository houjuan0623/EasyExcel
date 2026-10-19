// Copyright (c) 2026 AbeLin · MIT License
const { dirname } = require('node:path');

module.exports = {
  appId: 'com.excelrouter.desktop',
  productName: 'EasyExcel',
  copyright: 'Copyright (c) 2026 AbeLin',
  directories: { output: 'release', buildResources: '../packaging' },
  asar: true,
  // React 等依赖已由 Vite 写入 dist；主进程只引用 Electron / Node 内置模块。
  // 使用官方 hook 声明依赖已在构建阶段处理，不再收集运行时 node_modules。
  beforeBuild: async () => false,
  electronDist: dirname(require('electron')),
  files: ['dist/**/*', 'dist-electron/**/*', 'package.json', '!node_modules/**/*'],
  extraResources: [
    { from: 'build/backend/dist/EasyExcelBackend', to: 'backend', filter: ['**/*'] },
    { from: 'assets/app.png', to: 'app.png' },
    { from: 'assets/app.ico', to: 'app.ico' },
    { from: '../LICENSE', to: 'LICENSE.txt' },
  ],
  win: {
    target: [{ target: 'nsis', arch: ['x64'] }],
    icon: 'assets/app.ico',
    executableName: 'EasyExcel',
    artifactName: 'EasyExcel-Setup-${version}-${arch}.${ext}',
  },
  nsis: {
    oneClick: false,
    perMachine: false,
    allowElevation: false,
    allowToChangeInstallationDirectory: true,
    createDesktopShortcut: true,
    createStartMenuShortcut: true,
    shortcutName: 'EasyExcel',
    installerLanguages: ['zh_CN', 'en_US'],
    runAfterFinish: false,
    deleteAppDataOnUninstall: false,
  },
  publish: null,
};
