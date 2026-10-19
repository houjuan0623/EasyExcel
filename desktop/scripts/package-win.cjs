// Copyright (c) 2026 AbeLin · MIT License
const { existsSync } = require('node:fs');
const { join, resolve } = require('node:path');

// 默认只产生本地未签名产物，不自动使用系统证书或发布到发行平台。
process.env.CSC_IDENTITY_AUTO_DISCOVERY = 'false';
const { build, Platform, Arch } = require('electron-builder');
const project = resolve(__dirname, '..');
if (process.platform !== 'win32' || process.arch !== 'x64') throw new Error('请在 Windows x64 上执行打包命令。');
const backend = join(project, 'build/backend/dist/EasyExcelBackend/EasyExcelBackend.exe');
if (!existsSync(backend)) throw new Error('缺少 Python 后端，请先执行 npm run build:backend。');
const target = process.argv.includes('--zip') ? 'zip' : process.argv.includes('--dir') ? 'dir' : 'nsis';
const config = require('../electron-builder.cjs');
if (target === 'zip') {
  config.win = { ...config.win, artifactName: 'EasyExcel-Portable-${version}-${arch}.${ext}' };
  config.extraFiles = [{ from: 'packaging/portable-readme.txt', to: '使用说明.txt' }];
}

build({
  projectDir: project,
  config,
  targets: Platform.WINDOWS.createTarget(target, Arch.x64),
  publish: 'never',
}).then((artifacts) => {
  console.log('构建完成：');
  for (const artifact of artifacts) console.log(artifact);
  console.log(join(project, 'release/win-unpacked/EasyExcel.exe'));
}).catch((error) => { console.error(error); process.exitCode = 1; });
