// Copyright (c) 2026 AbeLin · MIT License
const { spawnSync } = require('node:child_process');
const { existsSync, mkdirSync } = require('node:fs');
const { join, resolve } = require('node:path');

const desktop = resolve(__dirname, '..');
const project = resolve(desktop, '..');
if (process.platform !== 'win32' || process.arch !== 'x64') {
  throw new Error('当前打包入口面向 Windows x64，请在对应的 Windows 机器上构建。');
}
const venv = join(project, '.venv/Scripts/python.exe');
const python = process.env.EXCELROUTER_PYTHON || (existsSync(venv) ? venv : 'python');
function run(executable, args, options = {}) {
  const result = spawnSync(executable, args, {
    cwd: project, windowsHide: true, shell: false,
    env: { ...process.env, PYTHONUTF8: '1', PYTHONIOENCODING: 'utf-8' },
    ...options,
  });
  if (result.error) throw result.error;
  if (result.status !== 0) throw new Error(result.stderr?.toString().trim() || `打包命令失败，退出码 ${result.status}`);
  return result;
}

try {
  run(python, ['-c', 'import struct, PyInstaller; from desktop.runtime_libraries import environment; environment(); assert struct.calcsize("P") == 8, "Please use 64-bit Python"'], { encoding: 'utf8' });
} catch (error) {
  console.error('请先在项目根目录安装打包依赖：.venv\\Scripts\\python.exe -m pip install -r desktop/requirements-build.txt');
  throw error;
}

// 专用目录避免覆盖根目录原 Python GUI 的 dist / build。
const staging = join(desktop, 'build/backend');
mkdirSync(staging, { recursive: true });
run(python, [
  '-m', 'PyInstaller', '--noconfirm', '--clean', '--onedir', '--console',
  '--name', 'EasyExcelBackend', '--icon', join(desktop, 'assets/app.ico'),
  '--distpath', join(staging, 'dist'), '--workpath', join(staging, 'work'), '--specpath', staging,
  '--paths', project, '--hidden-import', 'core.collect', '--hidden-import', 'xlrd',
  '--hidden-import', 'matplotlib.backends.backend_agg', '--hidden-import', 'pandas.plotting._matplotlib',
  '--hidden-import', 'win32timezone', '--hidden-import', 'PIL.ImageGrab',
  '--hidden-import', 'statistics', '--hidden-import', 'sqlite3', '--hidden-import', 'fractions',
  '--hidden-import', 'decimal', '--hidden-import', 'csv', '--hidden-import', 'random',
  '--hidden-import', 'statsmodels.api', '--hidden-import', 'statsmodels.formula.api',
  '--hidden-import', 'scipy.optimize', '--hidden-import', 'scipy.stats', '--hidden-import', 'scipy.special',
  '--hidden-import', 'babel.numbers', '--collect-data', 'babel',
  '--exclude-module', 'lxml',
  '--exclude-module', 'pyarrow', '--exclude-module', 'numba', '--exclude-module', 'IPython',
  '--exclude-module', 'PyQt5', '--exclude-module', 'PyQt6', '--exclude-module', 'PySide6', '--exclude-module', 'PySide2',
  '--exclude-module', 'tkinter', '--exclude-module', 'customtkinter',
  '--exclude-module', 'pytest', '--exclude-module', '_pytest',
  join(desktop, 'bridge.py'),
], { stdio: 'inherit' });

// 保留 console 子系统以使用 stdin/stdout 协议，Electron 通过 windowsHide 隐藏后台窗口。
run(python, [join(desktop, 'scripts/runtime_metadata.py'), join(staging, 'dist/EasyExcelBackend')], { stdio: 'inherit' });
const backend = join(staging, 'dist/EasyExcelBackend/EasyExcelBackend.exe');
const checked = run(backend, [], { encoding: 'utf8', input: '{"command":"health"}\n', timeout: 30000 });
const event = JSON.parse(checked.stdout.trim());
if (event.type !== 'result' || event.result.frozen !== true) throw new Error('内置 Python 后端自检失败');
console.log(`Python 后端已打包并通过自检：${backend}`);
const environment = run(backend, [], { encoding: 'utf8', input: '{"command":"code_environment"}\n', timeout: 30000 });
console.log(`离线脚本环境：${environment.stdout.trim()}`);
