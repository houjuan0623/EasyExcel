// Copyright (c) 2026 AbeLin · MIT License
import { app, BrowserWindow, clipboard, dialog, ipcMain, shell, type IpcMainInvokeEvent } from 'electron';
import { randomUUID } from 'node:crypto';
import { existsSync, mkdirSync, readFileSync, statSync, writeFileSync } from 'node:fs';
import { basename, join, resolve } from 'node:path';
import type { ChildProcessWithoutNullStreams } from 'node:child_process';
import type { CodeEvent, CodeRequest, CodeResult, HistoryItem, JobEvent, JobResult, RunRequest } from '../shared/types';
import { callPython, desktopRoot, stopPython, terminatePython } from './python';
import { ScriptLibrary } from './script-library';

let window: BrowserWindow;
let active: { id: string; child: ChildProcessWithoutNullStreams; cancel?: () => void } | null = null;
let quitting = false;
let history: HistoryItem[] = [];
let scriptLibrary: ScriptLibrary;
let libraryWarning = '';
const allowedOutputs = new Set<string>();
const reads = new Map<string, ChildProcessWithoutNullStreams>();
const devUrl = app.isPackaged ? undefined : process.env.ELECTRON_RENDERER_URL;
if (devUrl && new URL(devUrl).origin !== 'http://127.0.0.1:5173') throw new Error('仅允许本机开发服务器');
app.setName('EasyExcel');
// 品牌更名后仍使用原目录，保留脚本库、历史和 Chromium 中的编辑草稿。
const userData = process.env.EXCELROUTER_USER_DATA || join(app.getPath('appData'), 'excelrouter-desktop');
mkdirSync(userData, { recursive: true });
app.setPath('userData', userData);
app.setPath('sessionData', userData);
if (process.platform === 'win32') app.setAppUserModelId('com.excelrouter.desktop');

function trusted(event: IpcMainInvokeEvent) {
  if (!window || event.sender !== window.webContents || event.senderFrame !== window.webContents.mainFrame) {
    throw new Error('不允许的界面请求');
  }
}

function handle(channel: string, fn: (...args: any[]) => unknown) {
  ipcMain.handle(channel, (event, ...args: unknown[]) => { trusted(event); return fn(...args); });
}

function send(event: JobEvent) {
  if (!window.isDestroyed()) window.webContents.send('job:event', event);
}

function saveHistory() {
  writeFileSync(join(app.getPath('userData'), 'history.json'), JSON.stringify(history.slice(0, 20), null, 2), 'utf8');
}

function sendCode(event: CodeEvent) {
  if (!window.isDestroyed()) window.webContents.send('code:event', event);
}

function textContent(value: unknown): string {
  if (typeof value !== 'string' || value.length > 200000) throw new Error('文本过长，最多 200,000 字符');
  return value;
}

function registerHandlers() {
  handle('python:health', () => callPython({ command: 'health' }).promise);
  handle('code:environment', () => callPython({ command: 'code_environment' }).promise);
  handle('scripts:list', () => {
    const snapshot = scriptLibrary.list();
    if (libraryWarning) snapshot.warnings.push(libraryWarning);
    return snapshot;
  });
  handle('scripts:get', (id: string) => scriptLibrary.get(id));
  handle('scripts:save', (request) => scriptLibrary.save(request));
  handle('scripts:update', (id: string, changes) => scriptLibrary.update(id, changes));
  handle('scripts:export', async () => {
    const content = scriptLibrary.exportBackup();
    const result = await dialog.showSaveDialog(window, { title: '备份脚本库', defaultPath: 'EasyExcel-脚本库.json', filters: [{ name: '脚本库备份', extensions: ['json'] }] });
    if (result.canceled || !result.filePath) return null;
    writeFileSync(result.filePath, content, 'utf8');
    return result.filePath;
  });
  handle('scripts:import', async () => {
    const result = await dialog.showOpenDialog(window, { title: '导入脚本库备份', properties: ['openFile'], filters: [{ name: '脚本库备份', extensions: ['json'] }] });
    if (result.canceled) return null;
    const path = result.filePaths[0];
    if (statSync(path).size > 64 * 1024 * 1024) throw new Error('备份最多 64 MB');
    return scriptLibrary.importBackup(readFileSync(path, 'utf8'));
  });
  handle('code:import', async () => {
    const result = await dialog.showOpenDialog(window, { title: '导入 Python 脚本', properties: ['openFile'], filters: [{ name: 'Python 脚本', extensions: ['py'] }] });
    if (result.canceled) return null;
    const path = result.filePaths[0];
    if (statSync(path).size > 800000) throw new Error('脚本过大，最多 200,000 字符');
    const code = new TextDecoder('utf-8', { fatal: true }).decode(readFileSync(path));
    return { name: basename(path), code: textContent(code) };
  });
  handle('code:save-text', async (kind: unknown, value: unknown) => {
    const content = textContent(value);
    if (kind !== 'script' && kind !== 'instructions') throw new Error('不支持的文件类型');
    const script = kind === 'script';
    const result = await dialog.showSaveDialog(window, {
      title: script ? '保存 Python 脚本' : '导出 AI 编程说明',
      defaultPath: script ? '分析脚本.py' : 'EasyExcel-AI编程说明.txt',
      filters: [{ name: script ? 'Python 脚本' : '文本说明', extensions: [script ? 'py' : 'txt'] }],
    });
    if (result.canceled || !result.filePath) return null;
    writeFileSync(result.filePath, content, 'utf8');
    return result.filePath;
  });
  handle('code:copy-text', (value: unknown) => { clipboard.writeText(textContent(value)); });
  handle('code:start', (request: CodeRequest) => {
    if (active) throw new Error('已有任务正在运行');
    if (!request || !textContent(request.code).trim() || !Number.isInteger(request.timeoutSeconds) || request.timeoutSeconds < 5 || request.timeoutSeconds > 3600) throw new Error('请提供代码，并设置 5～3600 秒的运行上限');
    const id = randomUUID();
    const savedScript = scriptLibrary.recordRun(request.code, request.scriptName || '未命名脚本');
    const started = Date.now();
    let output = '';
    let stopped: 'stopped' | 'timeout' | null = null;
    const run = callPython<CodeResult>({ ...request, command: 'code' }, (event) => {
      if (event.type === 'started' && event.output) { output = event.output; allowedOutputs.add(resolve(output)); }
      if (event.type === 'log' || event.type === 'output') sendCode({ jobId: id, type: event.type, message: event.message, data: event.data });
    });
    const stop = (reason: 'stopped' | 'timeout') => {
      if (active?.id !== id || stopped) return;
      stopped = reason;
      terminatePython(run.child);
    };
    active = { id, child: run.child, cancel: () => stop('stopped') };
    const timer = setTimeout(() => stop('timeout'), request.timeoutSeconds * 1000);
    const finish = (result: CodeResult) => {
      clearTimeout(timer);
      if (active?.id === id) active = null;
      try { scriptLibrary.finishRun(savedScript.id, result.status); }
      catch { sendCode({ jobId: id, type: 'log', message: '脚本已收录，但无法更新脚本库中的运行状态。' }); }
      if (result.output) {
        allowedOutputs.add(resolve(result.output));
        history.unshift({ id, name: basename(request.inputPath), mode: 'code', date: new Date().toISOString(), output: result.output,
          summary: { rows: result.rows, files: result.artifactCount, status: result.status, scriptId: savedScript.id }, stopped: result.status === 'stopped' || result.status === 'timeout' });
        history = history.slice(0, 20);
        try { saveHistory(); } catch { sendCode({ jobId: id, type: 'log', message: '无法保存最近任务记录；结果已保存在输出目录。' }); }
      }
      sendCode({ jobId: id, type: 'result', result });
    };
    void run.promise.then((result) => finish(stopped ? { ...result, status: stopped } : result)).catch((error: Error) => finish({
      output, status: stopped || 'error', error: stopped ? undefined : error.message,
      artifacts: [], artifactCount: 0, rows: 0, duration: (Date.now() - started) / 1000,
    }));
    return id;
  });
  handle('code:cancel', (id: string) => { if (active?.id === id) active.cancel?.(); });
  handle('dialog:source', async (kind: unknown) => {
    if (kind !== 'file' && kind !== 'directory') throw new Error('不支持的选择方式');
    const result = await dialog.showOpenDialog(window, {
      title: kind === 'file' ? '选择 Excel 工作簿' : '选择 Excel 文件夹',
      properties: [kind === 'file' ? 'openFile' : 'openDirectory'],
      filters: [{ name: 'Excel 工作簿', extensions: ['xlsx', 'xls'] }],
    });
    return result.canceled ? null : result.filePaths[0];
  });
  handle('dialog:output', async () => {
    const result = await dialog.showOpenDialog(window, { title: '选择结果保存位置', properties: ['openDirectory', 'createDirectory'] });
    return result.canceled ? null : result.filePaths[0];
  });
  for (const command of ['inspect', 'preview', 'groups']) {
    handle(`excel:${command}`, async (request: Record<string, unknown>) => {
      if (!request || typeof request !== 'object') throw new Error('请求格式无效');
      // 新选择替换旧查询，避免快速切换 Sheet 时积压 Python 子进程。
      reads.get(command)?.kill();
      const run = callPython({ ...request, command });
      reads.set(command, run.child);
      try { return await run.promise; }
      finally { if (reads.get(command) === run.child) reads.delete(command); }
    });
  }
  handle('job:start', (request: RunRequest) => {
    if (active) throw new Error('已有任务正在运行');
    if (!request || !['split', 'merge'].includes(request.mode)) throw new Error('请求格式无效');
    const id = randomUUID();
    const run = callPython<JobResult>({ ...request, command: 'run' }, (event) => {
      if (event.type === 'log' || event.type === 'progress') send({ type: event.type, jobId: id, message: event.message, value: event.value });
    });
    active = { id, child: run.child };
    run.promise.then((result) => {
      if (result.output) {
        allowedOutputs.add(resolve(result.output));
        history.unshift({ ...result, id, name: basename(request.inputPath), mode: request.mode, date: new Date().toISOString() });
        history = history.slice(0, 20);
        try { saveHistory(); } catch { send({ jobId: id, type: 'log', message: '结果已生成，但无法保存最近任务记录。' }); }
      }
      send({ jobId: id, type: 'result', result });
    }).catch((error: Error) => send({ jobId: id, type: 'error', message: error.message }))
      .finally(() => { if (active?.id === id) active = null; });
    return id;
  });
  handle('job:cancel', (id: string) => {
    if (active?.id !== id) return;
    active.child.stdin.write('{"cancel":true}\n');
  });
  handle('history:list', () => history);
  handle('output:open', async (path: unknown) => {
    if (typeof path !== 'string' || !allowedOutputs.has(resolve(path))) throw new Error('请选择已生成的结果');
    if (!existsSync(path)) throw new Error('结果已被移动或删除');
    if (statSync(path).isDirectory()) {
      const error = await shell.openPath(path);
      if (error) throw new Error(error);
    } else shell.showItemInFolder(path);
  });
  ipcMain.on('window:action', (event, action: string) => {
    if (event.sender !== window.webContents || event.senderFrame !== window.webContents.mainFrame) return;
    if (action === 'minimize') window.minimize();
    if (action === 'maximize') window.isMaximized() ? window.unmaximize() : window.maximize();
    if (action === 'close') window.close();
  });
}

function createWindow() {
  window = new BrowserWindow({
    width: 1440, height: 960, minWidth: 1040, minHeight: 720,
    backgroundColor: '#f6f8fa', title: 'EasyExcel', frame: false, show: false,
    icon: join(app.isPackaged ? process.resourcesPath : join(desktopRoot, 'assets'), process.platform === 'win32' ? 'app.ico' : 'app.png'),
    webPreferences: {
      preload: join(__dirname, 'preload.js'), contextIsolation: true,
      nodeIntegration: false, sandbox: true, webSecurity: true,
    },
  });
  window.webContents.setWindowOpenHandler(() => ({ action: 'deny' }));
  window.webContents.on('will-navigate', (event) => event.preventDefault());
  window.webContents.session.setPermissionRequestHandler((_contents, _permission, callback) => callback(false));
  window.webContents.on('will-attach-webview', (event) => event.preventDefault());
  window.once('ready-to-show', () => window.show());
  window.on('close', (event) => {
    if (!active || quitting) return;
    event.preventDefault();
    void dialog.showMessageBox(window, {
      type: 'question', title: '任务正在处理', message: '关闭窗口会中断当前任务，可能留下不完整的结果。',
      buttons: ['继续处理', '关闭并中断'], defaultId: 0, cancelId: 0,
    }).then(({ response }) => { if (response === 1) { quitting = true; app.quit(); } });
  });
  if (devUrl) void window.loadURL(devUrl);
  else void window.loadFile(join(desktopRoot, 'dist/index.html'));
}

void app.whenReady().then(() => {
  try {
    const stored = JSON.parse(readFileSync(join(app.getPath('userData'), 'history.json'), 'utf8'));
    if (Array.isArray(stored)) history = stored.filter((item) => typeof item?.output === 'string' &&
      typeof item?.date === 'string' && typeof item?.id === 'string' && typeof item?.name === 'string' &&
      ['split', 'merge', 'code'].includes(item?.mode) && item?.summary && typeof item.summary === 'object').slice(0, 20);
  } catch { /* 首次启动或损坏的历史记录不阻止打开界面。 */ }
  history.forEach((item) => allowedOutputs.add(resolve(item.output)));
  scriptLibrary = new ScriptLibrary(join(app.getPath('userData'), 'script-library'));
  try { scriptLibrary.migrateHistory(history); }
  catch { libraryWarning = '部分旧任务脚本未能自动收录，可从旧结果目录导入 script.py；原文件已保留。'; }
  registerHandlers();
  createWindow();
  app.on('activate', () => { if (BrowserWindow.getAllWindows().length === 0) createWindow(); });
});
app.on('will-quit', stopPython);
app.on('window-all-closed', () => { if (process.platform !== 'darwin') app.quit(); });
