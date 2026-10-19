// Copyright (c) 2026 AbeLin · MIT License
import { execFile, spawn, type ChildProcessWithoutNullStreams } from 'node:child_process';
import { existsSync } from 'node:fs';
import { join, resolve } from 'node:path';
import { createInterface } from 'node:readline';
import { app } from 'electron';
import type { CodeOutput } from '../shared/types';

export const projectRoot = app.isPackaged ? process.resourcesPath : resolve(__dirname, '../../..');
export const desktopRoot = app.isPackaged ? app.getAppPath() : join(projectRoot, 'desktop');
export type PythonEvent = { type: string; message?: string; value?: number; result?: unknown; data?: CodeOutput; output?: string };
const children = new Set<ChildProcessWithoutNullStreams>();

export function pythonExecutable(): string {
  if (app.isPackaged) return join(process.resourcesPath, 'backend', process.platform === 'win32' ? 'EasyExcelBackend.exe' : 'EasyExcelBackend');
  if (process.env.EXCELROUTER_PYTHON) return process.env.EXCELROUTER_PYTHON;
  const venv = join(projectRoot, '.venv', process.platform === 'win32' ? 'Scripts/python.exe' : 'bin/python');
  return existsSync(venv) ? venv : process.platform === 'win32' ? 'python' : 'python3';
}

export function callPython<T>(request: object, onEvent?: (event: PythonEvent) => void) {
  // JSON 通过 stdin 传入，路径、列名和单元格内容从不拼入命令行。
  const env: NodeJS.ProcessEnv = { ...process.env, PYTHONUTF8: '1', PYTHONIOENCODING: 'utf-8' };
  if (app.isPackaged) {
    delete env.PYTHONHOME;
    delete env.PYTHONPATH;
  }
  const child = spawn(pythonExecutable(), app.isPackaged ? [] : ['-u', join(desktopRoot, 'bridge.py')], {
    cwd: app.isPackaged ? app.getPath('userData') : projectRoot,
    windowsHide: true,
    shell: false,
    env,
    stdio: ['pipe', 'pipe', 'pipe'],
  });
  children.add(child);
  const promise = new Promise<T>((resolveResult, reject) => {
    let stderr = '';
    let received = false;
    let result: T;
    let failure: Error | undefined;
    child.stderr.setEncoding('utf8');
    child.stderr.on('data', (data: string) => { stderr = (stderr + data).slice(-6000); });
    const lines = createInterface({ input: child.stdout });
    lines.on('line', (line) => {
      try {
        const event = JSON.parse(line) as PythonEvent;
        if (event.type === 'result') { received = true; result = event.result as T; }
        else if (event.type === 'error') failure = new Error(event.message || 'Python 处理失败');
        else onEvent?.(event);
      } catch { failure = new Error('Python 返回了无法识别的数据，请查看运行环境'); }
    });
    child.on('error', () => reject(new Error(app.isPackaged
      ? '无法启动内置表格引擎。请确认安装目录内的 backend 文件完整，或重新安装应用。'
      : '无法启动 Python。请安装 Python 3.9+ 和 desktop/requirements.txt，或设置 EXCELROUTER_PYTHON。')));
    child.stdin.on('error', () => { /* 退出原因由 close 统一报告，避免 EPIPE 崩溃。 */ });
    child.on('close', (code) => {
      lines.close();
      children.delete(child);
      if (failure) reject(failure);
      else if (received && code === 0) resolveResult(result!);
      else reject(new Error(stderr.includes('ModuleNotFoundError')
        ? app.isPackaged ? '内置表格引擎缺少依赖，请重新安装应用。' : 'Python 缺少依赖，请用当前解释器安装 desktop/requirements.txt 中的依赖。'
        : stderr.trim() || 'Python 处理进程已退出，请重试'));
    });
    child.stdin.write(`${JSON.stringify(request)}\n`);
  });
  return { child, promise };
}

export function terminatePython(child: ChildProcessWithoutNullStreams) {
  if (child.exitCode !== null || child.signalCode !== null) return;
  if (process.platform === 'win32' && child.pid) {
    // Only a child owned by this app; also stop subprocesses created by a script.
    execFile('taskkill', ['/pid', String(child.pid), '/T', '/F'], { windowsHide: true }, (error) => {
      if (error && child.exitCode === null) child.kill();
    });
  } else child.kill('SIGKILL');
}

export function stopPython() {
  for (const child of children) terminatePython(child);
  children.clear();
}
