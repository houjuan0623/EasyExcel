// Copyright (c) 2026 AbeLin · MIT License
const { _electron: electron } = require('playwright');
const assert = require('node:assert/strict');
const { mkdtempSync, mkdirSync, readFileSync, writeFileSync, existsSync } = require('node:fs');
const { tmpdir } = require('node:os');
const { dirname, join, resolve } = require('node:path');
const { spawnSync } = require('node:child_process');

const root = resolve(__dirname, '..');
const packaged = process.argv.includes('--packaged');
const packagedExecutable = resolve(process.env.EXCELROUTER_PACKAGED_EXE || join(root, 'release', 'win-unpacked', 'EasyExcel.exe'));
const scratchRoot = process.env.EXCELROUTER_TEST_DIR || tmpdir();
mkdirSync(scratchRoot, { recursive: true });
const scratch = mkdtempSync(join(scratchRoot, 'excelrouter-desktop-'));
const python = process.env.EXCELROUTER_PYTHON || join(root, '..', '.venv', process.platform === 'win32' ? 'Scripts/python.exe' : 'bin/python');
function fixture(command) {
  const result = spawnSync(python, [join(__dirname, 'smoke_fixture.py'), command, scratch], { encoding: 'utf8', windowsHide: true, env: { ...process.env, PYTHONIOENCODING: 'utf-8' } });
  assert.equal(result.status, 0, result.stderr);
  return JSON.parse(result.stdout);
}
async function main() {
  const { path } = fixture('create');
  const env = { ...process.env, EXCELROUTER_PYTHON: python, EXCELROUTER_USER_DATA: join(scratch, 'app-data') };
  delete env.ELECTRON_RUN_AS_NODE;
  delete env.ELECTRON_RENDERER_URL;
  if (packaged) {
    // 打包版必须独立运行；故意提供无效解释器并移除 PATH 中的开发工具。
    env.EXCELROUTER_PYTHON = join(scratch, 'no-system-python.exe');
    for (const key of Object.keys(env)) {
      if (['path', 'pythonpath', 'pythonhome'].includes(key.toLowerCase())) delete env[key];
    }
    env.Path = join(process.env.SystemRoot || 'C:\\Windows', 'System32');
  }
  let server;
  if (!packaged && process.env.EXCELROUTER_SMOKE_DEV === '1') {
    const { createServer } = await import('vite');
    server = await createServer({ configFile: join(root, 'vite.config.mts') });
    await server.listen();
    env.ELECTRON_RENDERER_URL = 'http://127.0.0.1:5173';
  }
  const launchOptions = {
    ...(packaged ? { executablePath: packagedExecutable, args: [], cwd: scratch } : { args: [root] }),
    env, timeout: 30000,
  };
  let app = await electron.launch(launchOptions);
  try {
    const page = await app.firstWindow();
    const errors = [];
    page.on('pageerror', (error) => errors.push(error.message));
    page.on('console', (entry) => { if (entry.type() === 'error') errors.push(entry.text()); });
    await page.getByText('本地引擎已就绪', { exact: true }).waitFor();
    assert.equal(await app.evaluate(({ app }) => app.isPackaged), packaged);
    const engine = await page.evaluate(() => window.desktop.health());
    assert.equal(engine.frozen, packaged);
    if (packaged) assert.equal(engine.python, join(dirname(packagedExecutable), 'resources', 'backend', 'EasyExcelBackend.exe'));
    assert.equal(await app.evaluate(({ app }) => app.getName()), 'EasyExcel');
    assert.equal(await page.title(), 'EasyExcel · 表格工作台');
    assert.equal(await page.getByAltText('EasyExcel 图标').evaluate((img) => img.complete && img.naturalWidth > 0), true);
    assert.equal(await page.getByRole('button', { name: '开始拆分', exact: true }).isEnabled(), false);
    await page.screenshot({ path: join(scratch, '01-empty.png') });
    // 只替代系统文件对话框，后续全部通过真实 UI、IPC、Python 与文件系统。
    await app.evaluate(({ dialog }, fixturePath) => {
      dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [fixturePath] });
    }, path);
    await page.getByRole('button', { name: '选择文件', exact: true }).click();
    await page.getByRole('checkbox', { name: '选择 2026年销售明细.xlsx / 10月销售明细', exact: true }).waitFor();
    await page.getByRole('button', { name: '开始拆分', exact: true }).waitFor();
    await page.waitForFunction(() => !document.querySelector('.run-button').disabled);
    assert.equal(await page.getByRole('checkbox', { name: '选择 2026年销售明细.xlsx / 归档数据', exact: true }).isEnabled(), false);
    await page.getByRole('checkbox', { name: '选择 2026年销售明细.xlsx / 9月销售明细', exact: true }).check();
    await page.getByRole('checkbox', { name: '选择 2026年销售明细.xlsx / 10月销售明细', exact: true }).uncheck();
    await page.getByRole('button', { name: '预览 9月销售明细', exact: true }).click();
    await page.waitForFunction(() => document.querySelector('.preview-tab')?.textContent.includes('9月销售明细') && !document.querySelector('.run-button').disabled);
    await page.getByRole('cell', { name: 'SO-090001', exact: true }).waitFor();
    await page.waitForFunction(() => !document.querySelector('.run-button').disabled && !document.querySelector('.group-heading .spin'));
    await page.screenshot({ path: join(scratch, '02-selected.png') });
    await page.locator('.preview-card').scrollIntoViewIfNeeded();
    await page.screenshot({ path: join(scratch, '05-data-preview.png') });
    await page.getByRole('button', { name: '预览 原样预览', exact: true }).click();
    await page.getByRole('cell', { name: '十月销售总览', exact: true }).waitFor();
    assert.equal(await page.locator('.sheet-grid tbody tr').count(), 99);
    assert.equal(await page.locator('[data-cell="F100"]').textContent(), '第 100 行，完整保留');
    assert.equal(await page.getByText('第 101 行，不应显示', { exact: true }).count(), 0);
    assert.equal(await page.getByText('隐藏字段', { exact: true }).count(), 0);
    const titleStyle = await page.locator('[data-cell="A1"]').evaluate((cell) => ({
      colspan: cell.colSpan, rowspan: cell.rowSpan,
      background: getComputedStyle(cell).backgroundColor, color: getComputedStyle(cell).color,
      size: Number.parseFloat(getComputedStyle(cell).fontSize),
    }));
    assert.deepEqual(titleStyle, { colspan: 6, rowspan: 2, background: 'rgb(21, 93, 69)', color: 'rgb(255, 255, 255)', size: 26.67 });
    assert.equal(await page.locator('[data-cell="C5"]').textContent(), '¥1,827.50');
    await page.getByRole('button', { name: '展开预览', exact: true }).click();
    const expanded = page.getByRole('dialog', { name: '原样预览 展开预览', exact: true });
    await expanded.waitFor();
    await page.screenshot({ path: join(scratch, '06-original-preview.png') });
    await expanded.getByRole('button', { name: '放大预览', exact: true }).click();
    assert.equal(await expanded.locator('.sheet-grid').evaluate((table) => getComputedStyle(table).zoom), '1.1');
    await expanded.getByRole('button', { name: '恢复原始比例', exact: true }).click();
    await expanded.locator('[data-cell="F100"]').scrollIntoViewIfNeeded();
    await page.screenshot({ path: join(scratch, '07-original-row100.png') });
    await expanded.getByRole('button', { name: '关闭展开预览', exact: true }).click();
    await page.getByRole('button', { name: '预览 9月销售明细', exact: true }).click();
    await page.getByRole('cell', { name: 'SO-090001', exact: true }).waitFor();
    await page.locator('.sheet-toolbar').getByRole('button', { name: '清空', exact: true }).click();
    assert.equal(await page.getByRole('button', { name: '开始拆分', exact: true }).isEnabled(), false);
    await page.getByRole('checkbox', { name: '选择 2026年销售明细.xlsx / 9月销售明细', exact: true }).check();
    await page.waitForFunction(() => !document.querySelector('.run-button').disabled);
    await page.getByRole('switch', { name: /同时生成总工作簿/ }).check();
    await page.getByRole('button', { name: '开始拆分', exact: true }).click();
    await page.getByText('处理完成，结果已保存', { exact: true }).waitFor({ timeout: 60000 });
    await page.getByText(/总工作簿已生成 · 4 个 Sheet/).waitFor();
    const verified = fixture('verify');
    await page.screenshot({ path: join(scratch, '03-result.png') });
    await page.getByRole('button', { name: 'Excel 汇总', exact: true }).click();
    await page.getByRole('checkbox', { name: '选择 2026年销售明细.xlsx / 10月销售明细', exact: true }).check();
    await page.getByRole('button', { name: '开始汇总', exact: true }).click();
    await page.getByText('处理完成，结果已保存', { exact: true }).waitFor({ timeout: 60000 });
    const history = await page.evaluate(() => window.desktop.history());
    assert.equal(history.length, 2);
    assert.equal(history[0].summary.rows, 48);
    await page.getByRole('button', { name: '开始汇总', exact: true }).click();
    await page.getByRole('button', { name: '停止任务', exact: true }).click();
    await page.getByText('任务已停止', { exact: true }).waitFor({ timeout: 60000 });
    await page.getByRole('button', { name: /最近任务/ }).click();
    assert.equal(await page.locator('.history-row').count(), 2);
    await page.getByRole('button', { name: 'Excel 汇总', exact: true }).click();
    await page.getByLabel('汇总方式', { exact: true }).selectOption('sheets');
    await page.getByRole('button', { name: '开始汇总', exact: true }).click();
    await page.getByText('处理完成，结果已保存', { exact: true }).waitFor({ timeout: 60000 });
    await page.getByText('已整理为 1 个工作簿 · 2 个独立 Sheet', { exact: true }).waitFor();
    const sheetHistory = await page.evaluate(() => window.desktop.history());
    assert.equal(sheetHistory[0].summary.sheets, 2);
    assert.equal(sheetHistory[0].summary.rows, 48);
    assert.ok(existsSync(sheetHistory[0].output));
    await page.locator('.config-card').scrollIntoViewIfNeeded();
    await page.screenshot({ path: join(scratch, '13-sheet-bundle.png') });
    await page.getByRole('button', { name: 'Excel 拆分', exact: true }).click();
    await app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows()[0].setSize(1040, 720));
    await page.waitForFunction(() => window.innerWidth === 1040);
    await page.waitForFunction(() => !document.querySelector('.run-button').disabled);
    assert.equal(await page.evaluate(() => document.documentElement.scrollWidth > window.innerWidth), false);
    await page.screenshot({ path: join(scratch, '04-small-window.png') });
    await page.getByRole('button', { name: 'Python 工作台', exact: true }).click();
    await page.waitForFunction(() => !Array.from(document.querySelectorAll('button')).find((button) => button.textContent === '运行代码')?.disabled);
    assert.equal(await page.evaluate(() => document.documentElement.scrollWidth > window.innerWidth), false);
    await page.screenshot({ path: join(scratch, '08-code-small.png') });
    await app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows()[0].setSize(1440, 960));
    await page.getByRole('button', { name: '运行代码', exact: true }).click();
    await page.getByText('脚本运行完成', { exact: true }).waitFor({ timeout: 60000 });
    assert.ok(await page.locator('.code-output table').count() >= 2);
    const analysisHistory = await page.evaluate(() => window.desktop.history());
    assert.equal(analysisHistory[0].mode, 'code');
    assert.equal(analysisHistory[0].summary.rows, 48);
    assert.ok(existsSync(join(analysisHistory[0].output, '数据概况.xlsx')));

    await page.getByRole('button', { name: '绘制图表', exact: true }).click();
    await page.getByRole('textbox', { name: 'Python 代码', exact: true }).press('Control+Enter');
    await page.getByText('脚本运行完成', { exact: true }).waitFor({ timeout: 60000 });
    await page.locator('.code-output img').waitFor();
    assert.ok(await page.locator('.code-output img').evaluate((img) => img.naturalWidth > 200));
    await page.locator('.code-results').scrollIntoViewIfNeeded();
    await page.screenshot({ path: join(scratch, '09-python-chart.png') });

    // Export a reusable prompt, containing schema but no real cell values.
    const instructions = join(scratch, 'AI编程说明.txt');
    await app.evaluate(({ dialog }, savePath) => { dialog.showSaveDialog = async () => ({ canceled: false, filePath: savePath }); }, instructions);
    await page.getByRole('button', { name: '给 AI 的说明', exact: true }).click();
    await page.getByRole('button', { name: '导出说明 .txt', exact: true }).click();
    await page.waitForFunction(() => document.querySelector('.code-info')?.textContent.includes('已保存'));
    assert.ok(readFileSync(instructions, 'utf8').includes('目标电脑没有网络'));
    assert.ok(!readFileSync(instructions, 'utf8').includes('SO-090001'));
    await page.getByRole('button', { name: '关闭编程说明', exact: true }).click();

    // Import a script as it would arrive from another PC; run actual split + merge helpers.
    const imported = join(scratch, 'AI生成的脚本.py');
    const sourceCode = 'import json, csv, math, statistics, datetime, pathlib, re, collections, itertools, functools, random, decimal, fractions, sqlite3\nmerged = merge_tables()\nsave_excel(merged, "AI合并.xlsx")\npaths = split_table(df, df.columns[1])\ndisplay(merged, "AI分析结果")\nprint("脚本导入成功", len(paths))\n';
    writeFileSync(imported, '\uFEFF' + sourceCode, 'utf8');
    await app.evaluate(({ dialog }, filePath) => { dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [filePath] }); }, imported);
    await page.getByRole('button', { name: '导入 .py', exact: true }).click();
    await page.getByText('脚本已导入，可检查代码后运行。', { exact: true }).waitFor();
    const savedScript = join(scratch, '保存的脚本.py');
    await app.evaluate(({ dialog }, savePath) => { dialog.showSaveDialog = async () => ({ canceled: false, filePath: savePath }); }, savedScript);
    await page.getByRole('button', { name: '保存脚本', exact: true }).click();
    await page.waitForFunction(() => document.querySelector('.code-info')?.textContent.includes('保存的脚本'));
    assert.equal(readFileSync(savedScript, 'utf8'), sourceCode);
    await page.getByRole('button', { name: '运行代码', exact: true }).click();
    await page.getByText('脚本运行完成', { exact: true }).waitFor({ timeout: 60000 });
    await page.getByText('AI分析结果', { exact: true }).waitFor();
    const scriptHistory = await page.evaluate(() => window.desktop.history());
    assert.ok(existsSync(join(scriptHistory[0].output, 'AI合并.xlsx')));
    assert.equal(scriptHistory[0].summary.rows, 48);
    await page.screenshot({ path: join(scratch, '10-python-results.png') });

    const editor = page.getByRole('textbox', { name: 'Python 代码', exact: true });
    async function editCode(text) { await editor.click(); await editor.press('Control+A'); await editor.fill(text); }
    await editCode('print("保留日志")\nraise ValueError("示例报错：字段不存在")');
    await page.getByRole('button', { name: '运行代码', exact: true }).click();
    await page.getByText('脚本运行失败', { exact: true }).waitFor({ timeout: 60000 });
    assert.ok((await page.locator('.code-error').textContent()).includes('line 2'));
    assert.ok((await page.locator('.code-logs').textContent()).includes('保留日志'));
    await page.getByRole('button', { name: '复制运行日志', exact: true }).click();
    assert.ok((await app.evaluate(({ clipboard }) => clipboard.readText())).includes('ValueError: 示例报错：字段不存在'));

    await editCode('print("开始长任务", flush=True)\nwhile True:\n    pass');
    await page.getByRole('button', { name: '运行代码', exact: true }).click();
    assert.equal(await page.getByRole('button', { name: 'Excel 拆分', exact: true }).isEnabled(), false);
    await page.getByRole('button', { name: '运行日志', exact: true }).click();
    await page.waitForFunction(() => document.querySelector('.code-logs')?.textContent.includes('开始长任务'));
    await page.getByRole('button', { name: '停止代码', exact: true }).click();
    await page.getByText('代码已停止', { exact: true }).waitFor({ timeout: 15000 });
    assert.equal(await page.getByRole('button', { name: '运行代码', exact: true }).isEnabled(), true);
    // Test the main-process watchdog with its minimum valid 5-second limit.
    const timeoutResult = await page.evaluate(async (filePath) => {
      const done = new Promise((resolve) => {
        const unsubscribe = window.desktop.onCode((event) => { if (event.type === 'result') { unsubscribe(); resolve(event.result); } });
      });
      await window.desktop.startCode({ inputPath: filePath, outputPath: filePath.replace(/[^\\/]+$/, '') + '代码超时结果', selection: [{ path: filePath, sheets: ['9月销售明细'] }], headerRow: null, code: 'while True:\n    pass', timeoutSeconds: 5 });
      return done;
    }, path);
    assert.equal(timeoutResult.status, 'timeout');
    // Finance capabilities must perform real work, including native extension modules.
    const financeCode = readFileSync(join(__dirname, 'finance_checks.py'), 'utf8');
    await editCode(financeCode);
    await page.getByRole('button', { name: '运行代码', exact: true }).click();
    await page.getByText('脚本运行完成', { exact: true }).waitFor({ timeout: 60000 });
    await page.getByText('财务库验证结果', { exact: true }).waitFor();
    const financeHistory = await page.evaluate(() => window.desktop.history());
    assert.ok(existsSync(join(financeHistory[0].output, '财务库报表.xlsx')));
    const financeScriptId = financeHistory[0].summary.scriptId;
    assert.ok(financeScriptId);
    const financeBefore = await page.evaluate((id) => window.desktop.script(id), financeScriptId);
    assert.equal(financeBefore.runCount, 1);

    await page.getByRole('button', { name: '我的脚本库', exact: true }).click();
    const library = page.getByRole('dialog', { name: '本地脚本库', exact: true });
    await library.getByRole('textbox', { name: '脚本名称', exact: true }).fill('财务综合检查');
    await library.getByRole('combobox', { name: '脚本分类', exact: true }).fill('财务核对');
    await library.getByRole('textbox', { name: '脚本使用说明', exact: true }).fill('离线计算、回归、估值及报表验证');
    await library.getByRole('button', { name: '保存信息', exact: true }).click();
    await library.getByText('脚本信息已保存。', { exact: true }).waitFor();
    await library.getByRole('button', { name: '收藏脚本', exact: true }).click();
    await library.getByRole('button', { name: '取消收藏', exact: true }).waitFor();
    await library.getByRole('button', { name: /我的收藏/ }).click();
    assert.equal(await library.locator('.library-item').count(), 1);
    await library.getByRole('textbox', { name: '搜索脚本', exact: true }).fill('不存在的脚本');
    assert.equal(await library.locator('.library-item').count(), 0);
    await library.getByRole('textbox', { name: '搜索脚本', exact: true }).fill('综合检查');
    await library.getByRole('button', { name: '载入编辑器', exact: true }).waitFor();
    await page.screenshot({ path: join(scratch, '11-script-library.png') });
    await library.getByRole('button', { name: '关闭脚本库', exact: true }).click();

    // Preserve a previously unrun editor draft when loading a library item.
    const unsavedDraft = 'print("未运行草稿，请保留")';
    await editCode(unsavedDraft);
    await page.getByRole('button', { name: '我的脚本库', exact: true }).click();
    await library.getByRole('button', { name: /我的收藏/ }).click();
    await library.getByRole('button', { name: '载入编辑器', exact: true }).click();
    await library.waitFor({ state: 'hidden' });
    // CodeMirror only renders the visible portion of a long document.
    await page.locator('.cm-content').press('Control+End');
    await page.waitForFunction(() => document.querySelector('.cm-content')?.textContent.includes('FINANCE_CHECKS_OK'));
    const draftSaved = await page.evaluate(async (draft) => {
      const items = (await window.desktop.scripts()).items.filter((item) => item.category === '草稿');
      for (const item of items) if ((await window.desktop.script(item.id)).code === draft) return true;
      return false;
    }, unsavedDraft);
    assert.equal(draftSaved, true);
    await page.getByRole('button', { name: '我的脚本库', exact: true }).click();
    await library.getByRole('button', { name: '载入并运行', exact: true }).click();
    await library.waitFor({ state: 'hidden' });
    await page.getByText('脚本运行完成', { exact: true }).waitFor({ timeout: 60000 });
    const financeAfter = await page.evaluate((id) => window.desktop.script(id), financeScriptId);
    assert.equal(financeAfter.runCount, 2);
    assert.equal(financeAfter.name, '财务综合检查');
    assert.equal(financeAfter.favorite, true);

    await page.getByRole('button', { name: '我的脚本库', exact: true }).click();
    await library.getByRole('button', { name: '移入回收站', exact: true }).click();
    await library.locator('.library-scopes').getByRole('button', { name: /回收站/ }).click();
    await library.getByRole('button', { name: '恢复脚本', exact: true }).click();
    await library.getByText('回收站为空', { exact: true }).waitFor();
    await library.getByRole('button', { name: /全部脚本/ }).click();
    const backupFile = join(scratch, '脚本库备份.json');
    await app.evaluate(({ dialog }, savePath) => { dialog.showSaveDialog = async () => ({ canceled: false, filePath: savePath }); }, backupFile);
    await library.getByRole('button', { name: '备份', exact: true }).click();
    await library.getByText(/备份已保存/).waitFor();
    const backup = JSON.parse(readFileSync(backupFile, 'utf8'));
    assert.ok(backup.scripts.some((item) => item.name === '财务综合检查' && item.favorite));
    await app.evaluate(({ dialog }, filePath) => { dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [filePath] }); }, backupFile);
    await library.getByRole('button', { name: '导入备份', exact: true }).click();
    await library.getByText(/已导入 0 份/).waitFor();
    await library.getByRole('button', { name: '关闭脚本库', exact: true }).click();
    const safety = await app.evaluate(({ BrowserWindow }) => {
      const preferences = BrowserWindow.getAllWindows()[0].webContents.getLastWebPreferences();
      return { isolated: preferences.contextIsolation, sandbox: preferences.sandbox, node: preferences.nodeIntegration };
    });
    assert.deepEqual(safety, { isolated: true, sandbox: true, node: false });
    assert.deepEqual(errors, []);
    const beforeRestart = await page.evaluate(() => window.desktop.scripts());
    await app.close();
    app = await electron.launch(launchOptions);
    const restarted = await app.firstWindow();
    await restarted.getByText('本地引擎已就绪', { exact: true }).waitFor();
    const afterRestart = await restarted.evaluate(() => window.desktop.scripts());
    assert.deepEqual(afterRestart.items, beforeRestart.items);
    await restarted.getByRole('button', { name: 'Python 工作台', exact: true }).click();
    await restarted.getByRole('button', { name: '我的脚本库', exact: true }).click();
    await restarted.getByRole('button', { name: /我的收藏/ }).click();
    await restarted.getByRole('button', { name: '载入编辑器', exact: true }).click();
    await restarted.getByRole('button', { name: '关闭脚本库', exact: true }).waitFor({ state: 'hidden' });
    await restarted.locator('.cm-content').press('Control+End');
    await restarted.waitForFunction(() => document.querySelector('.cm-content')?.textContent.includes('FINANCE_CHECKS_OK'));
    assert.equal(await restarted.getByRole('button', { name: '运行代码', exact: true }).isEnabled(), false);
    console.log(JSON.stringify({ ok: true, packaged, engine, scratch, ...verified, mergeRows: 48, rendererErrors: errors }, null, 2));
  } catch (error) {
    console.error(`失败现场：${scratch}`);
    try { await app.windows()[0]?.screenshot({ path: join(scratch, 'failure.png') }); } catch { /* Keep original test failure. */ }
    throw error;
  } finally { await app.close(); if (server) await server.close(); }
}
main().catch((error) => { console.error(error); process.exitCode = 1; });
