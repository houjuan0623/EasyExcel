// Copyright (c) 2026 AbeLin · MIT License
const { _electron: electron } = require('playwright');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const { join, resolve } = require('node:path');
const { tmpdir } = require('node:os');
const { spawnSync } = require('node:child_process');
const root = resolve(__dirname, '..');
const parent = process.env.EXCELROUTER_TEST_DIR || tmpdir();
fs.mkdirSync(parent, { recursive: true });
const scratch = fs.mkdtempSync(join(parent, 'native-ui-'));
const python = join(root, '../.venv/Scripts/python.exe');
const fixture = spawnSync(python, [join(__dirname, 'native_preview_fixture.py'), scratch], { encoding: 'utf8', windowsHide: true, env: { ...process.env, PYTHONIOENCODING: 'utf-8' } });
assert.equal(fixture.status, 0, fixture.stderr);
const source = JSON.parse(fixture.stdout).path;
const packaged = process.argv.includes('--packaged');
const env = { ...process.env, EXCELROUTER_USER_DATA: join(scratch, 'user-data'), EXCELROUTER_PYTHON: python };
for (const key of Object.keys(env)) if (['electron_run_as_node', 'electron_renderer_url'].includes(key.toLowerCase())) delete env[key];
if (packaged) {
  for (const key of Object.keys(env)) if (['path', 'pythonpath', 'pythonhome'].includes(key.toLowerCase())) delete env[key];
  env.Path = join(process.env.SystemRoot || 'C:/Windows', 'System32');
  env.EXCELROUTER_PYTHON = join(scratch, 'no-python.exe');
}
const launch = packaged ? { executablePath: process.env.EXCELROUTER_PACKAGED_EXE, args: [], cwd: scratch } : { args: [root] };
async function main() {
  const app = await electron.launch({ ...launch, env, timeout: 45000 });
  try {
    const page = await app.firstWindow(); page.setDefaultTimeout(90000);
    const errors = []; page.on('pageerror', error => errors.push(error.message));
    await page.getByText('本地引擎已就绪', { exact: true }).waitFor();
    await app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows()[0].setSize(1500, 1000));
    await page.getByRole('button', { name: '统计分析', exact: true }).click();
    await app.evaluate(({ dialog }, file) => { dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [file] }); }, source);
    await page.getByRole('button', { name: '选择分析文件', exact: true }).click();
    await page.getByLabel('数据块 1 区域', { exact: true }).fill('A3:C20');
    await page.getByRole('button', { name: '选择数据块 1 区域', exact: true }).click();
    const picker = page.getByRole('dialog', { name: '选择分析数据区域', exact: true });
    const cell = address => picker.locator(`[data-analysis-cell="${address}"]`);
    async function rendered(address) {
      await cell(address).waitFor();
      await page.waitForFunction(a => { const el = document.querySelector(`[data-analysis-cell="${a}"]`); return el && !el.disabled; }, address);
      assert.equal(await picker.locator('.analysis-error').count(), 0);
    }
    async function jump(row, col) {
      await picker.getByLabel('选区预览起始行', { exact: true }).fill(String(row));
      await picker.getByLabel('选区预览起始列', { exact: true }).fill(col);
      await picker.getByRole('button', { name: '跳转', exact: true }).click();
    }
    await rendered('A1');
    assert.equal(await picker.getByLabel('选区渲染方式').inputValue(), 'native');
    assert.equal(await picker.locator('.native-sheet-image').count(), 1);
    assert.equal(await cell('D4').count(), 0);
    assert.equal(await cell('B7').count(), 0);
    await cell('A1').click();
    assert.equal(await picker.getByLabel('所选分析区域').inputValue(), 'A1:F2');
    await cell('C4').click();
    assert.equal(await picker.getByLabel('单元格公式或内容').inputValue(), '=B4*2');
    await cell('B4').click();
    assert.equal(await picker.getByLabel('单元格公式或内容').inputValue(), '400.00');
    const start = await cell('A3').boundingBox(), end = await cell('C10').boundingBox();
    await page.mouse.move(start.x + 10, start.y + 5); await page.mouse.down();
    await page.mouse.move(end.x + 10, end.y + 5, { steps: 8 }); await page.mouse.up();
    assert.equal(await picker.getByLabel('所选分析区域').inputValue(), 'A3:C10');
    await picker.getByLabel('所选分析区域').fill('');
    await page.screenshot({ path: join(scratch, '01-native-view.png') });
    await picker.getByRole('button', { name: '选择整列 A', exact: true }).click();
    await picker.getByRole('button', { name: '选择整列 C', exact: true }).click({ modifiers: ['Shift'] });
    assert.equal(await picker.getByLabel('所选分析区域').inputValue(), 'A:C');
    await jump(101, 'A'); await rendered('C121');
    await picker.getByRole('button', { name: '放大选区预览', exact: true }).click();
    await picker.getByRole('button', { name: '放大选区预览', exact: true }).click();
    await cell('C121').click();
    assert.equal(await picker.getByLabel('单元格公式或内容').inputValue(), '=B121*2');
    assert.equal(await picker.getByLabel('所选分析区域').inputValue(), 'C121:C121');
    await picker.getByRole('button', { name: '选择整行 121', exact: true }).click();
    assert.equal(await picker.getByLabel('所选分析区域').inputValue(), '121:121');
    await jump(101, 'U'); await rendered('AB121'); await cell('AB121').click();
    assert.equal(await picker.getByLabel('单元格公式或内容').inputValue(), '第 28 列');
    await jump(1, 'A'); await rendered('C4');
    await picker.getByLabel('显示公式原文', { exact: true }).check();
    await rendered('C4'); await cell('C4').click();
    assert.equal(await picker.getByLabel('单元格公式或内容').inputValue(), '=B4*2');
    await page.screenshot({ path: join(scratch, '02-formula-view.png') });
    await picker.getByLabel('所选分析区域').fill('A3:C20');
    await picker.getByRole('button', { name: '使用此区域', exact: true }).click();
    assert.equal(await page.getByLabel('数据块 1 区域', { exact: true }).inputValue(), 'A3:C20');
    await page.getByRole('button', { name: '读取数据并检查字段', exact: true }).click();
    await page.getByLabel('数值字段 预算', { exact: true }).check();
    await page.getByRole('button', { name: '生成图表预览', exact: true }).click();
    await page.locator('.analysis-live-preview img').waitFor();
    assert.deepEqual(errors, []);
    console.log(JSON.stringify({ ok: true, packaged, nativeExcel: true, mergedSelection: true, formulas: true, nativeCharts: true, afterRow100: true, afterColumn20: true, zoomSelection: true, downstreamAnalysis: true, scratch }));
  } catch (error) { fs.writeFileSync(join(scratch, 'error.txt'), String(error.stack)); throw error; }
  finally { await app.close(); }
}
main().catch(error => { console.error(error); process.exitCode = 1; });
