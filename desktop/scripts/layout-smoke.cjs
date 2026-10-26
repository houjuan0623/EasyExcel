// Copyright (c) 2026 AbeLin · MIT License
// 由 analysis-smoke 调用，对真实 Electron 鼠标坐标、缩放和导出报告执行集成检查。
const assert = require('node:assert/strict');
const { join } = require('node:path');
const { pathToFileURL } = require('node:url');
const plan = page => page.evaluate(() => JSON.parse(localStorage.getItem('easyexcel.dashboard.v1')));
const close = (a, b, tolerance = 2) => assert.ok(Math.abs(a - b) <= tolerance, `${a} != ${b}`);

async function layoutInteractions(app, page, scratch) {
  const board = page.locator('.analysis-board'), canvas = board.locator('.dashboard-artboard');
  const before = await plan(page), id = before.cards[0].id;
  await page.getByLabel('布局图表', { exact: true }).selectOption(id);
  await page.getByLabel('吸附网格', { exact: true }).uncheck();
  const card = board.locator(`[data-card-id="${id}"]`);
  await card.scrollIntoViewIfNeeded();
  const dimensions = await canvas.boundingBox(); close(dimensions.width / dimensions.height, 16 / 9, 0.001);
  const scale = dimensions.width / 1920;
  const header = await card.locator('.dashboard-tile-header').boundingBox();
  await page.mouse.move(header.x + 20, header.y + header.height / 2); await page.mouse.down();
  await page.mouse.move(header.x + 20 + 72 * scale, header.y + header.height / 2 + 48 * scale, { steps: 7 }); await page.mouse.up();
  const moved = (await plan(page)).cards[0].layout;
  close(moved.x, before.cards[0].layout.x + 72); close(moved.y, before.cards[0].layout.y + 48);
  await card.locator('[data-handle="se"]').scrollIntoViewIfNeeded();
  const handle = await card.locator('[data-handle="se"]').boundingBox();
  await page.mouse.move(handle.x + handle.width / 2, handle.y + handle.height / 2); await page.mouse.down();
  await page.mouse.move(handle.x + handle.width / 2 + 120 * scale, handle.y + handle.height / 2 + 60 * scale, { steps: 7 }); await page.mouse.up();
  const resized = (await plan(page)).cards[0].layout;
  close(resized.width, moved.width + 120); close(resized.height, moved.height + 60);
  await page.getByRole('button', { name: '撤销布局', exact: true }).click();
  assert.deepEqual((await plan(page)).cards[0].layout, moved);
  await page.getByRole('button', { name: '重做布局', exact: true }).click();
  assert.deepEqual((await plan(page)).cards[0].layout, resized);
  // Esc 放弃尚未结束的手势，不将半成品写进本机草稿。
  await card.scrollIntoViewIfNeeded();
  const again = await card.locator('.dashboard-tile-header').boundingBox();
  await page.mouse.move(again.x + 20, again.y + 8); await page.mouse.down();
  await page.mouse.move(again.x + 80, again.y + 40); await page.keyboard.press('Escape'); await page.mouse.up();
  assert.deepEqual((await plan(page)).cards[0].layout, resized);
  for (const [index, rect] of [[0, [96, 144, 540, 216]], [1, [24, 408, 924, 648]], [2, [972, 408, 924, 648]]]) {
    await page.getByLabel('布局图表', { exact: true }).selectOption(before.cards[index].id);
    // 先改宽高，随后坐标，避免边界约束影响目标位置。
    for (const i of [2, 3, 0, 1]) {
      const field = page.getByLabel(`图表${['X', 'Y', '宽度', '高度'][i]}`, { exact: true });
      await field.fill(String(rect[i])); await field.press('Enter');
    }
  }
  await page.getByLabel('布局图表', { exact: true }).selectOption(id);
  const custom = await plan(page);
  assert.deepEqual(custom.cards[0].layout, { x: 96, y: 144, width: 540, height: 216 });
  assert.deepEqual(custom.cards.map(c => c.config), before.cards.map(c => c.config));
  await page.getByRole('button', { name: '自动排列', exact: true }).click();
  await page.getByRole('button', { name: '撤销布局', exact: true }).click();
  assert.deepEqual((await plan(page)).cards, custom.cards);
  // 切换窗口尺寸不改变逻辑坐标，图表相对位置始终只经过一个比例变换。
  for (const size of [[1200, 900], [1600, 900]]) {
    await app.evaluate(({ BrowserWindow }, [w, h]) => BrowserWindow.getAllWindows()[0].setSize(w, h), size);
    await page.waitForFunction(() => {
      const host = document.querySelector('.dashboard-editor-viewport'), stage = host.querySelector('.dashboard-artboard');
      const a = host.getBoundingClientRect(), b = stage.getBoundingClientRect();
      return Math.abs(b.width - (a.width - 2)) < 3;
    });
    const geometry = await card.evaluate(el => {
      const a = el.getBoundingClientRect(), b = el.closest('.dashboard-artboard').getBoundingClientRect();
      return { x: (a.x - b.x) / b.width * 1920, y: (a.y - b.y) / b.height * 1080, width: a.width / b.width * 1920, height: a.height / b.height * 1080 };
    });
    for (const key of ['x', 'y', 'width', 'height']) close(geometry[key], custom.cards[0].layout[key], 0.2);
    assert.deepEqual((await plan(page)).cards, custom.cards);
  }
  await page.getByRole('button', { name: '大屏预览', exact: true }).click();
  const modal = page.locator('.dashboard-presentation'); await modal.waitFor();
  const fit = await modal.locator('.dashboard-fit').boundingBox(), shown = await modal.locator('.dashboard-artboard').boundingBox();
  close(shown.width / shown.height, 16 / 9, 0.001);
  assert.ok(shown.width <= fit.width + 1 && shown.height <= fit.height + 1);
  await page.screenshot({ path: join(scratch, '03-layout-presentation.png') });
  await page.keyboard.press('Escape'); await modal.waitFor({ state: 'detached' });
  await page.getByRole('button', { name: '查看详情', exact: true }).click();
  assert.ok(await page.locator('.dashboard-dialog .analysis-method').textContent());
  await page.keyboard.press('Escape');
  await page.getByRole('button', { name: '预览布局', exact: true }).click();
  assert.equal(await board.locator('.dashboard-handle').count(), 0);
  await page.getByRole('button', { name: '编辑布局', exact: true }).click();
  await board.screenshot({ path: join(scratch, '05-layout-editor.png') });
  return custom;
}

async function verifyReport(app, htmlPath, expected, scratch) {
  // 同一 Electron 内创建测试窗口，实际从 file:// 打开离线 HTML。
  const pending = app.waitForEvent('window');
  await app.evaluate(({ BrowserWindow }, url) => {
    const win = new BrowserWindow({ width: 1400, height: 950, show: true, webPreferences: { nodeIntegration: false, contextIsolation: true, backgroundThrottling: false } });
    void win.loadURL(url);
  }, pathToFileURL(htmlPath).href);
  const report = await pending;
  await report.locator('.dashboard-tile').first().waitFor();
  assert.equal(await report.locator('.dashboard-tile').count(), expected.cards.length);
  const shapes = await report.locator('.dashboard-tile').evaluateAll(els => els.map(el => {
    const rect = el.getBoundingClientRect(), stage = el.closest('.dashboard-artboard').getBoundingClientRect();
    return { id: el.dataset.cardId, x: (rect.x - stage.x) / stage.width * 1920, y: (rect.y - stage.y) / stage.height * 1080,
      width: rect.width / stage.width * 1920, height: rect.height / stage.height * 1080, z: getComputedStyle(el).zIndex };
  }));
  shapes.forEach((shape, i) => {
    assert.equal(shape.id, expected.cards[i].id); assert.equal(Number(shape.z), i + 1);
    for (const key of ['x', 'y', 'width', 'height']) close(shape[key], expected.cards[i].layout[key], 0.1);
  });
  await report.waitForFunction(() => [...document.images].every(img => img.complete && img.naturalWidth > 0));
  await report.screenshot({ path: join(scratch, '04-offline-layout.png') });
  await report.close();
}
module.exports = { layoutInteractions, verifyReport, plan };
