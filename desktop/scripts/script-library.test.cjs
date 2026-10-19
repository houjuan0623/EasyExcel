// Copyright (c) 2026 AbeLin · MIT License
const { test } = require('node:test');
const assert = require('node:assert/strict');
const { mkdtempSync, mkdirSync, writeFileSync, readFileSync, readdirSync } = require('node:fs');
const { tmpdir } = require('node:os');
const { join } = require('node:path');
const { ScriptLibrary } = require('../dist-electron/electron/script-library.js');
const root = process.env.EXCELROUTER_TEST_DIR || tmpdir();
mkdirSync(root, { recursive: true });
const scratch = mkdtempSync(join(root, 'script-library-'));
function library(name) { return new ScriptLibrary(join(scratch, name)); }

test('same code deduplicates across CRLF and restarts, edited code stays separate', () => {
  const store = library('dedup');
  const first = store.recordRun('print("第一次")\r\n', '月度核对.py');
  store.finishRun(first.id, 'success');
  const next = library('dedup').recordRun('print("第一次")\n', '另一个名字.py');
  assert.equal(next.id, first.id);
  assert.equal(next.runCount, 2);
  assert.equal(next.name, '月度核对');
  store.recordRun('print("第二次")\n', '月度核对.py');
  assert.equal(store.list().items.length, 2);
  assert.equal(store.get(first.id).code, 'print("第一次")\n');
});

test('metadata, favorites, trash and restore persist without losing code', () => {
  const store = library('metadata');
  const item = store.save({ code: 'print(42)', name: '示例' });
  store.update(item.id, { name: '精确核对', category: '财务', notes: '金额按分汇总', favorite: true, archived: true });
  const same = store.save({ code: item.code, name: '临时名称', category: '草稿', rememberOnly: true });
  assert.equal(same.name, '精确核对');
  assert.equal(same.category, '财务');
  assert.equal(same.archived, true);
  const restored = library('metadata').update(item.id, { archived: false });
  assert.equal(restored.favorite, true);
  assert.equal(restored.code, item.code);
});

test('backup roundtrip and repeated imports do not multiply run counts', () => {
  const source = library('backup-source');
  const item = source.recordRun('print("备份")', '备份脚本');
  source.finishRun(item.id, 'error');
  source.update(item.id, { favorite: true, notes: '一条说明' });
  const content = source.exportBackup();
  const dest = library('backup-target');
  assert.deepEqual(dest.importBackup(content), { added: 1, merged: 0 });
  assert.deepEqual(dest.importBackup(content), { added: 0, merged: 1 });
  assert.equal(dest.get(item.id).runCount, 1);
  assert.equal(dest.get(item.id).notes, '一条说明');
  assert.equal(dest.get(item.id).status, 'error');
  const bad = JSON.parse(content);
  bad.scripts.push({ ...bad.scripts[0], id: '../../outside' });
  assert.throws(() => dest.importBackup(JSON.stringify(bad)), /编号/);
  assert.equal(dest.list().items.length, 1);
});

test('damaged records are retained and can be repaired by importing a backup', () => {
  const store = library('damage');
  const item = store.save({ code: 'print("保留")', name: '保留' });
  const backup = store.exportBackup();
  writeFileSync(join(store.path, item.id + '.json'), '{broken', 'utf8');
  assert.equal(store.list().warnings.length, 1);
  assert.equal(store.list().items.length, 0);
  store.importBackup(backup);
  assert.equal(store.get(item.id).code, item.code);
  assert.equal(store.list().warnings.length, 0);
  const damaged = readdirSync(store.path).find((name) => name.endsWith('.damaged'));
  assert.equal(readFileSync(join(store.path, damaged), 'utf8'), '{broken');
});

test('old run directories migrate once; new runs are not imported twice', () => {
  const store = library('legacy');
  const output = join(scratch, 'legacy-output'); mkdirSync(output);
  writeFileSync(join(output, 'script.py'), 'print("旧版本")', 'utf8');
  const entry = { id: 'old-1', mode: 'code', output, name: '月表.xlsx', summary: { status: 'success' }, stopped: false, date: '2026-10-19T10:00:00Z' };
  store.migrateHistory([entry]);
  const item = store.list().items[0];
  library('legacy').migrateHistory([entry]);
  store.migrateHistory([{ ...entry, id: 'new-1', summary: { scriptId: item.id } }]);
  assert.equal(store.get(item.id).runCount, 1);
  assert.equal(store.get(item.id).lastRunAt, new Date(entry.date).toISOString());
});

test('invalid identifiers and empty scripts cannot write outside the library', () => {
  const store = library('validation');
  assert.throws(() => store.get('../../test'), /编号/);
  assert.throws(() => store.save({ code: ' ', name: '空' }), /不能为空/);
  assert.throws(() => store.save({ code: 'x'.repeat(200001), name: '大' }), /最多/);
  const item = store.save({ code: 'print(1)', name: '有效' });
  assert.throws(() => store.update(item.id, { favorite: 'false' }), /格式/);
  assert.equal(store.list().items.length, 1);
});
