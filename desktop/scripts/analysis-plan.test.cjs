// Copyright (c) 2026 AbeLin · MIT License
const test = require('node:test');
const assert = require('node:assert/strict');
const { parseAnalysisPlan } = require('../dist-electron/shared/analysis');
const { arrangeLayouts, validLayout, transformLayout, nextLayout } = require('../dist-electron/shared/dashboard-layout');
const fixture = () => ({ version: 1, title: '看板', cards: [{ id: 'card-1', config: {
  inputPath: 'C:/data/book.xlsx', blocks: [{ id: 'block-1', name: '区域', path: 'C:/data/book.xlsx', sheet: 'Data', range: 'A1:B20', headerRows: 1, orientation: 'columns', includeHidden: false, skipTotals: true }],
  token: 'checked', kind: 'bar', title: '对比', unit: '元', note: '', combine: 'single', joinKeys: {}, joinType: 'inner', x: 'block-1:f0', y: ['block-1:f1'], aggregation: 'sum', sort: 'source', timeBucket: 'none', bins: 12, limit: 20, fieldLabels: { 'block-1:f1': '金额' },
} }] });
test('configuration roundtrip discards executable and image payloads', () => {
  const plan = fixture(); plan.cards[0].config.code = 'do not execute'; plan.cards[0].result = { image: 'data:image/png;base64,unused' };
  const parsed = parseAnalysisPlan(plan);
  assert.equal(parsed.cards[0].config.code, undefined);
  assert.equal(parsed.cards[0].result, undefined);
  assert.equal(parsed.cards[0].config.y[0], 'block-1:f1');
});
test('malformed, duplicated and prototype identifiers cannot enter dashboard state', () => {
  for (const id of ['__proto__', 'constructor', '', 'a/b']) {
    const plan = fixture(); plan.cards[0].id = id; assert.throws(() => parseAnalysisPlan(plan));
  }
  const plan = fixture(); plan.cards.push(plan.cards[0]); assert.throws(() => parseAnalysisPlan(plan));
  assert.throws(() => parseAnalysisPlan({ version: 5, cards: [] }));
});
test('limits and unsupported operations fail before import', () => {
  for (const [key, value] of [['bins', 100000], ['kind', 'execute-python'], ['combine', 'cross'], ['y', {}], ['limit', -1]]) {
    const plan = fixture(); plan.cards[0].config[key] = value; assert.throws(() => parseAnalysisPlan(plan));
  }
});

test('v1 migrates to a bounded 16:9 layout and v2 preserves geometry and order', () => {
  const plan = fixture();
  for (let i = 1; i < 12; i++) plan.cards.push({ ...plan.cards[0], id: `card-${i + 1}` });
  const migrated = parseAnalysisPlan(plan);
  assert.equal(migrated.version, 2);
  assert.deepEqual(migrated.canvas, { width: 1920, height: 1080 });
  assert.ok(migrated.cards.every(c => validLayout(c.layout)));
  migrated.cards[0].layout = { x: 100, y: 140, width: 600, height: 420 };
  assert.deepEqual(parseAnalysisPlan(migrated), migrated);
  for (let count = 1; count <= 12; count++) {
    const layouts = arrangeLayouts(count);
    assert.ok(layouts.every(validLayout));
    layouts.forEach((a, i) => layouts.slice(i + 1).forEach(b => assert.ok(a.x + a.width <= b.x || b.x + b.width <= a.x || a.y + a.height <= b.y || b.y + b.height <= a.y)));
  }
});

test('v2 refuses out-of-bounds, missing, nonfinite and executable layout values', () => {
  for (const [key, value] of [['x', -1], ['y', 119], ['width', 1920], ['height', 1], ['x', NaN], ['y', Infinity], ['x', '24px;display:none'], ['x', 24.5]]) {
    const plan = parseAnalysisPlan(fixture()); plan.cards[0].layout[key] = value;
    assert.throws(() => parseAnalysisPlan(plan));
  }
  const plan = parseAnalysisPlan(fixture()); delete plan.cards[0].layout;
  assert.throws(() => parseAnalysisPlan(plan));
  plan.canvas.width = 3840; assert.throws(() => parseAnalysisPlan(plan));
});

test('pointer deltas clamp, snap and resize against the anchored opposite edge', () => {
  const r = { x: 120, y: 240, width: 600, height: 420 };
  assert.deepEqual(transformLayout(r, 31, -17, 'move'), { ...r, x: 156, y: 228 });
  assert.deepEqual(transformLayout(r, 31, -17, 'move', false), { ...r, x: 151, y: 223 });
  assert.deepEqual(transformLayout(r, 10000, 10000, 'move'), { ...r, x: 1296, y: 636 });
  const nw = transformLayout(r, 10000, 10000, 'nw');
  assert.deepEqual(nw, { x: 480, y: 492, width: 240, height: 168 });
  assert.equal(nw.x + nw.width, r.x + r.width); assert.equal(nw.y + nw.height, r.y + r.height);
  for (const handle of ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w']) {
    for (const delta of [-9999, -113, 0, 13, 9999]) assert.ok(validLayout(transformLayout(r, delta, delta, handle)));
  }
  const next = nextLayout([r]); assert.ok(validLayout(next));
  assert.deepEqual(r, { x: 120, y: 240, width: 600, height: 420 });
});
