// Copyright (c) 2026 AbeLin · MIT License
const test = require('node:test');
const assert = require('node:assert/strict');
const { parseAnalysisPlan } = require('../dist-electron/shared/analysis');
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
