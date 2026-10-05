// engine.js 集成测试（Node 环境无 Worker，走降级同步路径）
// 运行：node --test test/
const assert = require('node:assert');
const { test } = require('node:test');

// 按 index.html 的顺序加载：xlsx -> core -> engine
globalThis.XLSX = require('../vendor/xlsx.full.min.js');
globalThis.self = globalThis;
globalThis.MergeCore = (function () {
  require('../core.js');
  return globalThis.MergeCore;
})();
require('../engine.js');

const Engine = globalThis.Engine;

test('engine 在 Node 无 Worker 时降级同步（supported=false）', () => {
  assert.strictEqual(Engine.supported, false);
});

test('engine.parse 返回轻量元数据', async () => {
  const aoa = [
    ['日期', '姓名', '内容'],
    ['2023-01-02', '张三', '你好'],
    ['2023-01-03', '李四', '世界']
  ];
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
  const buffer = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
  const meta = await Engine.parse(buffer, false, 'wb-1');
  assert.deepStrictEqual(meta.sheetNames, ['Sheet1']);
  assert.strictEqual(meta.totalRows, 3);
  assert.deepStrictEqual(meta.sheetHeaders['Sheet1'], ['日期', '姓名', '内容']);
  assert.strictEqual(meta.sheetRows['Sheet1'], 2);
});

test('engine.merge 端到端：多列合并+人员过滤+日期格式化', async () => {
  const aoa = [
    ['日期', '审核时间', '姓名', '内容'],
    ['2023-01-02', '2023-01-03', '张三', '你好'],
    ['abc', '2023-01-05', '李四', '世界']
  ];
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
  const buffer = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
  const meta = await Engine.parse(buffer, false, 'wb-2');

  const mapping = [
    { targetCol: '质检时间', sourceCol: ['日期', '审核时间'] },
    { targetCol: '人员', sourceCol: ['姓名'] },
    { targetCol: '内容', sourceCol: ['内容'] }
  ];
  const sources = [{ wbId: 'wb-2', sheet: 'Sheet1', headers: meta.sheetHeaders['Sheet1'] }];

  // 过滤开，名单含张三 -> 保留1行
  const out1 = await Engine.merge(sources, mapping, { staffEnable: true, staffList: ['张三'] });
  assert.strictEqual(out1.length, 1);
  assert.strictEqual(out1[0]['质检时间'], '01-02');
  assert.strictEqual(out1[0]['人员'], '张三');

  // 进度回调
  const pcts = [];
  const out2 = await Engine.merge(sources, mapping, { staffEnable: false, staffList: [] }, p => pcts.push(p));
  assert.strictEqual(out2.length, 2);
  assert.strictEqual(pcts[pcts.length - 1], 100);
});

test('engine.merge 支持多文件（多 wbId）跨源合并', async () => {
  const mk = (rows) => {
    const ws = XLSX.utils.aoa_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    return XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
  };
  const m1 = await Engine.parse(mk([['姓名', 'A'], ['张三', 'x']]), false, 'wb-a');
  const m2 = await Engine.parse(mk([['姓名', 'B'], ['李四', 'y']]), false, 'wb-b');
  const mapping = [{ targetCol: '姓名', sourceCol: ['姓名'] }, { targetCol: '数据', sourceCol: ['A'] }];
  const sources = [
    { wbId: 'wb-a', sheet: 'Sheet1', headers: m1.sheetHeaders['Sheet1'] },
    { wbId: 'wb-b', sheet: 'Sheet1', headers: m2.sheetHeaders['Sheet1'] }
  ];
  // 注意：'数据' 列 sourceCol [A]，wb-b 无此列 -> 留空
  const out = await Engine.merge(sources, mapping, { staffEnable: false }, () => {});
  assert.strictEqual(out.length, 2);
  assert.strictEqual(out[0]['姓名'], '张三');
  assert.strictEqual(out[0]['数据'], 'x');
  assert.strictEqual(out[1]['姓名'], '李四');
  assert.strictEqual(out[1]['数据'], '');
  // 来源sheet名列
  assert.strictEqual(out[1]['来源sheet名'], 'Sheet1');
});

test('engine.cleanup 释放 wbId（降级不报错）', async () => {
  await Engine.parse('x', true, 'wb-clean');  // csv 解析
  Engine.cleanup('wb-clean');                 // 不应抛错
  assert.ok(true);
});
