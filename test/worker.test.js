// worker.js 消息处理逻辑冒烟测试（Node 模拟 self/postMessage）
// 运行：node --test test/
const assert = require('node:assert');
const { test } = require('node:test');

test('worker 消息处理：parse + merge + cleanup 全流程', () => {
  const global = globalThis;
  global.XLSX = require('../vendor/xlsx.full.min.js');
  // 模拟 worker/浏览器环境：core.js 挂到 self
  global.self = global;
  // 手动加载 core.js（模拟 worker 中 importScripts 的产物）挂到全局
  require('../core.js');
  // 初始化 importScripts 桩
  global.importScripts = function () {};

  // 构造 worker 的 self 作用域
  const sent = [];
  const selfObj = {
    MergeCore: globalThis.MergeCore,
    postMessage: (msg) => sent.push(msg)
  };

  // 执行 worker.js
  const wfunc = new Function('self', 'XLSX', require('fs').readFileSync('worker.js', 'utf8'));
  wfunc(selfObj, global.XLSX);
  // 触发 onmessage
  const fire = (data) => selfObj.onmessage({ data });

  // --- parse ---
  const aoa = [['姓名', '值'], ['张三', 'a'], ['李四', 'b']];
  const ws = global.XLSX.utils.aoa_to_sheet(aoa);
  const wb = global.XLSX.utils.book_new();
  global.XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
  const buffer = global.XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
  fire({ type: 'parse', wbId: 'w1', data: buffer, isCSV: false, reqId: 1 });
  const parsed = sent.find(m => m.type === 'parsed');
  assert.ok(parsed, '应有 parsed 响应');
  assert.strictEqual(parsed.reqId, 1);
  assert.deepStrictEqual(parsed.sheetNames, ['Sheet1']);
  assert.deepStrictEqual(parsed.sheetHeaders['Sheet1'], ['姓名', '值']);
  assert.strictEqual(parsed.sheetRows['Sheet1'], 2);

  // --- merge ---
  const mapping = [{ targetCol: '姓名', sourceCol: ['姓名'] }];
  fire({ type: 'merge', reqId: 2, sources: [{ wbId: 'w1', sheet: 'Sheet1', headers: ['姓名', '值'] }], mapping, staffEnable: false, staffList: [] });
  const merged = sent.find(m => m.type === 'merged');
  assert.ok(merged, '应有 merged 响应');
  assert.strictEqual(merged.data.length, 2);
  assert.strictEqual(merged.data[0]['姓名'], '张三');
  // 进度消息
  assert.ok(sent.some(m => m.type === 'progress' && m.reqId === 2), '应有 progress 消息');

  // --- cleanup 后 merge 报错 ---
  fire({ type: 'cleanup', wbId: 'w1' });
  fire({ type: 'merge', reqId: 3, sources: [{ wbId: 'w1', sheet: 'Sheet1', headers: [] }], mapping, staffEnable: false, staffList: [] });
  const err = sent.find(m => m.type === 'mergeError' && m.reqId === 3);
  assert.ok(err, '清理后 merge 应报错');

  delete global.importScripts;
});
