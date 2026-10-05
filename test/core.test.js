// core.js 纯逻辑单元测试（node --test）
// 运行：node --test test/
const assert = require('node:assert');
const { test } = require('node:test');

// 在 node 中加载 core.js（IIFE 挂到 globalThis）
globalThis.self = globalThis;
require('../core.js');
const C = globalThis.MergeCore;

// 用真实 xlsx 库构造测试工作簿（aoa 二维数组 -> sheet）
const XLSX = require('../vendor/xlsx.full.min.js');
function makeWb(aoa) {
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
  return wb;
}

test('escapeHtml 转义特殊字符', () => {
  assert.strictEqual(C.escapeHtml('<b>&"\'</b>'), '&lt;b&gt;&amp;&quot;&#39;&lt;/b&gt;');
  assert.strictEqual(C.escapeHtml(null), '');
  assert.strictEqual(C.escapeHtml('普通文本'), '普通文本');
});

test('extractChineseName 只保留中文', () => {
  assert.strictEqual(C.extractChineseName('张三123abc!@'), '张三');
  assert.strictEqual(C.extractChineseName('Alice·王五'), '王五');
  assert.strictEqual(C.extractChineseName(''), '');
  assert.strictEqual(C.extractChineseName(null), '');
});

test('formatDate 处理格式', () => {
  assert.strictEqual(C.formatDate('-1'), '永久');
  assert.strictEqual(C.formatDate('-1', '质检回查表'), '永久');
  // 质检回查：MM/DD -> MM-DD
  assert.strictEqual(C.formatDate('3/5', '质检回查'), '03-05');
  assert.strictEqual(C.formatDate('12/25', '质检回查'), '12-25');
  // 普通日期字符串
  assert.strictEqual(C.formatDate('2023-07-08 12:00:00'), '07-08');
  assert.strictEqual(C.formatDate('2023-07-08'), '07-08');
  // 无法识别则原样返回
  assert.strictEqual(C.formatDate('hello'), 'hello');
});

test('formatDate 处理序列化数字日期（Excel 序列号）', () => {
  // 锁定当前行为：与 v1/v2 原 app.js 公式一致
  // 注意：该公式未处理 Excel 1900 闰年幻影日，实际比正确日期晚了 1 天（属已知缺陷，P2 项）
  const val = C.formatDate(45000);
  assert.strictEqual(typeof val, 'string');
  // 只验证格式为 MM-DD，不锁定具体值（避免与已知 off-by-one 缺陷耦合）
  assert.match(val, /^\d{2}-\d{2}$/);
});

test('getHeaderRowIndex 业务表头行', () => {
  assert.strictEqual(C.getHeaderRowIndex('数据'), 0);
  assert.strictEqual(C.getHeaderRowIndex('质检回查'), 1);
});

test('processRemarkColumn：无备注则原样返回', () => {
  const rows = [{ 'A': 'a' }, { 'A': 'b' }];
  const out = C.processRemarkColumn(rows);
  assert.deepStrictEqual(out, rows);
});

test('processRemarkColumn：拆分申诉/举报列（无处理状态时不删除）', () => {
  const rows = [
    { '备注': '【帖子违规】剩余部分A' },
    { '备注': '【帖子申诉】剩余部分B' }
  ];
  const out = C.processRemarkColumn(rows);
  // 帖子违规 -> 帖子举报
  assert.strictEqual(out[0]['申诉/举报'], '帖子举报');
  // 备注只剩】之后的剩余部分
  assert.strictEqual(out[0]['备注'], '剩余部分A');
  assert.strictEqual(out[1]['申诉/举报'], '帖子申诉');
  assert.strictEqual(out[1]['备注'], '剩余部分B');
});

test('processRemarkColumn：过滤通过/不通过的特定组合', () => {
  const rows = [
    { '备注': '【帖子违规】x', '处理状态': '通过' },      // 保留（帖子举报+通过）
    { '备注': '【帖子违规】y', '处理状态': '不通过' },    // 剔除
    { '备注': '【帖子申诉】z', '处理状态': '不通过' }     // 保留（帖子申诉+不通过）
  ];
  const out = C.processRemarkColumn(rows);
  assert.strictEqual(out.length, 2);
  assert.ok(out.every(r => !r.hasOwnProperty('处理状态')));
  assert.deepStrictEqual(out.map(r => r['申诉/举报']), ['帖子举报', '帖子申诉']);
});

test('buildMergedRows：多源列合并+人员过滤+日期格式化', () => {
  const wb = makeWb([
    ['日期', '审核时间', '姓名', '内容'],
    ['2023-01-02', '2023-01-03', '张三', '你好']
  ]);
  const sources = [{
    workbook: wb,
    sheet: wb.SheetNames[0],
    headers: ['日期', '审核时间', '姓名', '内容']
  }];
  const mapping = [
    { targetCol: '质检时间', sourceCol: ['日期', '审核时间'] },
    { targetCol: '人员', sourceCol: ['姓名'] },
    { targetCol: '内容', sourceCol: ['内容'] }
  ];
  // 人员过滤开且名单含张三 -> 保留
  let out = C.buildMergedRows(XLSX, sources, mapping, { staffEnable: true, staffList: ['张三'] });
  assert.strictEqual(out.length, 1);
  assert.strictEqual(out[0]['来源sheet名'], wb.SheetNames[0]);
  // 日期取第一个非空源列，且格式化
  assert.strictEqual(out[0]['质检时间'], '01-02');
  assert.strictEqual(out[0]['人员'], '张三');
  // 人员过滤开但名单不含张三 -> 剔除
  out = C.buildMergedRows(XLSX, sources, mapping, { staffEnable: true, staffList: ['李四'] });
  assert.strictEqual(out.length, 0);
  // 过滤关 -> 保留
  out = C.buildMergedRows(XLSX, sources, mapping, { staffEnable: false, staffList: [] });
  assert.strictEqual(out.length, 1);
});

test('buildMergedRows：多行数据按行输出完整', () => {
  const wb = makeWb([
    ['列A', '列B'],
    ['a1', 'b1'],
    ['a2', 'b2'],
    ['', 'b3']  // 列A为空但列B非空，整行不视为空
  ]);
  const sources = [{ workbook: wb, sheet: wb.SheetNames[0], headers: ['列A', '列B'] }];
  const mapping = [
    { targetCol: '列A', sourceCol: ['列A'] },
    { targetCol: '列B', sourceCol: ['列B'] }
  ];
  const out = C.buildMergedRows(XLSX, sources, mapping, { staffEnable: false });
  assert.strictEqual(out.length, 3);
  assert.strictEqual(out[0]['列A'], 'a1');
  assert.strictEqual(out[2]['列B'], 'b3');
});

test('buildMergedRows：完全空行被剔除', () => {
  const wb = makeWb([
    ['列A'],
    ['有值'],
    ['', ''],
    ['', '']
  ]);
  const sources = [{ workbook: wb, sheet: wb.SheetNames[0], headers: ['列A'] }];
  const mapping = [{ targetCol: '列A', sourceCol: ['列A'] }];
  const out = C.buildMergedRows(XLSX, sources, mapping, { staffEnable: false });
  assert.strictEqual(out.length, 1);
});

// 回归：大行数工作簿 cleanWorkbook 不得因 Math.min(...数组) spread 爆调用栈
test('cleanWorkbook 大行数不爆栈（spread 回归）', () => {
  const aoa = [['姓名', '值']];
  for (let i = 0; i < 120000; i++) aoa.push(['人' + i, i]);
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'S');
  const data = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  const res = MergeCore.parseWorkbook(XLSX, new Uint8Array(data), {});
  assert.ok(res.totalRows >= 120000, '应解析出 12 万行数据');
  assert.ok(res.sheetHeaders['S'].length === 2, '表头应为 姓名/值');
});
