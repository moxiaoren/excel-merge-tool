// app.js UI 层冒烟测试：用 Proxy 万能 stub 加载，验证重构后元数据驱动路径
// 运行：node --test test/
const assert = require('node:assert');
const { test } = require('node:test');
const vm = require('node:vm');

function makeEl() {
  const el = {
    _html: '', style: {}, dataset: {}, children: [],
    classList: { add() {}, remove() {}, toggle() {}, contains() { return false; } },
    _listeners: {},
    addEventListener(type, fn) { this._listeners[type] = fn; },
    appendChild(c) { this.children.push(c); return c; },
    querySelector() { return makeEl(); },
    querySelectorAll() { return []; },
    setAttribute() {}, getAttribute() { return ''; },
    closest() { return null; },
    focus() {},
    click() {},
    remove() {},
  };
  Object.defineProperty(el, 'innerHTML', { get() { return this._html; }, set(v) { this._html = v; } });
  Object.defineProperty(el, 'textContent', { get() { return this._html; }, set(v) { this._html = v; } });
  return el;
}

function makeDocument() {
  const reg = {};
  const doc = {
    readyState: 'loading',
    addEventListener(type, fn) { reg[type] = fn; },
    getElementById() { return makeEl(); },
    querySelectorAll() { return []; },
    querySelector() { return makeEl(); },
    createElement() { return makeEl(); },
    body: makeEl(),
  };
  return doc;
}

test('app.js 可加载且 getSelectedSources 使用元数据而非 workbook', () => {
  // 浏览器式全局
  const sandbox = {
    self: null,             // 占位，下面指向全局（core/engine 挂载目标）
    window: { scrollTo() {}, addEventListener() {} },
    document: makeDocument(),
    localStorage: { getItem() { return null; }, setItem() {} },
    alert() {}, confirm() { return true; }, prompt() { return ''; },
    setTimeout, clearTimeout, console,
  };
  sandbox.self = sandbox;   // root 指向全局对象
  // 依次加载 xlsx / core / engine / app
  vm.createContext(sandbox);
  vm.runInContext(require('fs').readFileSync('vendor/xlsx.full.min.js', 'utf8'), sandbox);
  vm.runInContext(require('fs').readFileSync('core.js', 'utf8'), sandbox);
  vm.runInContext(require('fs').readFileSync('engine.js', 'utf8'), sandbox);
  vm.runInContext(require('fs').readFileSync('app.js', 'utf8'), sandbox);

  // 触发 DOMContentLoaded（init）
  assert.strictEqual(sandbox.document.readyState, 'loading');
  // const state 为词法绑定，需在 context 内求值
  const run = (code) => vm.runInContext(code, sandbox);
  assert.strictEqual(run('typeof state'), 'object', 'state 对象存在');
  assert.strictEqual(run('state.uploadedFiles.length'), 0, 'uploadedFiles 初始为空');
  assert.strictEqual(run('state.staffList.length'), 0, 'staffList 初始为空名单（无硬编码姓名）');

  // 模拟已解析文件元数据，验证 getSelectedSources / collectAllSourceColumns 用 sheetHeaders
  run(`state.uploadedFiles.push({
    name: 'a.xlsx', wbId: 't1', status: 'loaded',
    sheetNames: ['S1'],
    selectedSheets: ['S1'],
    totalRows: 3,
    sheetHeaders: { S1: ['姓名', '值'] },
    sheetRows: { S1: 2 },
    largeFile: false
  });`);
  const sources = run('getSelectedSources()');
  assert.strictEqual(sources.length, 1);
  assert.deepStrictEqual(Array.from(sources[0].headers), ['姓名', '值']);
  assert.strictEqual(sources[0].wbId, 't1');

  run('collectAllSourceColumns()');
  assert.deepStrictEqual(Array.from(run('state.allSourceColumns')), ['姓名', '值']);

  // 移除文件应调用 cleanup（不抛错）
  run('removeFile(0)');
  assert.strictEqual(run('state.uploadedFiles.length'), 0);
});
