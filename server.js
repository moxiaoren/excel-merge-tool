/* ============================================================
   Excel 表格合并工具 · server.js  v3.1.0（可选本地/服务器后端）
   ------------------------------------------------------------
   用途：当公司服务器可运行 Node 时，启动本进程处理大文件解析/合并，
         浏览器前端（Engine）探测到本服务后自动把大文件交由其处理，
         彻底规避浏览器单页内存限制。

   纯 Node 原生实现，零第三方依赖，复用前端 core.js 核心逻辑。

   接口：
     GET  /api/health   -> {"ok":true}                （前端探测）
     POST /api/parse    -> 请求体为文件二进制；query 带 name
                           -> {"sheetNames","totalRows","sheetHeaders","sheetRows","wbId"}
                           （大 workbook 留存服务器内存，仅回传轻量元数据 + wbId）
     POST /api/merge    -> JSON {wbId, sheet, headers, mapping, staffEnable, staffList}
                           -> {"data":[...]}          （与本地格式完全一致）
     POST /api/cleanup  -> JSON/query {wbId} -> {"ok":true }（释放服务器内存）

   启动：node server.js [端口]   （默认 8787）
   内存：Node 进程默认堆即可，如需更大可 --max-old-space-size=8192
   ============================================================ */
'use strict';

// core.js 依赖 self/globalThis 挂载，先铺好环境再加载
globalThis.self = globalThis;
globalThis.XLSX = require('./vendor/xlsx.full.min.js');
require('./core.js');
const MergeCore = globalThis.MergeCore;

const http = require('http');
const PORT = parseInt(process.argv[2] || '8787', 10);

// wbId -> workbook 大对象（服务器内存中），解析后留存，合并完/移除时释放
const wbStore = {};
let seq = 0;

const json = (res, code, obj) => {
  res.writeHead(code, { 'Content-Type': 'application/json; charset=utf-8', 'Access-Control-Allow-Origin': '*', 'Access-Control-Allow-Methods': 'GET,POST,OPTIONS', 'Access-Control-Allow-Headers': 'Content-Type' });
  res.end(JSON.stringify(obj));
};

const readBody = (req) => new Promise((resolve, reject) => {
  const chunks = [];
  let size = 0;
  req.on('data', c => { chunks.push(c); size += c.length; });
  req.on('end', () => { try { resolve(Buffer.concat(chunks)); } catch (e) { reject(e); } });
  req.on('error', reject);
});

function handleParse(req, res, body) {
  let name = '';
  const q = new URL(req.url, 'http://x');
  name = q.searchParams.get('name') || 'upload.xlsx';
  const isCSV = /\.csv$/i.test(name);
  let result;
  try {
    result = isCSV
      ? MergeCore.parseWorkbook(XLSX, body.toString('utf8'), { csvText: body.toString('utf8') })
      : MergeCore.parseWorkbook(XLSX, new Uint8Array(body), {});
  } catch (e) {
    return json(res, 500, { error: '解析失败: ' + e.message });
  }
  const wbId = ++seq;
  wbStore[wbId] = result.workbook;
  json(res, 200, { wbId, sheetNames: result.sheetNames, totalRows: result.totalRows, sheetHeaders: result.sheetHeaders, sheetRows: result.sheetRows });
}

function handleMerge(req, res, body) {
  let payload;
  try { payload = JSON.parse(body.toString('utf8')); }
  catch (e) { return json(res, 400, { error: '请求体不是合法 JSON' }); }
  const { sources, wbId, sheet, headers, mapping, staffEnable, staffList } = payload || {};
  // 兼容两种入参：新格式 sources:[{wbId,sheet,headers}]；旧格式单 wbId+sheet+headers
  const items = Array.isArray(sources) && sources.length
    ? sources
    : (wbId != null ? [{ wbId, sheet, headers: headers || [] }] : []);
  if (!items.length) return json(res, 400, { error: '缺少 sources 或 wbId' });
  const real = items
    .filter(s => wbStore[s.wbId] && wbStore[s.wbId].Sheets && wbStore[s.wbId].Sheets[s.sheet])
    .map(s => ({ workbook: wbStore[s.wbId], sheet: s.sheet, headers: s.headers || [] }));
  if (!real.length) return json(res, 404, { error: '未找到文件数据（可能已清理）' });
  try {
    // 返回原始按行合并结果（不做 processRemarkColumn），由前端统一处理，保证多源/跨端语义一致
    const data = MergeCore.buildMergedRows(XLSX, real, mapping || [], { staffEnable: !!staffEnable, staffList: staffList || [] });
    json(res, 200, { data: data });
  } catch (e) {
    json(res, 500, { error: '合并失败: ' + e.message });
  }
}

function handleCleanup(req, res, body) {
  let payload;
  try { payload = JSON.parse(body.toString('utf8')); } catch (e) { payload = {}; }
  const wbId = payload && payload.wbId;
  if (wbId != null && wbStore[wbId]) delete wbStore[wbId];
  json(res, 200, { ok: true });
}

// 简单路由：按 URL 前缀分发
function route(req, res) {
  const url = req.url || '/';
  const path = url.split('?')[0];
  if (req.method === 'OPTIONS') { json(res, 204, {}); return; }
  if (path === '/api/health') { json(res, 200, { ok: true }); return; }

  const collect = (fn) => readBody(req).then(b => fn(req, res, b)).catch(e => json(res, 500, { error: '读取请求体失败: ' + e.message }));

  if (req.method === 'POST' && path === '/api/parse') return collect(handleParse);
  if (req.method === 'POST' && path === '/api/merge') return collect(handleMerge);
  if (req.method === 'POST' && path === '/api/cleanup') return collect(handleCleanup);

  json(res, 404, { error: '未知接口: ' + path });
}

const server = http.createServer(route);
server.listen(PORT, () => {
  console.log('Excel 合并工具后端已启动: http://localhost:' + PORT);
  console.log('  /api/health  /api/parse  /api/merge  /api/cleanup');
  console.log('(可在 index.html 的 BACKEND_URL 中填入本地址以启用后端处理)');
});
