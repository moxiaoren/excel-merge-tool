/* ============================================================
   Excel 表格合并工具 · worker.js  v3.0.0
   Web Worker：承担 解析 与 按行合并 两种重量级计算，
   主线程仅保留轻量元数据并通过 wbId 引用 Workbook，避免大对象驻留主线程造成卡死。
   ============================================================ */
'use strict';

// 主线程加载完 xlsx 后，worker 通过 importScripts 加载同一份库与纯逻辑
importScripts('vendor/xlsx.full.min.js');
importScripts('core.js');

const C = self.MergeCore;

// 解析好的 Workbook 存放于此（worker 私有堆），主线程只持有 wbId
const workbooks = {};

function post(type, payload) {
  self.postMessage(Object.assign({ type: type }, payload));
}

self.onmessage = function (e) {
  const msg = e.data || {};
  try {
    switch (msg.type) {
      case 'parse': {
        // msg: { wbId, data, isCSV }
        const result = C.parseWorkbook(XLSX, msg.data, { csvText: msg.isCSV ? msg.data : undefined });
        workbooks[msg.wbId] = result.workbook;
        // 只回传轻量元数据，不让大 workbook 跨线程拷贝
        post('parsed', {
          reqId: msg.reqId,
          wbId: msg.wbId,
          sheetNames: result.sheetNames,
          totalRows: result.totalRows,
          sheetHeaders: result.sheetHeaders,
          sheetRows: result.sheetRows
        });
        break;
      }
      case 'merge': {
        // msg: { sources:[{wbId, sheet, headers}], mapping, staffEnable, staffList, reqId }
        const sources = (msg.sources || [])
          .filter(s => workbooks[s.wbId] && workbooks[s.wbId].Sheets[s.sheet])
          .map(s => ({ workbook: workbooks[s.wbId], sheet: s.sheet, headers: s.headers || [] }));
        if (!sources.length) { post('mergeError', { reqId: msg.reqId, message: '没有可合并的数据源（可能已清理）' }); break; }
        const onProgress = function (pct) { post('progress', { reqId: msg.reqId, pct: pct, wbId: msg.wbId }); };
        const data = C.buildMergedRows(XLSX, sources, msg.mapping, { staffEnable: msg.staffEnable, staffList: msg.staffList || [] }, onProgress);
        // 注意：这里不再调用 processRemarkColumn，由 engine 层统一处理，保证多源合并语义一致
        post('merged', { reqId: msg.reqId, wbId: msg.wbId, data: data });
        break;
      }
      case 'cleanup': {
        delete workbooks[msg.wbId];
        break;
      }
      default:
        post('mergeError', { reqId: msg.reqId, message: '未知消息类型' });
    }
  } catch (err) {
    post('mergeError', { reqId: msg.reqId, message: (err && err.message) ? err.message : String(err) });
  }
};
