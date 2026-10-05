/* ============================================================
   Excel 表格合并工具 · engine.js  v3.1.0
   解析/合并执行引擎，三条路径按需切换：
     1) 后端服务（可选）：配置了 BACKEND_URL 且探测成功时，大文件 xlsx 交
        服务器解析/合并（/api/parse /api/merge /api/cleanup），彻底规避浏览器
        内存。小文件（≤ REMOTE_MIN_MB）仍本地处理，保护数据隐私。
     2) Web Worker（不冻结 UI）；不支持 Worker 的 file:// 降级为同步。
   app.js 只依赖本引擎，不关心底层在哪执行。后端 wbId 由内部映射隐藏，
   app.js 始终使用自己的 tempId 作为 wbId。
   ============================================================ */
'use strict';

const Engine = (function () {
  let worker = null;
  let seq = 0;
  const pending = {};       // reqId -> {resolve, reject, onProgress}
  const wbStore = {};       // wbId -> workbook（本地降级模式）

  // ---------- 后端配置 ----------
  const backendUrl = (typeof BACKEND_URL !== 'undefined' ? BACKEND_URL : '').replace(/\/+$/, '');
  const REMOTE_MIN_MB = 50; // xlsx 超过此大小且后端可用时，才走后端（隐私折中：小文件留本地）
  let backendAvailable = false;
  let backendChecked = false;
  const remoteMap = {};     // 本地tempId -> 服务器 wbId
  const remoteSet = new Set(); // 已走后端的本地 tempId 集合

  function checkBackend() {
    if (!backendUrl) { backendAvailable = false; backendChecked = true; return Promise.resolve(false); }
    return fetch(backendUrl + '/api/health', { method: 'GET' })
      .then(r => { backendAvailable = r.ok; backendChecked = true; return !!r.ok; })
      .catch(() => { backendAvailable = false; backendChecked = true; return false; });
  }

  // ---------- 尝试创建 Worker（失败则降级同步） ----------
  try {
    if (typeof Worker !== 'undefined') {
      worker = new Worker('worker.js');
      worker.onmessage = function (e) {
        const m = e.data || {};
        const p = pending[m.reqId];
        if (!p) return;
        if (m.type === 'progress' && p.onProgress) { p.onProgress(m.pct); return; }
        delete pending[m.reqId];
        if (m.type === 'parsed' || m.type === 'merged') p.resolve(m);
        else p.reject(new Error((m && m.message) || 'Worker 执行失败'));
      };
      worker.onerror = function (e) {
        Object.keys(pending).forEach(k => { pending[k].reject(new Error((e && e.message) || 'Worker 运行错误')); delete pending[k]; });
        try { worker.terminate(); } catch (x) {}
        worker = null;
      };
    }
  } catch (e) {
    worker = null;
  }

  function post(type, payload) {
    return new Promise(function (resolve, reject) {
      const id = ++seq;
      pending[id] = { resolve: resolve, reject: reject };
      worker.postMessage(Object.assign({ type: type, reqId: id }, payload));
    });
  }

  function backendParse(data, isCSV, wbId) {
    // 只走后端发 xlsx（csv 走本地）；文件名仅用于后端判类型，这里固定 .xlsx
    return fetch(backendUrl + '/api/parse?name=f.xlsx', { method: 'POST', body: data })
      .then(r => { if (!r.ok) throw new Error('服务器解析失败(' + r.status + ')'); return r.json(); })
      .then(j => {
        if (j.error) throw new Error(j.error);
        remoteMap[wbId] = j.wbId;
        remoteSet.add(wbId);
        return { sheetNames: j.sheetNames, totalRows: j.totalRows, sheetHeaders: j.sheetHeaders, sheetRows: j.sheetRows };
      });
  }

  function backendMerge(sources, mapping, opts) {
    const items = sources.map(s => ({ wbId: remoteMap[s.wbId], sheet: s.sheet, headers: s.headers || [] }));
    return fetch(backendUrl + '/api/merge', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ sources: items, mapping: mapping, staffEnable: !!opts.staffEnable, staffList: opts.staffList || [] })
    })
      .then(r => { if (!r.ok) return r.json().then(j => { throw new Error((j && j.error) || '服务器合并失败(' + r.status + ')'); }); return r.json(); })
      .then(j => { if (j.error) throw new Error(j.error); return j.data || []; });
  }

  return {
    get supported() { return !!worker; },
    // 是否已启用后端（UI 可据此提示）
    get backendOn() { return backendAvailable && remoteSet.size > 0; },
    initBackend: checkBackend,

    // 解析：data ArrayBuffer(xlsx)/字符串(csv)，wbId 用 app 自己的 tempId；size 为文件字节数
    parse(data, isCSV, wbId, size) {
      // CSV 永远本地（体积小，无内存压力，且格式编码走本地更稳）
      if (!isCSV && backendAvailable && size != null && size > REMOTE_MIN_MB * 1024 * 1024) {
        return backendParse(data, isCSV, wbId);
      }
      if (worker) return post('parse', { data: data, isCSV: isCSV, wbId: wbId });
      const result = MergeCore.parseWorkbook(XLSX, data, { csvText: isCSV ? data : undefined });
      wbStore[wbId] = result.workbook;
      return Promise.resolve({ sheetNames: result.sheetNames, totalRows: result.totalRows, sheetHeaders: result.sheetHeaders, sheetRows: result.sheetRows });
    },

    // 合并：sources=[{wbId,sheet,headers}]，wbId 为 app 的 tempId；可分远程/本地混合
    // 内部所有路径只返回「原始按行合并结果」，最终统一 processRemarkColumn，保证多源/跨端语义一致
    merge(sources, mapping, opts, onProgress) {
      const remoteSources = sources.filter(s => remoteSet.has(s.wbId));
      const localSources = sources.filter(s => !remoteSet.has(s.wbId));
      const runLocalRaw = () => {
        if (worker) {
          const id = ++seq;
          return new Promise(function (resolve, reject) {
            pending[id] = { resolve: resolve, reject: reject, onProgress: onProgress };
            worker.postMessage({ type: 'merge', reqId: id, sources: localSources, mapping: mapping, staffEnable: opts.staffEnable, staffList: opts.staffList });
          }).then(m => m.data || []);
        }
        const real = localSources
          .filter(s => wbStore[s.wbId] && wbStore[s.wbId].Sheets[s.sheet])
          .map(s => ({ workbook: wbStore[s.wbId], sheet: s.sheet, headers: s.headers || [] }));
        if (!real.length) return Promise.reject(new Error('没有可合并的数据源（可能已清理）'));
        return Promise.resolve(MergeCore.buildMergedRows(XLSX, real, mapping, { staffEnable: opts.staffEnable, staffList: opts.staffList || [] }, onProgress));
      };

      let rawPromise;
      if (!remoteSources.length) {
        rawPromise = runLocalRaw();
      } else {
        // 有远程源：请求后端（原始行）+ 本地源（原始行），拼接后统一处理备注
        const remoteDone = backendMerge(remoteSources, mapping, opts);
        const localDone = localSources.length ? runLocalRaw() : Promise.resolve([]);
        rawPromise = localDone.then(function (localRaw) {
          return remoteDone.then(function (remoteRaw) { return localRaw.concat(remoteRaw); });
        });
      }
      return rawPromise.then(raw => MergeCore.processRemarkColumn(raw));
    },

    // 释放 wbId（本地 tempId）：后端/worker/本地分别清理
    cleanup(wbId) {
      delete wbStore[wbId];
      if (remoteSet.has(wbId)) {
        const serverId = remoteMap[wbId];
        remoteSet.delete(wbId);
        delete remoteMap[wbId];
        if (backendUrl && serverId != null) {
          try {
            fetch(backendUrl + '/api/cleanup', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ wbId: serverId }) }).catch(() => {});
          } catch (e) {}
        }
        return;
      }
      if (worker) worker.postMessage({ type: 'cleanup', wbId: wbId });
    }
  };
})();

(typeof self !== 'undefined' ? self : (typeof window !== 'undefined' ? window : this)).Engine = Engine;
