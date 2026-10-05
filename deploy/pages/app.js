/* ============================================================
   Excel 表格合并工具 · app.js  v3.0.0
   四步向导 + 两套预设(字段模板/映射方案) + 源列名池(同名/独有·部分,N/M)
   + 目标字段卡(点选/拖拽/下拉补挂) + 人员过滤 + 合并导出
   v3.0.0：处理逻辑（解析/列裁剪/格式化/过滤/备注处理/合并）下沉至 core.js，
   解析与合并由 engine.js 调度到 Web Worker 执行（file:// 不支持时自动降级主线程），
   主线程仅持有轻量元数据，避免大文件驻留导致 UI 卡死。
   ============================================================ */

'use strict';

// ==================== 全局状态（单一 state 对象，统一数据流出口） ====================
// 所有模块级可变状态收敛于此，避免散落全局变量导致的数据流失控
const state = {
  uploadedFiles: [],       // {name, wbId, sheetNames, selectedSheets, totalRows, sheetRows, sheetHeaders, largeFile, status, tempId}
  currentMapping: [],      // [{targetCol, sourceCol:[...]}]
  currentTargetColumns: [],
  allSourceColumns: [],    // 全部 sheet 去重列名（用于下拉/选择器）
  staffList: [],           // 默认空名单；此前硬编码了大量真实姓名，涉及隐私，已移除，由用户通过界面维护
  sheetJSLoaded: false,
  activeStep: 1,
  currentPresetSheet: ''
};

const LARGE_FILE_MB = 10;
const LARGE_FILE_ROWS = 50000;

const STORAGE_KEYS = {
    staff: 'excelMergeTool_staff',
    mapping: 'excelMergeTool_mapping',
    customPresets: 'excelMergeTool_customPresets'
};

function escapeHtml(s) {
    return String(s ?? '').replace(/[&<>"']/g, c =>
        ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

// ==================== 持久化 ====================
function saveConfigToStorage() {
    try {
        localStorage.setItem(STORAGE_KEYS.staff, JSON.stringify(state.staffList));
        localStorage.setItem(STORAGE_KEYS.mapping, JSON.stringify({ mapping: state.currentMapping, targetColumns: state.currentTargetColumns }));
    } catch (e) { /* ignore */ }
}
function loadConfigFromStorage() {
    try {
        const s = localStorage.getItem(STORAGE_KEYS.staff);
        if (s) { const a = JSON.parse(s); if (Array.isArray(a) && a.length) state.staffList = a; }
        const m = localStorage.getItem(STORAGE_KEYS.mapping);
        if (m) { const o = JSON.parse(m); if (o && Array.isArray(o.mapping)) { state.currentMapping = o.mapping; state.currentTargetColumns = Array.isArray(o.targetColumns) && o.targetColumns.length ? o.targetColumns : state.currentMapping.map(x => x.targetCol); } }
    } catch (e) { /* ignore */ }
}

// ==================== 自定义预设 ====================
function loadCustomPresets() { try { const s = localStorage.getItem(STORAGE_KEYS.customPresets); return s ? JSON.parse(s) : {}; } catch (e) { return {}; } }
function saveCustomPresets(p) { try { localStorage.setItem(STORAGE_KEYS.customPresets, JSON.stringify(p)); } catch (e) {} }

// ==================== 内置预设（定义与数据统一在 core.js，此处仅引用） ====================
const PRESET_MAPPINGS = MergeCore.PRESET_MAPPINGS;
const FIELD_PRESET_ORDER = MergeCore.FIELD_PRESET_ORDER;
const MAP_PRESET_ORDER = MergeCore.MAP_PRESET_ORDER;

// ==================== 来源收集（基于轻量元数据 sheetHeaders，不触碰 workbook） ====================
// 勾选来源列表 [{fi, sheet, headers, file, wbId}]
function getSelectedSources() {
    const sources = [];
    state.uploadedFiles.forEach((f, fi) => {
        if (f.status !== 'loaded') return;
        const sheets = f.selectedSheets && f.selectedSheets.length ? f.selectedSheets : (f.sheetNames.length ? [f.sheetNames[0]] : []);
        sheets.forEach(sheet => {
            const headers = (f.sheetHeaders && f.sheetHeaders[sheet]) || [];
            if (headers.length) sources.push({ fi, sheet, headers, file: f, wbId: f.wbId });
        });
    });
    return sources;
}

function collectAllSourceColumns() {
    state.allSourceColumns = [];
    state.uploadedFiles.forEach(f => {
        if (f.status !== 'loaded' || !f.sheetHeaders) return;
        f.sheetNames.forEach(sheet => (f.sheetHeaders[sheet] || []).forEach(h => { if (h && !state.allSourceColumns.includes(h)) state.allSourceColumns.push(h); }));
    });
}

// ==================== 工具函数 ====================
function $(id) { return document.getElementById(id); }

function toggleCollapse(el) { if (el) el.classList.toggle('expanded'); }

// ==================== 向导切换 ====================
function goToStep(n) {
    state.activeStep = n;
    document.querySelectorAll('.step-panel').forEach(el => el.classList.remove('active'));
    const page = (n === 1) ? $('step-upload') : (n === 2) ? $('step-mapping') : (n === 3) ? $('step-preview') : $('step-export');
    if (page) page.classList.add('active');

    document.querySelectorAll('.step-item').forEach((el, i) => {
        const s = i + 1;
        el.classList.toggle('active', s === n);
        el.classList.toggle('done', s < n);
    });
    // 进度线
    const lines = document.querySelectorAll('.step-line i');
    lines.forEach((ln, i) => { ln.style.width = (i + 1 <= n - 1) ? '100%' : '0%'; });

    $('prevStepBtn').disabled = (n === 1);
    const next = $('nextStepBtn');
    next.textContent = (n === 4) ? '完成' : '下一步 →';

    if (n === 3) renderPreview();
    if (n === 4) renderExportSummary();
    window.scrollTo({ top: 0, behavior: 'smooth' });
}

// ==================== 文件上传 & Sheet 勾选 ====================
function initUpload() {
    $('uploadArea').addEventListener('click', () => $('fileInput').click());
    $('fileInput').addEventListener('change', handleFileSelect);
    const a = $('uploadArea');
    a.addEventListener('dragover', e => { e.preventDefault(); a.classList.add('dragover'); });
    a.addEventListener('dragleave', () => a.classList.remove('dragover'));
    a.addEventListener('drop', e => { e.preventDefault(); a.classList.remove('dragover'); handleFileSelect(e); });
}

function handleFileSelect(event) {
    const files = event.target.files || event.dataTransfer.files;
    if (!files.length) return;
    if (!state.sheetJSLoaded) { alert('正在加载必要的库，请稍后再试'); return; }
    Array.from(files).forEach(file => {
        if (!file.name.match(/\.(xlsx|xls|csv)$/i)) { alert('文件 ' + file.name + ' 不是有效的 Excel 或 CSV 文件'); return; }
        if (state.uploadedFiles.some(f => f.name === file.name)) { alert('文件 ' + file.name + ' 已经上传过了'); return; }
        const tempId = Date.now() + Math.random();
        state.uploadedFiles.push({ name: file.name, status: 'loading', tempId });
        renderFilesAndSheets();
        setTimeout(() => {
            if (file.name.match(/\.csv$/i)) readCSVFile(file, tempId);
            else {
                const reader = new FileReader();
                reader.onload = e => {
                    try {
                        const data = new Uint8Array(e.target.result);
                        Engine.parse(data, false, tempId, file.size).then(meta => finalizeFile(file.name, tempId, meta, file.size))
                            .catch(err => markFileError(tempId, file.name, (err && err.message) || String(err)));
                    } catch (err) { markFileError(tempId, file.name, err.message); }
                };
                reader.readAsArrayBuffer(file);
            }
        }, 30);
    });
    $('fileInput').value = '';
}

function readCSVFile(file, tempId) {
    const reader = new FileReader();
    reader.onload = e => {
        try {
            Engine.parse(e.target.result, true, tempId).then(meta => finalizeFile(file.name, tempId, meta, file.size))
                .catch(err => markFileError(tempId, file.name, (err && err.message) || String(err)));
        } catch (err) { markFileError(tempId, file.name, err.message); }
    };
    try { reader.readAsText(file, 'GBK'); } catch (e) { reader.readAsText(file); }
}

// 解析完成后：只保留轻量元数据（sheetNames/totalRows/sheetHeaders/sheetRows），
// 大 workbook 对象留在 Worker（或降级时的 engine 内部），绝不让大对象驻留主线程 UI 态。
function finalizeFile(name, tempId, meta, size) {
    const idx = state.uploadedFiles.findIndex(f => f.tempId === tempId);
    if (idx !== -1) {
        state.uploadedFiles[idx] = {
            name, wbId: tempId,
            sheetNames: meta.sheetNames, totalRows: meta.totalRows,
            sheetHeaders: meta.sheetHeaders, sheetRows: meta.sheetRows,
            largeFile: size > LARGE_FILE_MB * 1024 * 1024,
            selectedSheets: meta.sheetNames.length ? [meta.sheetNames[0]] : [],
            status: 'loaded'
        };
        renderFilesAndSheets();
        collectAllSourceColumns();
        buildSourcePool();
        renderFieldCards();
    }
}
function markFileError(tempId, name, msg) {
    const idx = state.uploadedFiles.findIndex(f => f.tempId === tempId);
    if (idx !== -1) { state.uploadedFiles[idx].status = 'error'; state.uploadedFiles[idx].error = msg; renderFilesAndSheets(); }
    alert('读取文件 ' + name + ' 时出错: ' + msg);
}

function renderFilesAndSheets() {
    const c = $('fileInfo');
    c.innerHTML = '';
    if (!state.uploadedFiles.length) return;
    state.uploadedFiles.forEach((f, fi) => {
        const card = document.createElement('div');
        card.className = 'file-card' + (f.status === 'loading' ? ' loading' : '');
        if (f.status === 'loading') {
            card.innerHTML = '<h4>' + escapeHtml(f.name) + '</h4><div class="file-meta">加载中…</div>';
        } else if (f.status === 'error') {
            card.innerHTML = '<button class="close-btn" data-del="' + fi + '">×</button><h4>' + escapeHtml(f.name) + '</h4><div class="file-warning" style="color:#e2595f">加载失败</div>';
        } else {
            const big = (f.largeFile || f.totalRows > LARGE_FILE_ROWS);
            const warnTxt = (f.largeFile ? '⚠ 较大文件 (' + (f.size / 1024 / 1024).toFixed(1) + 'MB)' : '') + (f.totalRows > LARGE_FILE_ROWS ? ' ⚠ 数据较多，处理可能较慢' : '');
            // 硬性上限：无后端且文件超大时，明确提示崩溃风险并引导措施
            const HARD_LIMIT_MB = 80;
            const noBackend = !(Engine && Engine.backendOn);
            const huge = noBackend && f.size > HARD_LIMIT_MB * 1024 * 1024;
            const hardTxt = huge
                ? '🔴 文件超过 ' + HARD_LIMIT_MB + 'MB 且未启用后端：当前纯静态解析内存峰值可能达 3-4GB，浏览器可能崩溃。建议拆分上传，或部署后端（server.js）并在 index.html 配置 BACKEND_URL。'
                : '';
            card.innerHTML =
                '<button class="close-btn" data-del="' + fi + '">×</button>' +
                '<h4>' + escapeHtml(f.name) + '</h4>' +
                '<div class="file-meta">' + f.sheetNames.length + ' 个工作表 · 约 ' + (f.totalRows || 0).toLocaleString() + ' 行</div>' +
                (big ? '<div class="file-warning">' + warnTxt + '</div>' : '') +
                (huge ? '<div class="file-warning" style="color:#e2595f;font-weight:700">' + hardTxt + '</div>' : '') +
                '<div class="sheet-list"><div class="sheet-label">参与合并的工作表（默认只勾选第一个）：</div></div>';
            const sl = card.querySelector('.sheet-list');
            f.sheetNames.forEach(sheet => {
                const on = f.selectedSheets && f.selectedSheets.includes(sheet);
                const chip = document.createElement('label');
                chip.className = 'sheet-chip' + (on ? ' checked' : '');
                chip.innerHTML = '<input type="checkbox" ' + (on ? 'checked' : '') + '><span>' + escapeHtml(sheet) + '</span>';
                chip.addEventListener('click', e => { e.preventDefault(); toggleSheet(fi, sheet); });
                sl.appendChild(chip);
            });
            card.querySelector('.close-btn').addEventListener('click', () => removeFile(fi));
        }
        c.appendChild(card);
    });
}

function toggleSheet(fi, sheet) {
    const f = state.uploadedFiles[fi];
    if (!f || f.status !== 'loaded') return;
    if (!f.selectedSheets) f.selectedSheets = f.sheetNames.length ? [f.sheetNames[0]] : [];
    const i = f.selectedSheets.indexOf(sheet);
    if (i >= 0) { if (f.selectedSheets.length === 1) { alert('至少保留一个工作表'); return; } f.selectedSheets.splice(i, 1); }
    else f.selectedSheets.push(sheet);
    renderFilesAndSheets();
    buildSourcePool();
    renderFieldCards();
}
function removeFile(index) {
    if (index < 0 || index >= state.uploadedFiles.length) return;
    const removed = state.uploadedFiles[index];
    state.uploadedFiles.splice(index, 1);
    if (removed && Engine.cleanup) Engine.cleanup(removed.wbId);   // 释放 worker 端大对象内存
    renderFilesAndSheets();
    collectAllSourceColumns();
    buildSourcePool();
    renderFieldCards();
}

// ==================== 源列名池 ====================
function buildSourcePool() {
    const sources = getSelectedSources();
    const M = sources.length;
    const sameEl = $('sameColumnList'), uniqEl = $('uniqueColumnList');

    if (!M) {
        sameEl.innerHTML = '<span class="pool-empty">请先在第 1 步上传文件并勾选工作表。</span>';
        uniqEl.innerHTML = '';
        $('sameCoverage').textContent = '—';
        $('partCoverage').textContent = '—';
        return;
    }

    // 列名 -> 出现来源
    const colSrc = {};
    sources.forEach((s, si) => s.headers.forEach(h => { if (h) { if (!colSrc[h]) colSrc[h] = []; colSrc[h].push({ fi: s.fi, sheet: s.sheet }); } }));

    const same = [], other = [];
    Object.keys(colSrc).forEach(col => (colSrc[col].length === M ? same : other).push(col));
    same.sort();

    $('sameCoverage').textContent = M + '/' + M;
    $('partCoverage').textContent = 'N/' + M;

    // 同名列
    sameEl.innerHTML = '';
    if (!same.length) sameEl.innerHTML = '<span class="pool-empty">— 无全覆盖列</span>';
    same.forEach(col => {
        const t = document.createElement('span');
        t.className = 'pool-tag green' + (isColUsed(col) ? '' : '');
        t.textContent = col;
        t.setAttribute('data-col', col);
        t.title = '所有勾选来源都有 · 点选建字段并全局自动匹配';
        t.addEventListener('click', () => pickColumn(col));
        t.draggable = true;
        t.addEventListener('dragstart', e => e.dataTransfer.setData('text/plain', 'SYNC::' + col));
        sameEl.appendChild(t);
    });

    // 独有/部分列：部分同名(2<=N<M)高亮 + 按文件/sheet分组其它
    const partCols = other.filter(c => colSrc[c].length >= 2 && colSrc[c].length < M);
    const rest = other.filter(c => !partCols.includes(c));
    let html = '';

    partCols.forEach(col => {
        const occ = colSrc[col];
        html += '<div class="pool-groups" style="border-left:3px solid var(--amber);padding-left:8px;gap:6px;">';
        html += '<div class="pool-group-label"><span class="fb" style="color:var(--amber)">⚠ ' + escapeHtml(col) + '</span><span class="coverage-badge">' + occ.length + '/' + M + ' 来源</span></div>';
        html += '<div class="pool-tags-sm">' + occ.map(o => '<span class="pool-tag purple" data-col="' + escapeHtml(col) + '" title="点选建字段">' + escapeHtml(col) + ' <span class="src-badge">' + escapeHtml(state.uploadedFiles[o.fi].name) + '/' + escapeHtml(o.sheet) + '</span></span>').join('') + '</div>';
        html += '<div class="sec-note pool-hint" style="font-size:11px;color:var(--amber);margin-bottom:4px">缺失来源该字段自动留空</div></div>';
    });

    // 按 文件→sheet 分组
    const byFile = {};
    rest.forEach(col => colSrc[col].forEach(o => { if (!byFile[o.fi]) byFile[o.fi] = []; if (!byFile[o.fi].includes(col)) byFile[o.fi].push(col); }));
    Object.keys(byFile).sort((a, b) => a - b).forEach(fiStr => {
        const fi = +fiStr, f = state.uploadedFiles[fi];
        html += '<div class="pool-groups"><div class="pool-group-label"><span class="fb">📄 ' + escapeHtml(f.name) + '</span></div>';
        const bySheet = {};
        byFile[fi].forEach(col => { const sKey = colSrc[col].find(o => o.fi === fi).sheet; if (!bySheet[sKey]) bySheet[sKey] = []; if (!bySheet[sKey].includes(col)) bySheet[sKey].push(col); });
        Object.keys(bySheet).forEach(sheet => {
            html += '<div class="pool-group-label" style="margin-top:4px">▸ <span class="fb">' + escapeHtml(sheet) + '</span></div>';
            html += '<div class="pool-tags-sm">';
            bySheet[sheet].sort().forEach(col => {
                const n = colSrc[col].length;
                html += '<span class="pool-tag purple' + (isColUsed(col) ? ' used' : '') + '" data-col="' + escapeHtml(col) + '" title="' + n + '/' + M + ' 来源 · 点选建字段">' + escapeHtml(col) + ' <span class="src-badge">' + n + '/' + M + '</span></span>';
            });
            html += '</div>';
        });
        html += '</div>';
    });

    uniqEl.innerHTML = html;
    uniqEl.querySelectorAll('.pool-tag[data-col]').forEach(t => {
        const col = t.getAttribute('data-col');
        t.addEventListener('click', () => pickColumn(col));
        t.draggable = true;
        t.addEventListener('dragstart', e => e.dataTransfer.setData('text/plain', 'SYNC::' + col));
    });
}

function isColUsed(col) { return state.currentMapping.some(m => m.sourceCol.includes(col)); }

function pickColumn(col) {
    if (!col) return;
    const ex = state.currentMapping.find(m => m.targetCol === col);
    if (ex) { if (!ex.sourceCol.includes(col)) ex.sourceCol.push(col); renderFieldCards(); return; }
    state.currentMapping.push({ targetCol: col, sourceCol: [col] });
    if (!state.currentTargetColumns.includes(col)) state.currentTargetColumns.push(col);
    renderFieldCards();
    buildSourcePool();
    saveConfigToStorage();
}

// ==================== 目标字段卡 ====================
function countCoverage(col) {
    const M = getSelectedSources().length;
    if (!M) return { n: 0, M };
    let n = 0; getSelectedSources().forEach(s => { if (s.headers.includes(col)) n++; });
    return { n, M };
}

function renderFieldCards() {
    const list = $('targetFieldList');
    list.innerHTML = '';
    if (!state.currentMapping.length) { list.innerHTML = '<span class="pool-empty">从左侧源列名池点选列名，或载入预设开始。</span>'; return; }
    const M = getSelectedSources().length;

    state.currentMapping.forEach((map, i) => {
        const card = document.createElement('div');
        card.className = 'field-card';
        card.dataset.index = i;
        let chips = '';
        map.sourceCol.forEach(col => {
            const { n } = countCoverage(col);
            let cls = 'all', lbl = escapeHtml(col);
            if (n === M) { cls = 'all'; lbl += ' ×' + M; }
            else if (n > 0) { cls = 'file'; lbl += ' ·' + n + '/' + M; }
            else { cls = 'blank'; lbl = escapeHtml(col) + '（无此列→留空）'; }
            chips += '<span class="src-chip ' + cls + '">' + lbl + '<span class="x" data-col="' + escapeHtml(col) + '">×</span></span>';
        });
        if (!map.sourceCol.length) chips = '<span class="src-chip blank" style="border:1px dashed var(--line)">尚未挂源列</span>';
        card.innerHTML =
            '<div class="field-card-head"><span class="field-name">' + escapeHtml(map.targetCol) + '</span><button class="field-del" data-del="' + i + '" title="删除字段">×</button></div>' +
            '<div class="src-chips">' + chips + '</div>' +
            '<div class="field-add"><select style="flex:1;padding:7px 9px;border:1px solid var(--line);border-radius:9px;font-size:12.5px;outline:none;background:#fff">' +
            '<option value="">＋ 从源列中挂列…</option>' + state.allSourceColumns.map(c => '<option value="' + escapeHtml(c) + '">' + escapeHtml(c) + '</option>').join('') +
            '</select><button class="btn btn-primary btn-sm" data-add="' + i + '">添加</button></div>';
        list.appendChild(card);

        // 移除源列
        card.querySelectorAll('.src-chip .x').forEach(x => x.addEventListener('click', e => {
            e.stopPropagation();
            const col = x.getAttribute('data-col');
            map.sourceCol = map.sourceCol.filter(c => c !== col);
            renderFieldCards(); saveConfigToStorage();
        }));
        // 添加源列
        card.querySelector('[data-add]').addEventListener('click', () => {
            const sel = card.querySelector('select');
            if (sel.value) { if (!map.sourceCol.includes(sel.value)) map.sourceCol.push(sel.value); renderFieldCards(); saveConfigToStorage(); }
        });
        // 删字段
        card.querySelector('[data-del]').addEventListener('click', () => {
            state.currentMapping.splice(i, 1);
            state.currentTargetColumns = state.currentMapping.map(m => m.targetCol);
            renderFieldCards(); buildSourcePool(); saveConfigToStorage();
        });
        // 拖入
        ['dragover', 'dragenter'].forEach(ev => card.addEventListener(ev, e => { e.preventDefault(); if (e.dataTransfer.types.includes('text/plain')) card.classList.add('dragover'); }));
        card.addEventListener('dragleave', () => card.classList.remove('dragover'));
        card.addEventListener('drop', e => {
            e.preventDefault(); card.classList.remove('dragover');
            const v = e.dataTransfer.getData('text/plain');
            if (v.indexOf('SYNC::') === 0) { const col = v.slice(6); if (!map.sourceCol.includes(col)) map.sourceCol.push(col); renderFieldCards(); saveConfigToStorage(); }
        });
    });
}

// ==================== 预设 ====================
function renderPresetSelect() {
    // 字段模板下拉（顺序来自 core 的 FIELD_PRESET_ORDER）
    const fs = $('fieldPresetSelect');
    fs.innerHTML = '';
    FIELD_PRESET_ORDER.forEach(n => { const o = document.createElement('option'); o.value = n; o.textContent = n; fs.appendChild(o); });
    const presets = loadCustomPresets();
    Object.keys(presets).forEach(nm => { const o = document.createElement('option'); o.value = '⭐' + nm; o.textContent = '⭐ ' + nm + (presets[nm].sheet ? ' · ' + presets[nm].sheet : ''); fs.appendChild(o); });
}
function renderPresetSelectMap() {
    const ms = $('mapPresetSelect');
    ms.innerHTML = '';
    MAP_PRESET_ORDER.forEach(n => { const o = document.createElement('option'); o.value = n; o.textContent = n; ms.appendChild(o); });
    const presets = loadCustomPresets();
    // 自定义预设也能从映射方案加载（含 sheet 绑定），但这里保留内置，自定义在字段模板下拉
}

function applyFieldPreset(key) {
    let mapping;
    if (key.indexOf('⭐') === 0) {
        const p = loadCustomPresets()[key.slice(1)];
        if (!p) { alert('预设不存在'); return; }
        state.currentPresetSheet = p.sheet || '';
        mapping = Array.isArray(p) ? p : (p.mapping || []);
    } else {
        const pd = PRESET_MAPPINGS[key];
        if (Array.isArray(pd) && pd.length && typeof pd[0] === 'string') mapping = pd.map(tc => ({ targetCol: tc, sourceCol: [tc] }));
        else mapping = pd || [];
    }
    if (!mapping.length) { alert('请选择有效的预设'); return; }
    state.currentMapping = JSON.parse(JSON.stringify(mapping));
    state.currentTargetColumns = state.currentMapping.map(m => m.targetCol);
    renderFieldCards(); buildSourcePool(); saveConfigToStorage();
    showStatus('已应用字段模板「' + key.replace(/^⭐/, '') + '」');
}

function applyMapPreset(key) {
    const pd = PRESET_MAPPINGS[key];
    if (!Array.isArray(pd) || !pd.length || typeof pd[0] === 'string') { alert('请选择有效的映射预设'); return; }
    state.currentMapping = JSON.parse(JSON.stringify(pd));
    state.currentTargetColumns = state.currentMapping.map(m => m.targetCol);
    state.currentPresetSheet = '';
    renderFieldCards(); buildSourcePool(); saveConfigToStorage();
    showStatus('已应用映射方案「' + key + '」');
}

function showStatus(msg) {
    const sb = $('statusBar'); $('statusMessage').textContent = msg;
    sb.style.display = 'block'; sb.style.opacity = '1';
    clearTimeout(showStatus._t);
    showStatus._t = setTimeout(() => { sb.style.display = 'none'; }, 2600);
}

function openPresetManager() {
    const presets = loadCustomPresets();
    const keys = Object.keys(presets);
    const body = document.createElement('div');
    body.style.cssText = 'position:fixed;inset:0;background:rgba(20,30,60,.4);backdrop-filter:blur(3px);display:flex;align-items:center;justify-content:center;z-index:60;padding:20px';
    let html = '<div style="width:440px;max-width:100%;background:#fff;border-radius:16px;overflow:hidden;box-shadow:0 24px 60px rgba(20,30,60,.3)">' +
        '<div style="display:flex;align-items:center;font-weight:800;color:#233047;background:#f7f8ff;padding:13px 16px;border-bottom:1px solid var(--line)">🗂 管理自定义预设<button data-c style="margin-left:auto;border:none;background:none;font-size:20px;color:#94a3b8;cursor:pointer">×</button></div>' +
        '<div style="padding:16px;max-height:60vh;overflow:auto">';
    if (!keys.length) html += '<p style="color:#6b7a93;font-size:13px">还没有自定义预设。</p>';
    keys.forEach(nm => {
        html += '<div style="display:flex;align-items:center;gap:8px;border:1px solid var(--line);border-radius:9px;padding:8px 10px;margin-bottom:7px;font-size:13px">' +
            '<span style="font-weight:700;flex:1">⭐ ' + escapeHtml(nm) + (presets[nm].sheet ? ' <small style="color:#94a3b8">· ' + escapeHtml(presets[nm].sheet) + '</small>' : '') + '</span>' +
            '<button data-load="' + escapeHtml(nm) + '" style="border:1px solid var(--line);background:#fff;color:#233047;border-radius:7px;padding:4px 10px;font-size:11px;cursor:pointer">读取</button>' +
            '<button data-delp="' + escapeHtml(nm) + '" style="border:1px solid #fecaca;background:#fff;color:#e2595f;border-radius:7px;padding:4px 10px;font-size:11px;cursor:pointer">删除</button></div>';
    });
    html += '</div></div>';
    body.innerHTML = html;
    document.body.appendChild(body);
    body.querySelector('[data-c]').addEventListener('click', () => body.remove());
    body.addEventListener('click', e => { if (e.target === body) body.remove(); });
    body.querySelectorAll('[data-load]').forEach(b => b.addEventListener('click', () => { applyFieldPreset('⭐' + b.getAttribute('data-load')); body.remove(); }));
    body.querySelectorAll('[data-delp]').forEach(b => b.addEventListener('click', () => {
        if (!confirm('删除自定义预设？')) return;
        const presets2 = loadCustomPresets(); delete presets2[b.getAttribute('data-delp')]; saveCustomPresets(presets2); renderPresetSelect(); body.remove(); openPresetManager();
    }));
}

function saveCurrentAsPreset() {
    if (!state.currentMapping.length) { alert('当前没有可保存的映射配置'); return; }
    const name = prompt('请输入预设名称：');
    if (!name || !name.trim()) return;
    const sheet = prompt('关联工作表名（可空，如：质检回查）：', state.currentPresetSheet || '');
    const presets = loadCustomPresets();
    presets[name.trim()] = { mapping: JSON.parse(JSON.stringify(state.currentMapping)), sheet: (sheet && sheet.trim()) ? sheet.trim() : '' };
    saveCustomPresets(presets);
    renderPresetSelect();
    alert('预设「' + name.trim() + '」已保存');
}

// ==================== 预览 ====================
function renderPreview() {
    const c = $('previewTableContainer');
    const sources = getSelectedSources();
    let html = '';
    if (!sources.length) { c.innerHTML = '<div class="preview-note">请先在①上传文件并勾选工作表。</div>'; return; }
    if (!state.currentMapping.length) { c.innerHTML = '<div class="preview-note">请先在②设置映射或应用预设。</div>'; return; }
    html += '<div class="table-responsive"><table><thead><tr><th>来源（文件 / 工作表）</th>';
    state.currentMapping.forEach(m => { html += '<th>' + escapeHtml(m.targetCol) + '</th>'; });
    html += '</tr></thead><tbody>';
    sources.forEach(s => {
        html += '<tr><td>' + escapeHtml(state.uploadedFiles[s.fi].name) + ' / ' + escapeHtml(s.sheet) + '</td>';
        state.currentMapping.forEach(m => {
            const matched = m.sourceCol.filter(cname => s.headers.includes(cname));
            if (matched.length) html += '<td class="match">' + matched.map(x => escapeHtml(x)).join(', ') + '</td>';
            else if (m.sourceCol.length) html += '<td class="nomatch">无此列 → 留空</td>';
            else html += '<td class="nomatch">—</td>';
        });
        html += '</tr>';
    });
    html += '</tbody></table></div>';
    html += '<div class="preview-note">预览基于<b>已勾选的工作表</b>。「无此列 → 留空」表示该来源缺少对应源列，导出时此项为空，可在字段上用「添加」补挂。</div>';
    html += '<div class="legend"><span style="color:var(--green)">■ 已匹配</span><span style="color:var(--ink-sub)">■ 无此列·留空</span><span style="color:var(--violet)">■ 独有列按来源</span></div>';
    c.innerHTML = html;
}

// ==================== 导出 ====================
function renderExportSummary() {
    const s = getSelectedSources();
    const rows = s.reduce((sum, so) => {
        const rowCount = (so.file.sheetRows && so.file.sheetRows[so.sheet]) || 0;
        return sum + rowCount;
    }, 0);
    const big = s.some(so => so.file.largeFile || (so.file.totalRows || 0) > LARGE_FILE_ROWS);
    $('exportSummary').innerHTML =
        '<div class="row"><span class="k">文件 / 工作表</span><span class="v">' + s.length + ' 个来源 (' + s.map(x => escapeHtml(state.uploadedFiles[x.fi].name)).filter((v, i, a) => a.indexOf(v) === i).join('、') + ')</span></div>' +
        '<div class="row"><span class="k">目标字段</span><span class="v">' + (state.currentMapping.length ? state.currentMapping.map(m => escapeHtml(m.targetCol)).join('、') : '（未设置）') + '</span></div>' +
        '<div class="row"><span class="k">预计行数</span><span class="v">约 ' + rows.toLocaleString() + ' 行 ' + (big ? '<span class="warn">⚠ 数据量较大，处理稍慢</span>' : '<span class="ok">✓</span>') + '</span></div>';
    $('mergeAndExportBtn').disabled = !s.length || !state.currentMapping.length;
}

function mergeAndExport() {
    if (!state.uploadedFiles.length || !getSelectedSources().length) { alert('请先上传文件并勾选工作表'); return; }
    if (!state.currentMapping.length) { alert('请先设置映射关系或应用预设'); return; }
    const sources = getSelectedSources();
    const total = sources.reduce((s, so) => s + (so.file.totalRows || 0), 0);
    const big = sources.some(so => so.file.largeFile || (so.file.totalRows || 0) > LARGE_FILE_ROWS);
    if (big && !confirm('检测到较大数据（合计约 ' + total.toLocaleString() + ' 行）。合并导出可能需要一些时间，是否继续？')) return;

    const staffEnable = $('staffFilterToggle') && $('staffFilterToggle').checked;

    const sb = $('statusBar'), pb = $('progressBar'), sm = $('statusMessage');
    sb.style.display = 'block'; pb.style.width = '0%'; sm.textContent = '正在合并数据…';

    // 合并由 Engine 调度（Worker 优先 / 主线程降级），期间主线程可交互不冻结
    const payload = sources.map(s => ({ wbId: s.wbId, sheet: s.sheet, headers: s.headers }));
    Engine.merge(payload, state.currentMapping, { staffEnable: staffEnable, staffList: state.staffList }, pct => {
        pb.style.width = Math.min(100, pct) + '%';
    }).then(processed => {
        exportToExcel(processed);
        sm.textContent = '合并完成！文件已下载。';
        setTimeout(() => { sb.style.display = 'none'; pb.style.width = '0%'; }, 4000);
    }).catch(err => {
        console.error(err);
        sm.textContent = '合并出错：' + ((err && err.message) || '未知错误');
    });
}

function exportToExcel(data) {
    if (!data.length) { alert('没有数据可导出'); return; }
    const ws = XLSX.utils.json_to_sheet(data);
    ws['!cols'] = Object.keys(data[0]).map(col => ({ wch: (col === '标题' || col === '内容') ? 45 : 12 }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '合并数据');
    XLSX.writeFile(wb, '合并数据.xlsx');
}

// ==================== 人员过滤 ====================
function renderStaffList() {
    const box = $('staffDisplay');
    box.innerHTML = '';
    state.staffList.forEach(s => {
        const t = document.createElement('span');
        t.className = 'staff-tag';
        t.innerHTML = escapeHtml(s) + ' <button data-rm="' + escapeHtml(s) + '">×</button>';
        t.querySelector('button').addEventListener('click', () => removeStaffMember(s));
        box.appendChild(t);
    });
    $('staffCount').textContent = '共 ' + state.staffList.length + ' 名人员';
}

function toggleStaffEdit() {
    const display = $('staffDisplay'), inputBox = $('staffInput');
    const edit = $('editStaffBtn'), save = $('saveStaffBtn'), cancel = $('cancelStaffBtn');
    const editing = inputBox.classList.contains('active');
    if (editing) {
        display.style.display = ''; inputBox.classList.remove('active'); inputBox.innerHTML = '<input type="text" id="staffNameInput" class="input-field glass" placeholder="输入姓名后按回车">';
        edit.style.display = ''; save.classList.add('hidden'); cancel.classList.add('hidden');
        bindStaffInput();
    } else {
        display.style.display = 'none'; inputBox.classList.add('active');
        const ta = document.createElement('textarea');
        ta.id = 'staffNameInput'; ta.value = state.staffList.join(', ');
        inputBox.innerHTML = ''; inputBox.appendChild(ta);
        edit.style.display = 'none'; save.classList.remove('hidden'); cancel.classList.remove('hidden');
    }
}
function bindStaffInput() {
    const inp = $('staffNameInput');
    if (inp) inp.addEventListener('keydown', e => { if (e.key === 'Enter') addStaffMember(); });
}
function addStaffMember() {
    const inp = $('staffNameInput'); const name = inp.value.trim();
    if (name && !state.staffList.includes(name)) { state.staffList.push(name); saveConfigToStorage(); }
    toggleStaffEdit(); renderStaffList();
}
function removeStaffMember(name) { state.staffList = state.staffList.filter(s => s !== name); saveConfigToStorage(); renderStaffList(); }
function saveStaffList() {
    const ta = $('staffNameInput');
    const names = ta.value.split(/[,，\n]/).map(s => s.trim()).filter(Boolean);
    if (!names.length) { alert('请输入至少一个姓名'); return; }
    state.staffList = names; saveConfigToStorage(); toggleStaffEdit(); renderStaffList();
}
function cancelStaffEdit() { toggleStaffEdit(); }

// ==================== 初始化 ====================
function init() {
    loadConfigFromStorage();

    initUpload();
    renderFilesAndSheets();
    renderPresetSelect();
    renderPresetSelectMap();
    buildSourcePool();
    renderFieldCards();

    // 步骤条点击
    document.querySelectorAll('.step-item').forEach(el => el.addEventListener('click', () => goToStep(+el.dataset.step)));
    $('nextStepBtn').addEventListener('click', () => goToStep(Math.min(4, state.activeStep + 1)));
    $('prevStepBtn').addEventListener('click', () => goToStep(Math.max(1, state.activeStep - 1)));

    // 预设
    $('loadFieldPresetBtn').addEventListener('click', () => applyFieldPreset($('fieldPresetSelect').value));
    $('loadMapPresetBtn').addEventListener('click', () => applyMapPreset($('mapPresetSelect').value));
    $('savePresetBtn').addEventListener('click', saveCurrentAsPreset);
    $('managePresetBtn').addEventListener('click', openPresetManager);

    // 新建字段
    $('addNewFieldBtn').addEventListener('click', () => {
        const v = $('newFieldNameInput').value.trim();
        if (v) { newField(v); $('newFieldNameInput').value = ''; }
    });
    $('newFieldNameInput').addEventListener('keydown', e => { if (e.key === 'Enter') $('addNewFieldBtn').click(); });

    // 导出
    $('mergeAndExportBtn').addEventListener('click', mergeAndExport);

    // 人员
    renderStaffList();
    bindStaffInput();
    $('editStaffBtn').addEventListener('click', toggleStaffEdit);
    $('saveStaffBtn').addEventListener('click', saveStaffList);
    $('cancelStaffBtn').addEventListener('click', cancelStaffEdit);

    // 源列名池：收纳 / 展开（记忆用户偏好）
    const pool = $('sourcePool');
    const poolToggle = $('sourcePoolToggle');
    if (pool && poolToggle) {
        const applyCollapsed = collapsed => {
            pool.classList.toggle('collapsed', collapsed);
            poolToggle.textContent = collapsed ? '⏵' : '⏴';
            poolToggle.title = collapsed ? '展开源列名池' : '收起源列名池';
            try { localStorage.setItem('poolCollapsed', collapsed ? '1' : '0'); } catch (e) {}
        };
        let collapsed = false;
        try { collapsed = localStorage.getItem('poolCollapsed') === '1'; } catch (e) {}
        applyCollapsed(collapsed);
        poolToggle.addEventListener('click', () => applyCollapsed(!pool.classList.contains('collapsed')));
    }

    state.sheetJSLoaded = true;

    // 探测可选后端（纯静态部署时 backendUrl 为空，自动跳过）
    if (Engine.initBackend) Engine.initBackend();
    goToStep(1);
}

function newField(name) {
    if (state.currentMapping.some(m => m.targetCol === name)) { pickColumn(name); return; }
    state.currentMapping.push({ targetCol: name, sourceCol: [name] });
    if (!state.currentTargetColumns.includes(name)) state.currentTargetColumns.push(name);
    renderFieldCards(); buildSourcePool(); saveConfigToStorage();
}

if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
else init();
