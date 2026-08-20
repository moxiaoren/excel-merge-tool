/* ============================================================
   Excel 表格合并工具 · app.js  v2.0.0
   四步向导 + 两套预设(字段模板/映射方案) + 源列名池(同名/独有·部分,N/M)
   + 目标字段卡(点选/拖拽/下拉补挂) + 人员过滤 + 合并导出
   处理逻辑（解析/列裁剪/格式化/过滤/备注处理/导出）与 v1 一致
   ============================================================ */

'use strict';

// ==================== 全局状态 ====================
let uploadedFiles = [];       // {name, workbook, sheetNames, selectedSheets, totalRows, largeFile, status, tempId}
let currentMapping = [];      // [{targetCol, sourceCol:[...]}]
let currentTargetColumns = [];
let allSourceColumns = [];    // 全部 sheet 去重列名（用于下拉/选择器）
let staffList = ['汪玲玲','张文','高平翠','邱慧敏','杨刚','申攀城','刘梦云','唐凯文','王成玲','方玲玲','蒋好','许修严','汤金萍','刘玲玲','王泗洪','金啸海','潘晨','朱静','吴亮','杨逍','徐攀','刘述功','古波','李積业','杨娜','武强','阳鑫','陈朋宇','余杰','石雪枫','邓梦婕','郭玉丹','刘以斌','张婷婷','侯倩','王近','张露霖','蔺艳秋','谢力笠','程星宇','陈杰','张海','关在兴','唐镜恒','殷伟','龙香莲','张彩凤','徐浩凡','谢小颖','首清华','郭金燕','赵梦婷','文嘉鑫'];
let sheetJSLoaded = false;
let activeStep = 1;
let currentPresetSheet = '';

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
        localStorage.setItem(STORAGE_KEYS.staff, JSON.stringify(staffList));
        localStorage.setItem(STORAGE_KEYS.mapping, JSON.stringify({ mapping: currentMapping, targetColumns: currentTargetColumns }));
    } catch (e) { /* ignore */ }
}
function loadConfigFromStorage() {
    try {
        const s = localStorage.getItem(STORAGE_KEYS.staff);
        if (s) { const a = JSON.parse(s); if (Array.isArray(a) && a.length) staffList = a; }
        const m = localStorage.getItem(STORAGE_KEYS.mapping);
        if (m) { const o = JSON.parse(m); if (o && Array.isArray(o.mapping)) { currentMapping = o.mapping; currentTargetColumns = Array.isArray(o.targetColumns) && o.targetColumns.length ? o.targetColumns : currentMapping.map(x => x.targetCol); } }
    } catch (e) { /* ignore */ }
}

// ==================== 自定义预设 ====================
function loadCustomPresets() { try { const s = localStorage.getItem(STORAGE_KEYS.customPresets); return s ? JSON.parse(s) : {}; } catch (e) { return {}; } }
function saveCustomPresets(p) { try { localStorage.setItem(STORAGE_KEYS.customPresets, JSON.stringify(p)); } catch (e) {} }

// ==================== 内置预设 ====================
const PRESET_MAPPINGS = {
    '质检映射预设': [
        { targetCol: "质检时间", sourceCol: ["日期", "审核时间"] },
        { targetCol: "sequenceld ID", sourceCol: ["sequenceId", "sequenceld ID （音视频用）"] },
        { targetCol: "业务产品", sourceCol: ["业务产品"] },
        { targetCol: "标题", sourceCol: ["贴图", "标题", "图片URL", "标题名称"] },
        { targetCol: "内容", sourceCol: ["出错内容", "内容", "替换图片url", "资讯URL", "URL"] },
        { targetCol: "原审核人员", sourceCol: ["审核人员", "原审核员", "原审核人姓名", "原审核人"] },
        { targetCol: "原审核结果", sourceCol: ["原审核状态", "原审核结果", "审核误判/ 原审核结果"] },
        { targetCol: "质检人员", sourceCol: ["质检人员", "质检员", "质检员姓名"] },
        { targetCol: "质检结果", sourceCol: ["质检判定垃圾类型", "垃圾类型", "人工高敏类别"] },
        { targetCol: "危险等级", sourceCol: ["危险等级"] }
    ],
    '敏感词映射预设': [
        { targetCol: "标题敏感词", sourceCol: ["标题敏感词"] },
        { targetCol: "敏感词", sourceCol: ["敏感词"] },
        { targetCol: "业务产品", sourceCol: ["业务产品", "业务组"] },
        { targetCol: "业务类型", sourceCol: ["业务类型"] },
        { targetCol: "作用目标", sourceCol: ["作用目标"] },
        { targetCol: "执行动作", sourceCol: ["执行动作"] },
        { targetCol: "敏感词类型", sourceCol: ["敏感词类型"] },
        { targetCol: "敏感词分类", sourceCol: ["敏感词分类"] },
        { targetCol: "原审核人员", sourceCol: ["添加人姓名", "添加人工号"] },
        { targetCol: "生效时间", sourceCol: ["生效时间"] },
        { targetCol: "失效时间", sourceCol: ["失效时间"] },
        { targetCol: "备注", sourceCol: ["备注"] }
    ],
    '回查映射预设': [
        { targetCol: "业务方", sourceCol: ["业务方"] },
        { targetCol: "被回查人", sourceCol: ["被回查人"] },
        { targetCol: "标题备注", sourceCol: ["标题备注", "标题"] },
        { targetCol: "回查人", sourceCol: ["回查人"] },
        { targetCol: "内容", sourceCol: ["内容"] },
        { targetCol: "sequence ID", sourceCol: ["sequenceld ID", "sequenceId", "sequence ID"] },
        { targetCol: "隐晦红线备注", sourceCol: ["隐晦红线备注"] }
    ],
    // —— 字段模板（仅目标列，同名自动匹配） ——
    '质检目标列预设': ['质检时间', 'sequenceld ID', '业务产品', '标题', '内容', '原审核人员', '原审核结果', '质检人员', '质检结果', '危险等级'],
    '敏感词目标列映射': ['标题敏感词', '敏感词', '业务产品', '业务类型', '作用目标', '执行动作', '敏感词类型', '敏感词分类', '原审核人员', '生效时间', '失效时间', '备注'],
    '回查目标列映射': ['业务方', '被回查人', '标题备注', '回查人', '内容', 'sequence ID', '隐晦红线备注'],
    '文本通用': ['业务产品', '抄送时间', 'sequenceId', '标题', '内容', 'AI垃圾类型', '垃圾类型', '命中原因', '审核人工号'],
    '图片通用': ['业务产品', '抄送时间', 'sequenceId', '图片URL', '替换图片url', 'AI垃圾类型', '垃圾类型', '命中原因', '审核人工号'],
    'AI反馈-文本': ['标题', '内容', 'sequenceId', '垃圾类型', 'AI垃圾类型', '业务产品'],
    'AI反馈-图片': ['sequenceId', '图片URL', '替换图片url', '垃圾类型', 'AI垃圾类型', '业务产品'],
    '帖子举报': ['抄送时间', 'sequenceId', '举报内容', '标题名称', '备注', '辅助信息', '处理状态'],
    '文本举报': ['抄送时间', 'sequenceId', '举报内容', '标题名称', '疑似类型'],
    '大模型': ['抄送时间', 'sequenceId', '业务产品', '标题', '内容', '审核状态', '垃圾类型', '大模型审核状态', '大模型垃圾类型']
};

const SPECIAL_COLUMNS = { REMARK: '备注', APPEAL: '申诉/举报', STATUS: '处理状态' };

// ==================== 列裁剪辅助 ====================
function getHeaderRowIndex(sheetName) { return sheetName.includes('质检回查') ? 1 : 0; }

function getSheetHeaderRow(worksheet, sheetName) {
    if (!worksheet || !worksheet['!ref']) return [];
    const range = XLSX.utils.decode_range(worksheet['!ref']);
    const hr = getHeaderRowIndex(sheetName);
    if (range.e.r < hr) return [];
    const headers = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
        const cell = worksheet[XLSX.utils.encode_cell({ r: hr, c: c })];
        let v = cell ? cell.v : undefined;
        if (v !== undefined && v !== null) v = typeof v === 'string' ? v.replace(/[\r\n]+/g, ' ').trim() : String(v);
        headers.push(v);
    }
    return headers;
}
function getColValue(worksheet, r, c) { const cell = worksheet[XLSX.utils.encode_cell({ r: r, c: c })]; return cell ? cell.v : ''; }

// 勾选来源列表 [{fi, sheet, headers, file}]
function getSelectedSources() {
    const sources = [];
    uploadedFiles.forEach((f, fi) => {
        if (f.status !== 'loaded') return;
        const sheets = f.selectedSheets && f.selectedSheets.length ? f.selectedSheets : (f.sheetNames.length ? [f.sheetNames[0]] : []);
        sheets.forEach(sheet => {
            const headers = getSheetHeaderRow(f.workbook.Sheets[sheet], sheet);
            if (headers.length) sources.push({ fi, sheet, headers, file: f });
        });
    });
    return sources;
}

function collectAllSourceColumns() {
    allSourceColumns = [];
    uploadedFiles.forEach(f => {
        if (f.status !== 'loaded') return;
        f.sheetNames.forEach(sheet => getSheetHeaderRow(f.workbook.Sheets[sheet], sheet).forEach(h => { if (h && !allSourceColumns.includes(h)) allSourceColumns.push(h); }));
    });
}

// ==================== 工具函数 ====================
function $(id) { return document.getElementById(id); }

function toggleCollapse(el) { if (el) el.classList.toggle('expanded'); }

// ==================== 向导切换 ====================
function goToStep(n) {
    activeStep = n;
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
    if (!sheetJSLoaded) { alert('正在加载必要的库，请稍后再试'); return; }
    Array.from(files).forEach(file => {
        if (!file.name.match(/\.(xlsx|xls|csv)$/i)) { alert('文件 ' + file.name + ' 不是有效的 Excel 或 CSV 文件'); return; }
        if (uploadedFiles.some(f => f.name === file.name)) { alert('文件 ' + file.name + ' 已经上传过了'); return; }
        const tempId = Date.now() + Math.random();
        uploadedFiles.push({ name: file.name, status: 'loading', tempId });
        renderFilesAndSheets();
        setTimeout(() => {
            if (file.name.match(/\.csv$/i)) readCSVFile(file, tempId);
            else {
                const reader = new FileReader();
                reader.onload = e => {
                    try {
                        const data = new Uint8Array(e.target.result);
                        const wb = XLSX.read(data, { type: 'array', cellStyles: false, sheetStubs: false });
                        cleanWorkbook(wb);
                        finalizeFile(file.name, tempId, wb, wb.SheetNames, file.size);
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
            const wb = XLSX.read(e.target.result, { type: 'string', codepage: 936, raw: true, cellStyles: false, sheetStubs: false });
            cleanWorkbook(wb);
            finalizeFile(file.name, tempId, wb, wb.SheetNames, file.size);
        } catch (err) { markFileError(tempId, file.name, err.message); }
    };
    try { reader.readAsText(file, 'GBK'); } catch (e) { reader.readAsText(file); }
}

function cleanWorkbook(wb) {
    wb.SheetNames.forEach(sn => {
        const ws = wb.Sheets[sn];
        if (!ws || !ws['!ref']) return;
        const range = XLSX.utils.decode_range(ws['!ref']);
        const dr = new Set(), dc = new Set();
        for (let c in ws) { if (c[0] === '!') continue; const a = XLSX.utils.decode_cell(c); dr.add(a.r); dc.add(a.c); }
        if (!dr.size || !dc.size) return;
        ws['!ref'] = XLSX.utils.encode_range({ s: { r: Math.min(...dr), c: Math.min(...dc) }, e: { r: Math.max(...dr), c: Math.max(...dc) } });
        for (let c in ws) { if (c[0] === '!') continue; ws[c] = { v: ws[c].v, t: ws[c].t }; }
    });
}

function finalizeFile(name, tempId, wb, sheetNames, size) {
    let totalRows = 0;
    sheetNames.forEach(sn => { const ws = wb.Sheets[sn]; if (ws && ws['!ref']) { const r = XLSX.utils.decode_range(ws['!ref']); totalRows += (r.e.r - r.s.r + 1); } });
    const idx = uploadedFiles.findIndex(f => f.tempId === tempId);
    if (idx !== -1) {
        uploadedFiles[idx] = { name, workbook: wb, sheetNames, totalRows, largeFile: size > LARGE_FILE_MB * 1024 * 1024, selectedSheets: sheetNames.length ? [sheetNames[0]] : [], status: 'loaded' };
        renderFilesAndSheets();
        collectAllSourceColumns();
        buildSourcePool();
        renderFieldCards();
    }
}
function markFileError(tempId, name, msg) {
    const idx = uploadedFiles.findIndex(f => f.tempId === tempId);
    if (idx !== -1) { uploadedFiles[idx].status = 'error'; uploadedFiles[idx].error = msg; renderFilesAndSheets(); }
    alert('读取文件 ' + name + ' 时出错: ' + msg);
}

function renderFilesAndSheets() {
    const c = $('fileInfo');
    c.innerHTML = '';
    if (!uploadedFiles.length) return;
    uploadedFiles.forEach((f, fi) => {
        const card = document.createElement('div');
        card.className = 'file-card' + (f.status === 'loading' ? ' loading' : '');
        if (f.status === 'loading') {
            card.innerHTML = '<h4>' + escapeHtml(f.name) + '</h4><div class="file-meta">加载中…</div>';
        } else if (f.status === 'error') {
            card.innerHTML = '<button class="close-btn" data-del="' + fi + '">×</button><h4>' + escapeHtml(f.name) + '</h4><div class="file-warning" style="color:#e2595f">加载失败</div>';
        } else {
            const big = (f.largeFile || f.totalRows > LARGE_FILE_ROWS);
            const warnTxt = (f.largeFile ? '⚠ 较大文件 (' + (f.size / 1024 / 1024).toFixed(1) + 'MB)' : '') + (f.totalRows > LARGE_FILE_ROWS ? ' ⚠ 数据较多，处理可能较慢' : '');
            card.innerHTML =
                '<button class="close-btn" data-del="' + fi + '">×</button>' +
                '<h4>' + escapeHtml(f.name) + '</h4>' +
                '<div class="file-meta">' + f.sheetNames.length + ' 个工作表 · 约 ' + (f.totalRows || 0).toLocaleString() + ' 行</div>' +
                (big ? '<div class="file-warning">' + warnTxt + '</div>' : '') +
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
    const f = uploadedFiles[fi];
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
    if (index < 0 || index >= uploadedFiles.length) return;
    uploadedFiles.splice(index, 1);
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
        html += '<div class="pool-tags-sm">' + occ.map(o => '<span class="pool-tag purple" data-col="' + escapeHtml(col) + '" title="点选建字段">' + escapeHtml(col) + ' <span class="src-badge">' + escapeHtml(uploadedFiles[o.fi].name) + '/' + escapeHtml(o.sheet) + '</span></span>').join('') + '</div>';
        html += '<div class="sec-note pool-hint" style="font-size:11px;color:var(--amber);margin-bottom:4px">缺失来源该字段自动留空</div></div>';
    });

    // 按 文件→sheet 分组
    const byFile = {};
    rest.forEach(col => colSrc[col].forEach(o => { if (!byFile[o.fi]) byFile[o.fi] = []; if (!byFile[o.fi].includes(col)) byFile[o.fi].push(col); }));
    Object.keys(byFile).sort((a, b) => a - b).forEach(fiStr => {
        const fi = +fiStr, f = uploadedFiles[fi];
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

function isColUsed(col) { return currentMapping.some(m => m.sourceCol.includes(col)); }

function pickColumn(col) {
    if (!col) return;
    const ex = currentMapping.find(m => m.targetCol === col);
    if (ex) { if (!ex.sourceCol.includes(col)) ex.sourceCol.push(col); renderFieldCards(); return; }
    currentMapping.push({ targetCol: col, sourceCol: [col] });
    if (!currentTargetColumns.includes(col)) currentTargetColumns.push(col);
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
    if (!currentMapping.length) { list.innerHTML = '<span class="pool-empty">从左侧源列名池点选列名，或载入预设开始。</span>'; return; }
    const M = getSelectedSources().length;

    currentMapping.forEach((map, i) => {
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
            '<option value="">＋ 从源列中挂列…</option>' + allSourceColumns.map(c => '<option value="' + escapeHtml(c) + '">' + escapeHtml(c) + '</option>').join('') +
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
            currentMapping.splice(i, 1);
            currentTargetColumns = currentMapping.map(m => m.targetCol);
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
    // 字段模板下拉
    const fs = $('fieldPresetSelect');
    fs.innerHTML = '';
    ['质检目标列预设', '敏感词目标列映射', '回查目标列映射', '文本通用', '图片通用', 'AI反馈-文本', 'AI反馈-图片', '帖子举报', '文本举报', '大模型'].forEach(n => { const o = document.createElement('option'); o.value = n; o.textContent = n; fs.appendChild(o); });
    const presets = loadCustomPresets();
    Object.keys(presets).forEach(nm => { const o = document.createElement('option'); o.value = '⭐' + nm; o.textContent = '⭐ ' + nm + (presets[nm].sheet ? ' · ' + presets[nm].sheet : ''); fs.appendChild(o); });
}
function renderPresetSelectMap() {
    const ms = $('mapPresetSelect');
    ms.innerHTML = '';
    ['质检映射预设', '敏感词映射预设', '回查映射预设'].forEach(n => { const o = document.createElement('option'); o.value = n; o.textContent = n; ms.appendChild(o); });
    const presets = loadCustomPresets();
    // 自定义预设也能从映射方案加载（含 sheet 绑定），但这里保留内置，自定义在字段模板下拉
}

function applyFieldPreset(key) {
    let mapping;
    if (key.indexOf('⭐') === 0) {
        const p = loadCustomPresets()[key.slice(1)];
        if (!p) { alert('预设不存在'); return; }
        currentPresetSheet = p.sheet || '';
        mapping = Array.isArray(p) ? p : (p.mapping || []);
    } else {
        const pd = PRESET_MAPPINGS[key];
        if (Array.isArray(pd) && pd.length && typeof pd[0] === 'string') mapping = pd.map(tc => ({ targetCol: tc, sourceCol: [tc] }));
        else mapping = pd || [];
    }
    if (!mapping.length) { alert('请选择有效的预设'); return; }
    currentMapping = JSON.parse(JSON.stringify(mapping));
    currentTargetColumns = currentMapping.map(m => m.targetCol);
    renderFieldCards(); buildSourcePool(); saveConfigToStorage();
    showStatus('已应用字段模板「' + key.replace(/^⭐/, '') + '」');
}

function applyMapPreset(key) {
    const pd = PRESET_MAPPINGS[key];
    if (!Array.isArray(pd) || !pd.length || typeof pd[0] === 'string') { alert('请选择有效的映射预设'); return; }
    currentMapping = JSON.parse(JSON.stringify(pd));
    currentTargetColumns = currentMapping.map(m => m.targetCol);
    currentPresetSheet = '';
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
    if (!currentMapping.length) { alert('当前没有可保存的映射配置'); return; }
    const name = prompt('请输入预设名称：');
    if (!name || !name.trim()) return;
    const sheet = prompt('关联工作表名（可空，如：质检回查）：', currentPresetSheet || '');
    const presets = loadCustomPresets();
    presets[name.trim()] = { mapping: JSON.parse(JSON.stringify(currentMapping)), sheet: (sheet && sheet.trim()) ? sheet.trim() : '' };
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
    if (!currentMapping.length) { c.innerHTML = '<div class="preview-note">请先在②设置映射或应用预设。</div>'; return; }
    html += '<div class="table-responsive"><table><thead><tr><th>来源（文件 / 工作表）</th>';
    currentMapping.forEach(m => { html += '<th>' + escapeHtml(m.targetCol) + '</th>'; });
    html += '</tr></thead><tbody>';
    sources.forEach(s => {
        html += '<tr><td>' + escapeHtml(uploadedFiles[s.fi].name) + ' / ' + escapeHtml(s.sheet) + '</td>';
        currentMapping.forEach(m => {
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
        const ws = uploadedFiles[so.fi].workbook.Sheets[so.sheet];
        if (!ws || !ws['!ref']) return sum;
        const r = XLSX.utils.decode_range(ws['!ref']); const hr = getHeaderRowIndex(so.sheet);
        return sum + Math.max(0, r.e.r - (hr + 1) + 1);
    }, 0);
    const big = s.some(so => so.file.largeFile || (so.file.totalRows || 0) > LARGE_FILE_ROWS);
    $('exportSummary').innerHTML =
        '<div class="row"><span class="k">文件 / 工作表</span><span class="v">' + s.length + ' 个来源 (' + s.map(x => escapeHtml(uploadedFiles[x.fi].name)).filter((v, i, a) => a.indexOf(v) === i).join('、') + ')</span></div>' +
        '<div class="row"><span class="k">目标字段</span><span class="v">' + (currentMapping.length ? currentMapping.map(m => escapeHtml(m.targetCol)).join('、') : '（未设置）') + '</span></div>' +
        '<div class="row"><span class="k">预计行数</span><span class="v">约 ' + rows.toLocaleString() + ' 行 ' + (big ? '<span class="warn">⚠ 数据量较大，处理稍慢</span>' : '<span class="ok">✓</span>') + '</span></div>';
    $('mergeAndExportBtn').disabled = !s.length || !currentMapping.length;
}

function mergeAndExport() {
    if (!uploadedFiles.length || !getSelectedSources().length) { alert('请先上传文件并勾选工作表'); return; }
    if (!currentMapping.length) { alert('请先设置映射关系或应用预设'); return; }
    const sources = getSelectedSources();
    const total = sources.reduce((s, so) => s + (so.file.totalRows || 0), 0);
    const big = sources.some(so => so.file.largeFile || (so.file.totalRows || 0) > LARGE_FILE_ROWS);
    if (big && !confirm('检测到较大数据（合计约 ' + total.toLocaleString() + ' 行）。合并导出可能需要一些时间，是否继续？')) return;

    const staffEnable = $('staffFilterToggle') && $('staffFilterToggle').checked;

    const sb = $('statusBar'), pb = $('progressBar'), sm = $('statusMessage');
    sb.style.display = 'block'; pb.style.width = '0%'; sm.textContent = '正在合并数据…';

    setTimeout(() => {
        try {
            const mergedData = [];
            sources.forEach((s, si) => {
                const ws = s.file.workbook.Sheets[s.sheet];
                if (!ws || !ws['!ref']) return;
                const range = XLSX.utils.decode_range(ws['!ref']);
                const hr = getHeaderRowIndex(s.sheet), start = hr + 1, last = range.e.r;
                if (last < start) return;
                const headers = getSheetHeaderRow(ws, s.sheet);
                const idx = currentMapping.map(m => m.sourceCol.map(sc => headers.indexOf(sc)).filter(x => x >= 0));
                const span = last - start + 1;
                for (let r = start; r <= last; r++) {
                    const row = { '来源sheet名': s.sheet };
                    let ok = true, empty = true;
                    currentMapping.forEach((m, mi) => {
                        let v = '';
                        for (const ci of idx[mi]) { const cv = getColValue(ws, r, ci); if (cv !== undefined && cv !== null && cv !== '') { v = cv; empty = false; break; } }
                        if ((/日期|时间/.test(m.targetCol)) && v === '-1') v = '永久';
                        else if (/日期|时间/.test(m.targetCol)) v = formatDate(v, s.sheet);
                        else if (/人员|姓名/.test(m.targetCol)) { v = extractChineseName(v); if (staffEnable && !staffList.includes(v)) ok = false; }
                        else if (m.targetCol === '疑似类型') { const i = String(v).indexOf('违规备注：'); if (i !== -1) v = String(v).substring(i + 5).trim(); }
                        row[m.targetCol] = v;
                    });
                    if (!empty && ok) mergedData.push(row);
                    pb.style.width = Math.min(100, Math.floor(((r - start + 1) / span + si) / sources.length * 100)) + '%';
                }
            });
            const processed = processRemarkColumn(mergedData);
            exportToExcel(processed);
            sm.textContent = '合并完成！文件已下载。';
            setTimeout(() => { sb.style.display = 'none'; pb.style.width = '0%'; }, 4000);
        } catch (e) { console.error(e); sm.textContent = '合并出错：' + e.message; }
    }, 60);
}

function processRemarkColumn(data) {
    const hasRemark = data.some(r => r.hasOwnProperty(SPECIAL_COLUMNS.REMARK));
    if (!hasRemark) return data;
    let hasAppeal = false;
    const out = data.map(row => {
        if (row[SPECIAL_COLUMNS.REMARK]) {
            const remark = String(row[SPECIAL_COLUMNS.REMARK]);
            const bi = remark.indexOf('】');
            if (bi !== -1) row[SPECIAL_COLUMNS.REMARK] = remark.substring(bi + 1).trim();
            const ms = remark.match(/【(.*?)】/g);
            if (ms && ms.length) { hasAppeal = true; let ac = ms[0].replace(/【|】/g, ''); ac = ac.replace('帖子违规', '帖子举报'); row[SPECIAL_COLUMNS.APPEAL] = ac; }
        }
        Object.keys(row).forEach(k => { if (typeof row[k] === 'string') row[k] = row[k].trim(); });
        return row;
    });
    if (!hasAppeal) return out;
    if (!out.some(r => r.hasOwnProperty(SPECIAL_COLUMNS.STATUS))) return out;
    const f = out.filter(r => (r[SPECIAL_COLUMNS.APPEAL] === '帖子举报' && r[SPECIAL_COLUMNS.STATUS] === '通过') || (r[SPECIAL_COLUMNS.APPEAL] === '帖子申诉' && r[SPECIAL_COLUMNS.STATUS] === '不通过'));
    f.forEach(r => delete r[SPECIAL_COLUMNS.STATUS]);
    return f;
}

function formatDate(dateStr, sheetName) {
    if (dateStr === "-1") return "永久";
    if (sheetName && sheetName.includes('质检回查')) {
        const m = String(dateStr).match(/^(\d{1,2})[\/\.](\d{1,2})$/);
        if (m) return m[1].padStart(2, '0') + '-' + m[2].padStart(2, '0');
        return dateStr;
    }
    if (typeof dateStr === 'string') {
        const d = new Date(dateStr.split(' ')[0]);
        if (!isNaN(d.getTime())) return (d.getMonth() + 1 + '').padStart(2, '0') + '-' + (d.getDate() + '').padStart(2, '0');
    }
    const num = parseFloat(dateStr);
    if (!isNaN(num)) { const d = new Date(new Date(1899, 11, 31).getTime() + (num - 1) * 24 * 60 * 60 * 1000); return (d.getMonth() + 1 + '').padStart(2, '0') + '-' + (d.getDate() + '').padStart(2, '0'); }
    const m = String(dateStr).match(/^(\d{1,2})[\/\.](\d{1,2})$/);
    if (m) return m[1].padStart(2, '0') + '-' + m[2].padStart(2, '0');
    return dateStr;
}
function extractChineseName(str) { if (!str) return ''; return String(str).replace(/[^\u4e00-\u9fa5]/g, ''); }

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
    staffList.forEach(s => {
        const t = document.createElement('span');
        t.className = 'staff-tag';
        t.innerHTML = escapeHtml(s) + ' <button data-rm="' + escapeHtml(s) + '">×</button>';
        t.querySelector('button').addEventListener('click', () => removeStaffMember(s));
        box.appendChild(t);
    });
    $('staffCount').textContent = '共 ' + staffList.length + ' 名人员';
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
        ta.id = 'staffNameInput'; ta.value = staffList.join(', ');
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
    if (name && !staffList.includes(name)) { staffList.push(name); saveConfigToStorage(); }
    toggleStaffEdit(); renderStaffList();
}
function removeStaffMember(name) { staffList = staffList.filter(s => s !== name); saveConfigToStorage(); renderStaffList(); }
function saveStaffList() {
    const ta = $('staffNameInput');
    const names = ta.value.split(/[,，\n]/).map(s => s.trim()).filter(Boolean);
    if (!names.length) { alert('请输入至少一个姓名'); return; }
    staffList = names; saveConfigToStorage(); toggleStaffEdit(); renderStaffList();
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
    $('nextStepBtn').addEventListener('click', () => goToStep(Math.min(4, activeStep + 1)));
    $('prevStepBtn').addEventListener('click', () => goToStep(Math.max(1, activeStep - 1)));

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

    sheetJSLoaded = true;
    goToStep(1);
}

function newField(name) {
    if (currentMapping.some(m => m.targetCol === name)) { pickColumn(name); return; }
    currentMapping.push({ targetCol: name, sourceCol: [name] });
    if (!currentTargetColumns.includes(name)) currentTargetColumns.push(name);
    renderFieldCards(); buildSourcePool(); saveConfigToStorage();
}

if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
else init();
