/* ============================================================
   Excel 表格合并工具 · core.js  v3.0.0
   纯业务逻辑层：解析、表头读取、按行合并、格式化、过滤、备注处理。
   不依赖 DOM / 不依赖全局状态，可在主线程与 Web Worker 中复用。
   通过 IIFE 挂载到 self / window，便于脚本标签方式加载。
   ============================================================ */
(function (root) {
  'use strict';

  const LARGE_FILE_MB = 10;
  const LARGE_FILE_ROWS = 50000;

  const STORAGE_KEYS = {
    staff: 'excelMergeTool_staff',
    mapping: 'excelMergeTool_mapping',
    customPresets: 'excelMergeTool_customPresets'
  };

  const SPECIAL_COLUMNS = { REMARK: '备注', APPEAL: '申诉/举报', STATUS: '处理状态' };

  // 内置预设（字段模板 + 映射方案）
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

  const FIELD_PRESET_ORDER = ['质检目标列预设', '敏感词目标列映射', '回查目标列映射', '文本通用', '图片通用', 'AI反馈-文本', 'AI反馈-图片', '帖子举报', '文本举报', '大模型'];
  const MAP_PRESET_ORDER = ['质检映射预设', '敏感词映射预设', '回查映射预设'];

  // ==================== 通用工具 ====================
  function escapeHtml(s) {
    return String(s ?? '').replace(/[&<>"']/g, c =>
      ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  }

  // ==================== 表头 / 单元格读取 ====================
  // 某些业务 sheet（如质检回查）表头在第二行
  function getHeaderRowIndex(sheetName) { return sheetName.includes('质检回查') ? 1 : 0; }

  function getSheetHeaderRow(XLSX, worksheet, sheetName) {
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

  function getColValue(XLSX, worksheet, r, c) { const cell = worksheet[XLSX.utils.encode_cell({ r: r, c: c })]; return cell ? cell.v : ''; }

  // ==================== 解析 ====================
  function cleanWorkbook(XLSX, wb) {
    wb.SheetNames.forEach(sn => {
      const ws = wb.Sheets[sn];
      if (!ws || !ws['!ref']) return;
      // 注意：不能用 Math.min(...dr)/Math.max(...dr) —— 大文件行数/列数众多时 spread 会爆调用栈
      let minR = Infinity, maxR = -Infinity, minC = Infinity, maxC = -Infinity;
      let cellCount = 0;
      for (let c in ws) {
        if (c[0] === '!') continue;
        const a = XLSX.utils.decode_cell(c);
        if (a.r < minR) minR = a.r; if (a.r > maxR) maxR = a.r;
        if (a.c < minC) minC = a.c; if (a.c > maxC) maxC = a.c;
        cellCount++;
      }
      if (cellCount === 0) return;
      ws['!ref'] = XLSX.utils.encode_range({ s: { r: minR, c: minC }, e: { r: maxR, c: maxC } });
      for (let c in ws) { if (c[0] === '!') continue; ws[c] = { v: ws[c].v, t: ws[c].t }; }
    });
  }

  // 读取并清理工作簿，返回 { workbook, sheetNames, totalRows, sheetHeaders }
  function parseWorkbook(XLSX, data, opts) {
    opts = opts || {};
    const wb = opts.csvText !== undefined
      ? XLSX.read(opts.csvText, { type: 'string', codepage: 936, raw: true, cellStyles: false, sheetStubs: false })
      : XLSX.read(data, { type: 'array', cellStyles: false, sheetStubs: false });
    cleanWorkbook(XLSX, wb);
    return describeWorkbook(XLSX, wb);
  }

  function describeWorkbook(XLSX, wb) {
    const sheetNames = wb.SheetNames || [];
    let totalRows = 0;
    const sheetHeaders = {};
    const sheetRows = {};
    sheetNames.forEach(sn => {
      const ws = wb.Sheets[sn];
      if (ws && ws['!ref']) {
        const r = XLSX.utils.decode_range(ws['!ref']);
        totalRows += (r.e.r - r.s.r + 1);
      }
      sheetHeaders[sn] = getSheetHeaderRow(XLSX, ws, sn);
      // 数据行数 = 表头行之后的实际行数
      sheetRows[sn] = countDataRows(XLSX, ws, sn);
    });
    return { workbook: wb, sheetNames, totalRows, sheetHeaders, sheetRows };
  }

  // 计算某 sheet 的数据行数（去掉表头行）
  function countDataRows(XLSX, ws, sheetName) {
    if (!ws || !ws['!ref']) return 0;
    const r = XLSX.utils.decode_range(ws['!ref']);
    const hr = getHeaderRowIndex(sheetName);
    return Math.max(0, r.e.r - (hr + 1) + 1);
  }

  // ==================== 按行合并（可在 Worker 中执行） ====================
  // sources: [{ workbook, sheet, headers }]
  // mapping: [{ targetCol, sourceCol: [] }]
  // opts: { staffEnable, staffList }
  // onProgress: (pct) => void  可选进度回调（0-100）
  // 返回合并后的行数组（普通对象）
  function buildMergedRows(XLSX, sources, mapping, opts, onProgress) {
    opts = opts || {};
    const staffEnable = !!opts.staffEnable;
    const staffList = opts.staffList || [];
    const mergedData = [];
    let done = 0;
    sources.forEach(s => {
      const ws = s.workbook.Sheets[s.sheet];
      if (ws && ws['!ref']) {
        const range = XLSX.utils.decode_range(ws['!ref']);
        const hr = getHeaderRowIndex(s.sheet), start = hr + 1, last = range.e.r;
        if (last >= start) {
          const headers = s.headers && s.headers.length ? s.headers : getSheetHeaderRow(XLSX, ws, s.sheet);
          const idx = mapping.map(m => m.sourceCol.map(sc => headers.indexOf(sc)).filter(x => x >= 0));
          const span = last - start + 1;
          for (let r = start; r <= last; r++) {
            const row = { '来源sheet名': s.sheet };
            let ok = true, empty = true;
            mapping.forEach((m, mi) => {
              let v = '';
              for (const ci of idx[mi]) { const cv = getColValue(XLSX, ws, r, ci); if (cv !== undefined && cv !== null && cv !== '') { v = cv; empty = false; break; } }
              if ((/日期|时间/.test(m.targetCol)) && v === '-1') v = '永久';
              else if (/日期|时间/.test(m.targetCol)) v = formatDate(v, s.sheet);
              else if (/人员|姓名/.test(m.targetCol)) { v = extractChineseName(v); if (staffEnable && !staffList.includes(v)) ok = false; }
              else if (m.targetCol === '疑似类型') { const i = String(v).indexOf('违规备注：'); if (i !== -1) v = String(v).substring(i + 5).trim(); }
              row[m.targetCol] = v;
            });
            if (!empty && ok) mergedData.push(row);
          }
        }
      }
      done++;
      if (onProgress) onProgress(Math.round(done / sources.length * 100));
    });
    return mergedData;
  }

  // 备注列拆分/过滤（业务规则集中在此，便于后续改为可配置）
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

  // ==================== 格式化 ====================
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

  root.MergeCore = {
    LARGE_FILE_MB,
    LARGE_FILE_ROWS,
    STORAGE_KEYS,
    SPECIAL_COLUMNS,
    PRESET_MAPPINGS,
    FIELD_PRESET_ORDER,
    MAP_PRESET_ORDER,
    escapeHtml,
    getHeaderRowIndex,
    getSheetHeaderRow,
    getColValue,
    cleanWorkbook,
    parseWorkbook,
    describeWorkbook,
    countDataRows,
    buildMergedRows,
    processRemarkColumn,
    formatDate,
    extractChineseName
  };
})(typeof self !== 'undefined' ? self : (typeof window !== 'undefined' ? window : this));
