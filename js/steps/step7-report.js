/*
 * 第 7 步：總表（Coating：一般鍍膜）
 *   在活頁簿裡新增「光鍍人力需求」工作表，版面照使用者給的範本（B 到 K 行、第 2 到 49 列，
 *   設定在 js/excel-io.js 的 DM_REPORT、DM_ROWS），並填入：
 *     工作天數 - 常日班、做四休二   各月輸入（不會記住）；做四休二自動帶入該月的日曆天數（可以改）
 *     工作時間           各月輸入工作時數、休息時數（小時），填 =(工作時數-休息時數)*60（分鐘）
 *     機台每日可動工時   各月輸入小時，填 =小時*60（分鐘）
 *     工作時數、休息時數、機台每日可動工時、班別數、出勤率、作業效率、兩站的加班天數、共用操機人員   各月輸入，有「同步首月」按鈕
 *     共用操機人員 清洗、光譜檢測、大清機   各月輸入
 *     需求操機人員(GMO+Plastic+共用)   =GMO 站需求操機人員+Plastic 站需求操機人員+清洗+光譜檢測+大清機
 *     操機人力 Loading Rate            =需求操機人員(GMO+Plastic+共用)/現有操機人員
 *     目檢站 加班天數 - GMO、- Plastic、備料   各月輸入
 *     目檢站 需求人力 - GMO、- Plastic
 *       =ROUND(需求工時/((工作天數 做四休二+加班天數)*工作時間*班別數)*6/4/出勤率/作業效率,1)
 *     需求目檢人員       =ROUND(需求人力 - GMO+需求人力 - Plastic+備料,1)
 *     目檢人力 Loading Rate =需求目檢人員/現有目檢人員
 *     需求人力（On Line）(操機+共用+目檢) =需求操機人員(GMO+Plastic+共用)+需求目檢人員
 *     總需求人力 =需求人力（On Line）+需求人力（Off Line）；總現有人力 =現有人力（On Line）+現有人力（Off Line）
 *     人力需求差異 =總現有人力-總需求人力；總人力 Loading Rate =總需求人力/總現有人力
 *     現有人力（On Line）、（Off Line）  第 6 步確認的 on line／off line 列 × 一般鍍膜、＿＿鍍膜兩行，全部加總
 *                                        例如 =SUM('人力'!D5:D10,'人力'!E5:E10)
 *     需求人力（Off Line）               =現有人力（Off Line）
 *     MPS GMO、Plastic   公式：參照 MPS 工作表該段（GMO = 第 1 步的 Glass 段）第 2 步的 Total 列，例如 ='MPS'!E10；
 *                        沒有 Total 列就直接加總該段（當月也不加 WIP）
 *     GMO 站、Plastic 站
 *       機台需求工時   公式：參照 MPS 工作表該段「機台需求工時」第 5 步的 Total 列，例如 ='MPS'!P10
 *       人機比         輸入機台數、人員數（六個月共用），填 =機台數/人員數；Notes 寫「1 人顧 4 機」
 *       開機數         =MIN(ROUND(機台需求工時/((工作天數 做四休二+該站加班天數)*機台每日可動工時),1),最大開機數)
 *                      最大開機數在畫面上輸入（每一站一個）；沒填就不設上限
 *       加班天數       各月輸入
 *       需求操機人員   =ROUND(開機數/人機比*班別數*6/4/出勤率/作業效率,1)
 *     目檢站 需求工時 - GMO、- Plastic   公式：參照 MPS 工作表該段「目檢需求工時」第 5 步的 Total 列
 *   其他格（開機數、加班天數、需求操機人員、共用操機人員、Loading Rate、需求／現有人力…）之後再一項一項補上。
 *
 *   月份表頭用公式參照 MPS 工作表「機台需求工時」那一塊的月份（第 5 步輸入的），例如 ='MPS'!P3。
 *
 *   來源有兩種：
 *     a. 剛才第 6 步做好的結果（不用先匯出，直接沿用）
 *     b. 上傳之前第 6 步匯出的檔案
 *
 * 做完之後把結果放在 IEApp.state.withReport（這是最後一步，之後要接新的步驟時可以直接拿來用）。
 */
(function (global) {
  'use strict';

  var IEApp = global.IEApp;
  var XIO = global.ExcelIO;
  var ui = IEApp.ui;
  var h = IEApp.h;

  var STEP1_KEY = 'step1';
  var R = XIO.DM_REPORT;
  var SHEET_NAME = R.sheet;
  var MONTH_TITLE = XIO.DEMAND_TITLES.machine;      // 月份從「機台需求工時」那一塊抓
  var SUFFIX = '_總表';
  var EXTRA = ['P003', '人力', SHEET_NAME, '總表', 'RMA人力需求'];

  var S = {
    upload: null,         // 上傳的第 6 步檔案 { name, buffer, sheets }
    preferUpload: false,
    work: null            // { src, loading, error, wb, main, refs }
  };
  var els = {};

  /* ---------- 使用者輸入的數值 ---------- */
  // 每個月一格的列。percent：輸入的是 %（95 代表 95%）；copy：後面有「同步首月」按鈕；remember：記在瀏覽器裡
  // calendar：自動帶入該月的日曆天數（還沒改過的格子），後面有「帶入日曆天數」按鈕
  var GRID = [
    { key: 'workDaysDay', label: '工作天數 - 常日班', unit: '天', remember: false },
    { key: 'workDaysShift', label: '工作天數 - 做四休二', unit: '天', remember: false, calendar: true },
    { key: 'workHours', label: '工作時數', unit: '小時', remember: true, copy: true },
    { key: 'restHours', label: '休息時數', unit: '小時', remember: true, copy: true },
    { key: 'machineHours', label: '機台每日可動工時', unit: '小時', remember: true, copy: true },
    { key: 'shifts', label: '班別數', unit: '班/天', remember: true, copy: true },
    { key: 'attendance', label: '出勤率', unit: '%', remember: true, copy: true, percent: true },
    { key: 'efficiency', label: '作業效率', unit: '%', remember: true, copy: true, percent: true },
    { key: 'gmoOtDays', label: 'GMO 站 加班天數', unit: '天', remember: true, copy: true },
    { key: 'plasticOtDays', label: 'Plastic 站 加班天數', unit: '天', remember: true, copy: true },
    { key: 'sharedWash', label: '共用操機人員 - 清洗', unit: '人/天', remember: true, copy: true },
    { key: 'sharedSpectrum', label: '共用操機人員 - 光譜檢測', unit: '人/天', remember: true, copy: true },
    { key: 'sharedClean', label: '共用操機人員 - 大清機', unit: '人/天', remember: true, copy: true },
    { key: 'laborOtGmo', label: '目檢站 加班天數 - GMO', unit: '天', remember: true, copy: true },
    { key: 'laborOtPlastic', label: '目檢站 加班天數 - Plastic', unit: '天', remember: true, copy: true },
    { key: 'laborPplPrep', label: '目檢站 需求人力 - 備料', unit: '人/天', remember: true, copy: true }
  ];
  // 人機比：每一站輸入一組機台數、人員數（六個月共用）
  var STATIONS = [
    { key: 'glass', label: 'GMO 站', name: 'GMO', machines: 'gmoMachines', people: 'gmoPeople', max: 'gmoMaxMachines', id: 'gmo' },
    { key: 'plastic', label: 'Plastic 站', name: 'Plastic', machines: 'plasticMachines', people: 'plasticPeople', max: 'plasticMaxMachines', id: 'pla' }
  ];
  var INPUT_KEY = 'step7';             // 共用設定的組名（IEApp.settings）
  var INP = loadInputs();       // { workDaysDay: ['', …], workHours: [...], …, gmoMachines: '' }（輸入框的文字）

  function blankMonths() { var a = []; for (var i = 0; i < R.months; i++) a.push(''); return a; }
  function loadInputs() {
    var s = {};
    try { s = IEApp.settings.read(INPUT_KEY) || {}; } catch (e) { s = {}; }
    var out = {};
    STATIONS.forEach(function (st) {
      out[st.machines] = s[st.machines] || ''; out[st.people] = s[st.people] || ''; out[st.max] = s[st.max] || '';
    });
    GRID.forEach(function (g) {
      var saved = g.remember && Array.isArray(s[g.key]) ? s[g.key] : [];
      // 版本 5 以前工作時數、休息時數是六個月共用一個值：六個月都填同一個
      if (g.remember && !Array.isArray(s[g.key]) && s[g.key] !== undefined && s[g.key] !== null && s[g.key] !== '') {
        saved = blankMonths().map(function () { return s[g.key]; });
      }
      out[g.key] = blankMonths().map(function (v, i) { return saved[i] === undefined ? '' : String(saved[i]); });
    });
    return out;
  }
  function saveInputs() {
    var s = {};
    STATIONS.forEach(function (st) { s[st.machines] = INP[st.machines]; s[st.people] = INP[st.people]; s[st.max] = INP[st.max]; });
    GRID.forEach(function (g) { if (g.remember) s[g.key] = INP[g.key]; });
    try { IEApp.settings.write(INPUT_KEY, s); } catch (e) { /* 存不了就算了 */ }
  }

  // 輸入框的文字 → 數字。空白回傳 null；看不懂回傳 NaN。percent 時 95 或 95% 都回傳 0.95
  function toValue(text, percent) {
    var t = XIO.normalizeText(text || '').replace(/,/g, '');
    if (percent) t = t.replace(/[%％]\s*$/, '');
    t = t.trim();
    if (t === '') return null;
    var n = Number(t);
    if (isNaN(n) || n < 0) return NaN;
    return percent ? n / 100 : n;
  }

  // 讀出所有輸入：{ values: 給 addDmReportSheet 的 inputs, bad: [看不懂的], blank: [沒填的] }
  function readInputs() {
    var values = {}, bad = [], blank = [];
    var names = monthNames();
    GRID.forEach(function (g) {
      values[g.key] = INP[g.key].map(function (text, i) {
        var v = toValue(text, g.percent);
        if (v === null) blank.push(g.label + '（' + names[i] + '）');
        else if (isNaN(v)) bad.push(g.label + '（' + names[i] + '）「' + text + '」');
        return v === null || isNaN(v) ? null : v;
      });
    });
    STATIONS.forEach(function (st) {
      [[st.machines, st.label + '的機台數'], [st.people, st.label + '的人員數'], [st.max, st.name + ' 機台最大開機數']].forEach(function (p) {
        var v = toValue(INP[p[0]], false);
        if (v === null) blank.push(p[1]);
        else if (isNaN(v)) bad.push(p[1] + '「' + INP[p[0]] + '」');
        values[p[0]] = v === null || isNaN(v) ? null : v;
      });
      if (values[st.people] === 0) bad.push(st.label + '的人員數不能是 0');
    });
    values.workHours.forEach(function (w, i) {
      var r = values.restHours[i];
      if (w !== null && r !== null && r > w) bad.push('休息時數比工作時數還多（' + names[i] + '）');
    });
    return { values: values, bad: bad, blank: blank };
  }

  function monthNames() {
    var refs = S.work && S.work.refs;
    var out = [];
    for (var i = 0; i < R.months; i++) out.push(refs && refs[i] && refs[i].text ? refs[i].text : '第 ' + (i + 1) + ' 個月');
    return out;
  }

  // 輸入的表格：一列一個項目、一行一個月份；建立一次，之後只更新月份名稱
  function buildInputGrid() {
    els.monthHeads = [];
    els.gridInputs = {};
    var head = h('tr', null, h('th', { scope: 'col', text: '項目' }));
    for (var i = 0; i < R.months; i++) {
      var th = h('th', { scope: 'col', class: 'is-month', text: '第 ' + (i + 1) + ' 個月' });
      els.monthHeads.push(th);
      head.appendChild(th);
    }
    head.appendChild(h('th', { scope: 'col' }));
    var body = h('tbody');
    GRID.forEach(function (g) {
      els.gridInputs[g.key] = [];
      var tr = h('tr', null, h('th', { scope: 'row' },
        h('span', { class: 'grid-label', text: g.label }), h('span', { class: 'grid-unit', text: g.unit })));
      for (var m = 0; m < R.months; m++) {
        var input = h('input', {
          type: 'text', class: 'in in-grid', inputmode: 'decimal', autocomplete: 'off',
          value: INP[g.key][m], 'aria-label': g.label + ' 第 ' + (m + 1) + ' 個月'
        });
        (function (key, idx, el) {
          el.addEventListener('input', function () {
            INP[key][idx] = el.value;
            if (AUTO[key]) AUTO[key][idx] = false;      // 自己改過的格子，之後不再自動帶入
            onType();
          });
        })(g.key, m, input);
        els.gridInputs[g.key].push(input);
        tr.appendChild(h('td', null, input));
      }
      tr.appendChild(h('td', null, g.copy ? h('button', {
        type: 'button', class: 'btn-secondary btn-small', text: '同步首月',
        title: '把第一個月的值填到後面每個月',
        onclick: function () { copyFirst(g.key); }
      }) : g.calendar ? h('button', {
        type: 'button', class: 'btn-secondary btn-small', text: '帶入日曆天數',
        title: '六個月都重新填入該月的日曆天數',
        onclick: function () { fillCalendar(g.key, true); }
      }) : (g.note ? h('span', { class: 'grid-note', text: g.note }) : null)));
      body.appendChild(tr);
    });
    return h('div', { class: 'sheet-scroll grid-scroll' }, h('table', { class: 'grid-inputs' }, h('thead', null, head), body));
  }

  /* ---------- 工作天數 做四休二：自動帶入日曆天數 ---------- */
  // AUTO[key][i] = true：這一格是程式自動帶入的（換了月份會跟著重算）；使用者改過就變 false
  var AUTO = {};
  GRID.forEach(function (g) {
    if (!g.calendar) return;
    AUTO[g.key] = INP[g.key].map(function (v) { return v === ''; });
  });

  var MONTH_ABBR = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
  // 月份表頭的文字 → { month: 0–11, year: 西元年或 null }；看不懂回傳 null
  function parseMonth(text) {
    var t = XIO.normalizeText(text || '').toLowerCase();
    if (!t) return null;
    var m = -1, y = null, k;
    var iso = /(\d{4})\s*[\/\-.年]\s*(\d{1,2})/.exec(t);
    if (iso) { y = +iso[1]; m = +iso[2] - 1; }
    if (m < 0) {
      for (k = 0; k < 12; k++) if (t.indexOf(MONTH_ABBR[k]) >= 0) { m = k; break; }
    }
    if (m < 0) { var zh = /(\d{1,2})\s*月/.exec(t); if (zh) m = +zh[1] - 1; }
    if (m < 0 || m > 11) return null;
    if (y === null) {
      var y4 = /(\d{4})/.exec(t);
      var y2 = /[\-\/'\s]\s*(\d{2})\s*$/.exec(t);
      if (y4) y = +y4[1]; else if (y2) y = 2000 + +y2[1];
    }
    return { month: m, year: y };
  }
  // 六個月各自的日曆天數；月份看不懂的回傳 null。
  // 沒寫年份時：第一個月取離今天最近的那一年，後面的月份遇到跨年就加一年
  function calendarDays(names) {
    var today = new Date();
    var prev = null;
    return names.map(function (name) {
      var p = parseMonth(name);
      if (!p) { prev = null; return null; }
      var y = p.year;
      if (y === null) {
        if (prev) y = p.month <= prev.month ? prev.year + 1 : prev.year;
        else {
          y = today.getFullYear();
          var diff = p.month - today.getMonth();
          if (diff > 6) y--; else if (diff < -6) y++;
        }
      }
      prev = { month: p.month, year: y };
      return new Date(y, p.month + 1, 0).getDate();
    });
  }
  // 填入日曆天數：force = 六格都重填（按鈕）；否則只填還沒改過的格子
  function fillCalendar(key, force) {
    if (!S.work || !S.work.refs) {
      if (force) els.exportNote.textContent = '還沒有月份（要先有第 5 步的機台需求工時月份），沒辦法帶入日曆天數。';
      return false;
    }
    var days = calendarDays(monthNames());
    var changed = false;
    for (var i = 0; i < R.months; i++) {
      if (days[i] === null || (!force && !AUTO[key][i])) continue;
      var v = String(days[i]);
      AUTO[key][i] = true;
      if (INP[key][i] === v) continue;
      INP[key][i] = v;
      if (els.gridInputs && els.gridInputs[key]) els.gridInputs[key][i].value = v;
      changed = true;
    }
    if (force) onType();
    return changed;
  }

  function copyFirst(key) {
    var first = INP[key][0];
    for (var m = 1; m < R.months; m++) {
      INP[key][m] = first;
      els.gridInputs[key][m].value = first;
    }
    onType();
  }

  var typeTimer = null;
  function onType() {
    saveInputs();
    els.exportNote.textContent = '';
    clearTimeout(typeTimer);
    typeTimer = setTimeout(refresh, 250);
  }

  function sameName(a, b) { return String(a).toLowerCase() === String(b).toLowerCase(); }
  function isExtra(name) { return EXTRA.some(function (n) { return sameName(n, name); }); }

  function activeSource() {
    var prev = IEApp.state.withManpower;
    if (S.upload && (S.preferUpload || !prev)) return { kind: 'file', key: S.upload, name: S.upload.name };
    if (prev) return { kind: 'prev', key: prev, name: prev.outName, prev: prev };
    return null;
  }

  function buildFrom(src) {
    return src.kind === 'prev' ? Promise.resolve(src.prev.build()) : XIO.load(src.key.buffer.slice(0));
  }

  function pickMain(wb, src) {
    var others = wb.worksheets.filter(function (w) { return !isExtra(w.name); });
    var wanted = src.kind === 'prev' ? src.prev.mainSheet : null;
    return others.filter(function (w) { return w.name === wanted; })[0] || others[0] || null;
  }

  // MPS 的資料從第幾列開始：沿用前面步驟的；上傳的檔案用第 1 步記住的設定（預設第 5 列）
  function dataStartRow(src) {
    if (src.kind === 'prev' && src.prev.dataStartRow) return src.prev.dataStartRow;
    var n = 0;
    try {
      var s1 = IEApp.settings.read(STEP1_KEY) || {};
      n = parseInt(XIO.normalizeText(s1.startRow === undefined ? '' : s1.startRow), 10);
    } catch (e) { n = 0; }
    return n >= 1 ? n : 5;
  }

  // 各月產量在哪幾行：和第 2、5 步共用的設定（沒有設定就用預設的第 5 到 10 行）
  function mpsColumns() {
    var shared = ui.columns.get();
    var months = XIO.parseColumns(shared.monthCols || '').cols;
    if (!months.length) months = [5, 6, 7, 8, 9, 10];
    return { monthCols: months };
  }

  // 總表 MPS 那兩列的公式（GMO = 第 1 步的 Glass 段、Plastic）
  function mpsFormulas(main, src) {
    var cols = mpsColumns();
    return XIO.dmMpsFormulas(main, { headerEnd: dataStartRow(src) - 1, monthCols: cols.monthCols });
  }

  // GMO 站、Plastic 站 機台需求工時的公式（參照 MPS 工作表機台需求工時那一塊各段的 Total 列）
  function machineTotals(main, src) {
    return XIO.dmBlockTotalFormulas(main, { headerEnd: dataStartRow(src) - 1, title: MONTH_TITLE, months: R.months });
  }

  // 目檢站 需求工時的公式（參照 MPS 工作表目檢需求工時那一塊各段的 Total 列）
  var LABOR_TITLE = XIO.DEMAND_TITLES.labor;
  function laborTotals(main, src) {
    return XIO.dmBlockTotalFormulas(main, { headerEnd: dataStartRow(src) - 1, title: LABOR_TITLE, months: R.months });
  }

  // 月份表頭要參照的儲存格：MPS 工作表「機台需求工時」那一塊的月份（大表頭下一列）
  function monthRefs(main, src) {
    var headerEnd = dataStartRow(src) - 1;
    var block = XIO.findHeaderColumn(main, headerEnd, MONTH_TITLE);
    if (!block) return null;
    var titleRow = 0;
    for (var r = 1; r <= headerEnd && !titleRow; r++) {
      var cell = main.findCell(r, block);
      if (cell && XIO.lookupKey(cell.value) === XIO.lookupKey(MONTH_TITLE)) titleRow = r;
    }
    var headRow = titleRow ? titleRow + 1 : (headerEnd >= 3 ? 3 : headerEnd);
    if (headRow < 1 || headRow > headerEnd) return null;
    var refs = [];
    for (var i = 0; i < R.months; i++) {
      var c = block + i;
      var hc = main.findCell(headRow, c);
      refs.push({ sheet: main.name, col: c, row: headRow, text: hc ? XIO.displayText(hc) : '' });
    }
    return refs;
  }

  /* ---------- 畫面 ---------- */
  function render(panel) {
    els = {};

    els.srcInfo = h('p', { class: 'source-line', 'aria-live': 'polite' });
    els.drop = ui.fileDrop({
      id: 'f9-base',
      title: '上傳第 6 步匯出的 Excel',
      sub: '或把檔案拖進這裡。支援 .xlsx',
      onFile: openUpload
    });
    els.switchBtn = h('button', { type: 'button', class: 'btn-link', hidden: true, onclick: switchSource });
    els.backBtn = h('button', {
      type: 'button', class: 'btn-link', hidden: true, text: '回到第 6 步匯入人力報表',
      onclick: function () { IEApp.go('manpower'); }
    });
    els.srcMsg = h('p', { class: 'msg is-error', role: 'alert', hidden: true });

    els.summary = h('p', { class: 'summary', 'aria-live': 'polite' });
    els.notes = h('ul', { class: 'notes' });
    els.layout = h('div', { class: 'sheet-area' });
    els.exportBtn = h('button', { type: 'button', class: 'btn-primary', disabled: true, text: '匯出 Excel', onclick: onExport });
    els.exportNote = h('span', { class: 'export-note', 'aria-live': 'polite' });

    // 工作時間：工作時數、休息時數（六個月共用）
    // 人機比：每一站一列（機台數、人員數）
    els.ratioText = {};
    var ratioRows = STATIONS.map(function (st) {
      var mIn = h('input', { id: 'f9-' + st.id + '-m', type: 'text', class: 'in in-tiny', inputmode: 'decimal', autocomplete: 'off', value: INP[st.machines], placeholder: '例如 4' });
      var pIn = h('input', { id: 'f9-' + st.id + '-p', type: 'text', class: 'in in-tiny', inputmode: 'decimal', autocomplete: 'off', value: INP[st.people], placeholder: '例如 1' });
      mIn.addEventListener('input', function () { INP[st.machines] = mIn.value; onType(); });
      pIn.addEventListener('input', function () { INP[st.people] = pIn.value; onType(); });
      var xIn = h('input', { id: 'f9-' + st.id + '-x', type: 'text', class: 'in in-tiny', inputmode: 'decimal', autocomplete: 'off', value: INP[st.max], placeholder: '例如 8' });
      xIn.addEventListener('input', function () { INP[st.max] = xIn.value; onType(); });
      els.ratioText[st.key] = h('span', { class: 'hint-inline' });
      return h('div', { class: 'station-block' },
        h('div', { class: 'field-row' },
          h('span', { class: 'field-label label-fixed', text: st.label }),
          h('label', { for: 'f9-' + st.id + '-m', text: '人機比：機台數' }), mIn, h('span', { text: '台' }),
          h('label', { for: 'f9-' + st.id + '-p', text: '人員數' }), pIn, h('span', { text: '人' }),
          els.ratioText[st.key]),
        h('div', { class: 'field-row' },
          h('span', { class: 'field-label label-fixed', 'aria-hidden': 'true' }),
          h('label', { for: 'f9-' + st.id + '-x', text: st.name + ' 機台最大開機數' }), xIn, h('span', { text: '台' })));
    });

    panel.appendChild(h('div', { class: 'step-head' },
      h('h1', { text: '總表' }),
      h('p', { text: '在同一個檔案裡新增「光鍍人力需求」工作表（一般鍍膜）：MPS、GMO 站、Plastic 站、共用操機人員、目檢站、需求人力、現有人力、人力需求差異、Loading Rate。月份表頭會參照 MPS 工作表「機台需求工時」的月份。這一步先填工作天數、工作時間、機台每日可動工時、班別數、出勤率、作業效率、MPS 的兩列、GMO 站和 Plastic 站、共用操機人員、需求操機人員、操機人力 Loading Rate、目檢站、需求目檢人員、目檢人力 Loading Rate；其他格之後再補上。GMO 是第 1 步的 Glass 那一段。' })));

    panel.appendChild(h('ol', { class: 'tasks' },
      h('li', { class: 'task' },
        h('h2', { text: '第 6 步的 Excel' }),
        els.srcInfo, els.drop.el,
        h('div', { class: 'link-row' }, els.switchBtn, els.backBtn),
        els.srcMsg),
      h('li', { class: 'task' },
        h('h2', { text: '各月的數值' }),
        buildInputGrid(),
        h('p', { class: 'hint', text: '工作天數 - 做四休二會自動帶入該月的日曆天數（例如 Oct 是 31），不對可以直接改；按「帶入日曆天數」會六個月重新帶入。工作時數、休息時數輸入小時，總表的工作時間會填公式 =(工作時數-休息時數)*60（分鐘）。出勤率、作業效率輸入 95 或 95% 都代表 95%。每月都一樣的列只要填第一個月，再按「同步首月」。機台每日可動工時輸入小時，總表會換算成分鐘（例如 =22*60）。共用操機人員（清洗、光譜檢測、大清機）的人數會加進總表的「需求操機人員(GMO+Plastic+共用)」；目檢站的備料人數會加進「需求目檢人員」。工作天數每個月不同，不會記住；其他會記住上次輸入的值。沒填的格子在總表留空白。' })),
      h('li', { class: 'task' },
        h('h2', { text: 'GMO 站、Plastic 站' }),
        ratioRows,
        h('p', { class: 'hint', text: '每一站一組，六個月共用，會記住上次輸入的值。總表的人機比填公式 =機台數/人員數，Notes 寫「幾人顧幾機」。' +
          '開機數 =ROUND(機台需求工時/((工作天數 做四休二+該站加班天數)*機台每日可動工時),1)，算出來超過最大開機數就顯示最大開機數（沒填就不設上限）。' +
          '需求操機人員 =ROUND(開機數/人機比*班別數*6/4/出勤率/作業效率,1)。' })),
      h('li', { class: 'task' },
        h('h2', { text: '匯出' }),
        els.summary, els.notes, els.layout,
        h('div', { class: 'actions' }, els.exportBtn, els.exportNote))));

    refresh();
  }

  /* ---------- 來源 ---------- */
  function openUpload(file) {
    els.srcMsg.hidden = true;
    els.drop.setBusy(file.name);
    ui.readWorkbook(file)
      .then(function (res) {
        var names = res.workbook.worksheets.map(function (w) { return w.name; });
        if (!names.some(function (n) { return !isExtra(n); })) {
          throw new Error('這個檔案沒有 MPS 工作表。請上傳第 6 步匯出的檔案。');
        }
        S.upload = { name: file.name, buffer: res.buffer, sheets: names };
        S.preferUpload = true;
        els.exportNote.textContent = '';
        refresh();
      })
      .catch(function (err) {
        els.srcMsg.textContent = err && err.message ? err.message : String(err);
        els.srcMsg.hidden = false;
        refresh();
      });
  }

  function switchSource() {
    S.preferUpload = !S.preferUpload;
    els.exportNote.textContent = '';
    refresh();
  }

  function ensureWork(src) {
    if (S.work && S.work.src === src.key) return;
    var work = { src: src.key, loading: true };
    S.work = work;
    buildFrom(src).then(function (wb) {
      if (S.work !== work) return;
      var main = pickMain(wb, src);
      if (!main) throw new Error('找不到 MPS 工作表。');
      work.wb = wb;
      work.main = main;
      work.refs = monthRefs(main, src);
      work.mps = mpsFormulas(main, src);
      work.machineTotals = machineTotals(main, src);
      work.laborTotals = laborTotals(main, src);
      work.hasManpower = wb.worksheets.some(function (w) { return sameName(w.name, MP_SHEET); });
      work.loading = false;
      refresh();
    }).catch(function (err) {
      if (S.work !== work) return;
      work.loading = false;
      work.error = err && err.message ? err.message : String(err);
      refresh();
    });
  }

  /* ---------- 更新畫面 ---------- */
  function refresh() {
    var src = activeSource();
    var prev = IEApp.state.withManpower;
    IEApp.state.withReport = null;
    els.exportBtn.disabled = true;
    els.notes.textContent = '';
    els.layout.textContent = '';
    refreshInputLabels();

    if (!src) {
      els.srcInfo.textContent = '還沒有第 6 步的結果。回到第 6 步，或上傳之前匯出的檔案。';
      els.drop.setIdle();
    } else if (src.kind === 'prev') {
      els.srcInfo.textContent = '沿用第 6 步剛做好的結果：' + prev.outName + '。不用先匯出；要改用之前匯出的檔案，再從下面上傳。';
      if (S.upload) els.drop.setFile(S.upload.name, '已上傳，目前沒有使用。點這裡換一個檔案');
      else els.drop.setIdle();
    } else {
      els.srcInfo.textContent = '使用上傳的檔案。';
      els.drop.setFile(S.upload.name);
    }
    els.backBtn.hidden = !!src;
    els.switchBtn.hidden = !(S.upload && prev);
    if (src) els.switchBtn.textContent = src.kind === 'file' ? '改用第 6 步剛做好的結果' : '改用上傳的檔案';

    if (!src) { setSummary('先準備好第 6 步的 Excel。'); return; }
    ensureWork(src);
    var work = S.work;
    if (work.loading) { setSummary('準備資料中…'); return; }
    if (work.error) { setSummary(work.error, true); return; }

    var refs = work.refs;
    var first = XIO.numToCol(R.monthCol), last = XIO.numToCol(R.monthCol + R.months - 1);
    setSummary('會新增「' + SHEET_NAME + '」工作表（B 到 K 行、第 ' + R.titleRow + ' 到 ' + R.lastRow + ' 列）。' +
      (refs ? '月份表頭 ' + first + R.headRow + ' 到 ' + last + R.headRow + ' 參照 ' + work.main.name + ' 工作表「' + MONTH_TITLE + '」的月份：' +
        refs.map(function (x) { return x.text || '（空白）'; }).join('、') + '。' : ''), false, !refs);
    if (!refs) {
      els.notes.appendChild(h('li', { class: 'is-warn' },
        work.main.name + ' 工作表找不到「' + MONTH_TITLE + '」（第 5 步），月份表頭先留空白。'));
    }
    var mps = work.mps;
    var mpsText = [['glass', 'GMO'], ['plastic', 'Plastic']].map(function (p) {
      var list = mps && mps[p[0]];
      if (!list) return p[1] + '：找不到';
      return p[1] + '：' + (list[0].fromTotal ? '參照第 2 步的 Total 列' : '直接加總該段') + '，例如 =' + list[0].formula;
    });
    els.notes.appendChild(h('li', { class: mps && mps.glass && mps.plastic ? null : 'is-warn' },
      'MPS 的兩列（公式）：' + mpsText.join('；') + '。' +
      (mps && mps.glass && mps.plastic ? '' : '找不到的那一列先留空白（MPS 工作表裡要有第 1 步的「Glass（…）」「Plastic（…）」標題和料號；GMO 用的是 Glass 那一段）。')));

    var mt = work.machineTotals;
    STATIONS.forEach(function (st) {
      var list = mt && mt[st.key];
      if (list) {
        els.notes.appendChild(h('li', null, st.label + ' 機台需求工時（公式）：' +
          (list[0].fromTotal ? '參照第 5 步的 Total 列' : '直接加總該段') + '，例如 =' + list[0].formula + '。'));
      } else {
        els.notes.appendChild(h('li', { class: 'is-warn' },
          'MPS 工作表找不到 ' + (st.key === 'glass' ? 'Glass' : 'Plastic') + ' 段的「' + MONTH_TITLE + '」（第 5 步），' + st.label + '的機台需求工時先留空白。'));
      }
    });

    var mpos = manpowerPos(src);
    if (!work.hasManpower) {
      els.notes.appendChild(h('li', { class: 'is-warn' }, '檔案裡沒有「' + MP_SHEET + '」工作表（第 6 步），現有人力（On Line、Off Line）和需求人力（Off Line）先留空白。'));
    } else if (!mpos) {
      els.notes.appendChild(h('li', { class: 'is-warn' }, '找不到第 6 步確認的位置（on line、off line 的列，一般鍍膜、＿＿鍍膜的行），現有人力先留空白。請回到第 6 步確認一次。'));
    } else {
      var mc = [mpos.generalCol, mpos.otherCol];
      els.notes.appendChild(h('li', null, '現有人力（公式）：On Line =SUM(' + MP_SHEET + ' ' + rangeText(mpos.onlineRows, mc) + ')、Off Line =SUM(' +
        MP_SHEET + ' ' + rangeText(mpos.offlineRows, mc) + ')；需求人力（Off Line）= 現有人力（Off Line）。'));
    }

    var lt = work.laborTotals;
    STATIONS.forEach(function (st) {
      var list = lt && lt[st.key];
      var name = '目檢站 需求工時 - ' + st.name;
      if (list) {
        els.notes.appendChild(h('li', null, name + '（公式）：' +
          (list[0].fromTotal ? '參照第 5 步的 Total 列' : '直接加總該段') + '，例如 =' + list[0].formula + '。'));
      } else {
        els.notes.appendChild(h('li', { class: 'is-warn' },
          'MPS 工作表找不到 ' + (st.key === 'glass' ? 'Glass' : 'Plastic') + ' 段的「' + LABOR_TITLE + '」（第 5 步），' + name + '先留空白。'));
      }
    });

    var read = readInputs();
    if (read.bad.length) {
      setSummary('看不懂：' + read.bad.join('、') + '。請輸入數字。', true);
      return;
    }
    if (read.blank.length) {
      els.notes.appendChild(h('li', { class: 'is-warn' },
        '還沒填：' + read.blank.slice(0, 12).join('、') + (read.blank.length > 12 ? ' …等 ' + read.blank.length + ' 格' : '') +
        '。沒填的格子在總表留空白，一樣可以匯出。'));
    }
    els.notes.appendChild(h('li', null, '目前填入工作天數、工作時間、機台每日可動工時、班別數、出勤率、作業效率、MPS 的兩列、GMO 站和 Plastic 站（機台需求工時、人機比、開機數、加班天數、需求操機人員）、共用操機人員、需求操機人員(GMO+Plastic+共用)、操機人力 Loading Rate、目檢站（需求工時、加班天數、需求人力、備料）、需求目檢人員、目檢人力 Loading Rate、需求人力（On Line、Off Line）、總需求人力、現有人力（On Line、Off Line）、總現有人力、人力需求差異、總人力 Loading Rate；現有操機人員、現有目檢人員先留空白，之後再補上（還沒填時，操機和目檢的 Loading Rate 會顯示 #DIV/0!）。'));

    var inputs = read.values;
    els.exportBtn.disabled = false;
    IEApp.state.withReport = {
      outName: outputName(src),
      mainSheet: work.main.name,
      dataStartRow: dataStartRow(src),
      reportSheet: SHEET_NAME,
      reportRows: rowKeys(),
      monthCol: R.monthCol,
      monthCount: R.months,
      inputs: inputs,
      build: function () { return buildWorkbook(src, refs, inputs); }
    };
  }

  // 月份名稱（從 MPS 抓到的）、工作時間換算後的分鐘，顯示在輸入區
  function refreshInputLabels() {
    GRID.forEach(function (g) { if (g.calendar) fillCalendar(g.key, false); });
    var names = monthNames();
    (els.monthHeads || []).forEach(function (th, i) { th.textContent = names[i]; });
    STATIONS.forEach(function (st) {
      var mm = toValue(INP[st.machines], false), pp = toValue(INP[st.people], false);
      els.ratioText[st.key].textContent = mm !== null && pp !== null && !isNaN(mm) && !isNaN(pp) && pp > 0
        ? '＝ 人機比 ' + Math.round(mm / pp * 100) / 100 + '（' + pp + ' 人顧 ' + mm + ' 機）' : '';
    });
  }

  function rowKeys() {
    var out = {};
    XIO.DM_ROWS.forEach(function (x) { out[x.key] = x.r; });
    return out;
  }

  function setSummary(text, isError, isWarn) {
    els.summary.textContent = text;
    els.summary.className = 'summary' + (isError ? ' is-error' : '') + (isWarn ? ' is-warn' : '');
  }

  /* ---------- 產生活頁簿 ---------- */
  /* ---------- 現有人力：第 6 步確認的位置 ---------- */
  var MP_SHEET = '人力';
  // 沿用第 6 步剛做好的結果時用它確認的位置；上傳第 6 步的檔案時用第 6 步記住的位置
  function manpowerPos(src) {
    var m = src.kind === 'prev' && src.prev.manpower ? src.prev.manpower : null;
    if (!m) {
      var s = {};
      try { s = IEApp.settings.read('step6') || {}; } catch (e) { s = {}; }
      m = { onlineRows: s.online, offlineRows: s.offline, generalCol: parseInt(s.generalCol, 10) || 0, otherCol: parseInt(s.otherCol, 10) || 0 };
    }
    function ok(p) { return Array.isArray(p) && p[0] > 0 && p[1] >= p[0]; }
    if (!ok(m.onlineRows) || !ok(m.offlineRows) || !(m.generalCol || m.otherCol)) return null;
    return { sheet: MP_SHEET, onlineRows: m.onlineRows, offlineRows: m.offlineRows, generalCol: m.generalCol, otherCol: m.otherCol };
  }
  function rangeText(rows, cols) {
    return cols.filter(Boolean).map(function (c) { return XIO.numToCol(c) + rows[0] + ':' + XIO.numToCol(c) + rows[1]; }).join('、');
  }

  function buildWorkbook(src, refs, inputs) {
    return buildFrom(src).then(function (wb) {
      var main = pickMain(wb, src);
      var mt = main ? machineTotals(main, src) : null;
      var lt = main ? laborTotals(main, src) : null;
      XIO.addDmReportSheet(wb, {
        monthRefs: refs, inputs: inputs, mps: main ? mpsFormulas(main, src) : null,
        gmoHours: mt ? mt.glass : null,
        gmoRatio: { machines: inputs.gmoMachines, people: inputs.gmoPeople },
        plasticHours: mt ? mt.plastic : null,
        plasticRatio: { machines: inputs.plasticMachines, people: inputs.plasticPeople },
        gmoMax: inputs.gmoMaxMachines,
        plasticMax: inputs.plasticMaxMachines,
        laborHours: lt || null,
        manpower: manpowerPos(src)
      });
      return wb;
    });
  }

  function outputName(src) {
    var stem = src.name.replace(/\.[^.]+$/, '');
    if (stem.slice(-SUFFIX.length) === SUFFIX) stem = stem.slice(0, -SUFFIX.length);
    return stem + SUFFIX + '.xlsx';
  }

  /* ---------- 匯出 ---------- */
  function onExport() {
    var ready = IEApp.state.withReport;
    if (!ready) return;
    els.exportBtn.disabled = true;
    els.exportNote.textContent = '產生檔案中…';
    ready.build()
      .then(function (wb) { return XIO.toBlob(wb); })
      .then(function (blob) {
        ui.download(blob, ready.outName);
        els.exportNote.textContent = '已匯出 ' + ready.outName + '，在瀏覽器的下載資料夾。';
      })
      .catch(function (err) {
        els.exportNote.textContent = '匯出失敗：' + (err && err.message ? err.message : err);
      })
      .then(function () {
        els.exportBtn.disabled = false;
      });
  }

  IEApp.registerStep({ id: 'report', render: render });
})(window);
