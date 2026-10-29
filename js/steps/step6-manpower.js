/*
 * 第 6 步：匯入人力報表（Coating）
 *   把人力報表（現有人力 - 鍍膜）放進前面做好的 Excel，成為同一個檔案裡名為「人力」的工作表（整張照抄）。
 *
 *   人力報表的樣子：B 行是大項（on line、off line，上下合併），C 行是類別（操機、目檢、文書…），
 *   表頭那一列有「一般鍍膜」「___鍍膜」，下面是各類別的人數。
 *
 *   匯入之後請使用者確認四件事，之後做總表時才知道要去哪幾格抓資料：
 *     1. on line 是第幾列到第幾列
 *     2. off line 是第幾列到第幾列
 *     3.「一般鍍膜」在哪一行
 *     4.「___鍍膜」在哪一行
 *     5.「類別」在哪一行，以及類別行裡每個項目（操機、目檢、文書…）各在第幾列
 *   程式會先自動帶入：寫 on line／off line 的格子（有合併就用合併的範圍；沒合併時 on line 到 off line 的前一列、
 *   off line 到最後一列），表頭寫「一般鍍膜」的那一行，以及同一列另一個寫「鍍膜」的那一行；
 *   類別行：表頭寫「類別」的那一行（沒有就用 on line 那一格右邊那一行），on line、off line 範圍內這一行有文字的格子就是各個項目。
 *   使用者核對、不對就改。都選好時會記住，下個月自動帶入不到時拿來當預設值（項目的列依名稱記）。
 *
 *   來源有兩種：
 *     a. 剛才第 5 步做好的結果（不用先匯出，直接沿用）
 *     b. 上傳之前第 5 步匯出的檔案
 *
 * 做完之後把結果放在 IEApp.state.withManpower，下一步直接拿來用：
 *   manpower  { onlineRows: [起, 迄], offlineRows: [起, 迄], generalCol, otherCol, catCol,
 *               categories: [{ name: '操機', row: 5, group: 'online' }, …] }（在「人力」工作表的列號、行號，A=1）
 */
(function (global) {
  'use strict';

  var IEApp = global.IEApp;
  var XIO = global.ExcelIO;
  var ui = IEApp.ui;
  var h = IEApp.h;

  var SHEET_NAME = '人力';
  var REF_SHEET = 'P003';
  var SUFFIX = '_人力';
  var STORE_KEY = 'step6';             // 共用設定的組名（IEApp.settings）
  var MIN_PREVIEW_ROWS = 15;
  var MAX_PREVIEW_ROWS = 200;

  // 要確認的四項：兩組起迄列、兩個行
  var GROUPS = [
    { key: 'online', label: 'on line', match: 'online' },
    { key: 'offline', label: 'off line', match: 'offline' }
  ];
  var COLS = [
    { key: 'generalCol', label: '「一般鍍膜」' },
    { key: 'otherCol', label: '「＿＿鍍膜」' }
  ];

  var S = {
    base: null,           // 上傳的第 5 步檔案 { name, buffer, sheets, hasManpower }
    preferUpload: false,
    mp: null,             // 人力報表 { name, workbook, ws, ext, labels: { 列號: 文字 } }
    pick: null            // 使用者確認的位置 { online: [起, 迄], offline: [起, 迄], generalCol, otherCol, headRow }（還沒選的是 0）
  };
  var settings = loadSettings();
  var els = {};

  function loadSettings() {
    var s = {};
    try { s = IEApp.settings.read(STORE_KEY) || {}; } catch (e) { s = {}; }
    function pair(v) { return Array.isArray(v) && v.length === 2 ? [parseInt(v[0], 10) || 0, parseInt(v[1], 10) || 0] : [0, 0]; }
    return {
      online: pair(s.online), offline: pair(s.offline),
      generalCol: parseInt(s.generalCol, 10) || 0, otherCol: parseInt(s.otherCol, 10) || 0,
      catCol: parseInt(s.catCol, 10) || 0,
      items: s.items && typeof s.items === 'object' && !Array.isArray(s.items) ? s.items : {}
    };
  }
  function saveSettings() {
    if (!S.pick) return;
    var items = {};
    S.pick.items.forEach(function (it) { if (it.row) items[it.name] = it.row; });
    settings = {
      online: S.pick.online.slice(), offline: S.pick.offline.slice(),
      generalCol: S.pick.generalCol, otherCol: S.pick.otherCol,
      catCol: S.pick.catCol, items: items
    };
    try { IEApp.settings.write(STORE_KEY, settings); } catch (e) { /* 存不了就算了 */ }
  }

  function sameName(a, b) { return String(a).toLowerCase() === String(b).toLowerCase(); }
  function isExtra(name) { return sameName(name, REF_SHEET) || sameName(name, SHEET_NAME); }

  // 前面步驟的結果：第 5 步（計算需求工時）
  function previous() {
    return IEApp.state.withDemand || null;
  }
  function previousLabel() {
    return '第 5 步剛做好的結果';
  }

  function activeSource() {
    var prev = previous();
    if (S.base && (S.preferUpload || !prev)) return { kind: 'file', base: S.base, name: S.base.name };
    if (prev) return { kind: 'prev', prev: prev, name: prev.outName };
    return null;
  }

  /* ---------- 畫面 ---------- */
  function render(panel) {
    els = { rangeSelects: {}, colSelects: {} };

    els.baseInfo = h('p', { class: 'source-line', 'aria-live': 'polite' });
    els.baseDrop = ui.fileDrop({
      id: 'f8-base',
      title: '上傳第 5 步匯出的 Excel',
      sub: '或把檔案拖進這裡。支援 .xlsx',
      onFile: openBase
    });
    els.baseSwitch = h('button', { type: 'button', class: 'btn-link', hidden: true, onclick: switchSource });
    els.backBtn = h('button', {
      type: 'button', class: 'btn-link', hidden: true, text: '回到第 5 步計算需求工時',
      onclick: function () { IEApp.go('demand-hours'); }
    });
    els.baseMsg = h('p', { class: 'msg is-error', role: 'alert', hidden: true });

    els.mpDrop = ui.fileDrop({
      id: 'f8-file',
      title: '選擇人力報表的 Excel 檔',
      sub: '或把檔案拖進這裡。支援 .xlsx',
      onFile: openManpower
    });
    els.sheetSelect = h('select', { id: 'f8-sheet', 'aria-label': '人力報表的工作表' });
    els.sheetRow = h('div', { class: 'field-row', hidden: true },
      h('label', { for: 'f8-sheet', text: '工作表' }), els.sheetSelect);
    els.mpMsg = h('p', { class: 'msg is-error', role: 'alert', hidden: true });

    var rows = GROUPS.map(function (g) {
      var from = h('select', { id: 'f8-' + g.key + '-from', class: 'in-select in-range', 'aria-label': g.label + ' 從第幾列' });
      var to = h('select', { id: 'f8-' + g.key + '-to', class: 'in-select in-range', 'aria-label': g.label + ' 到第幾列' });
      from.addEventListener('change', function () { onRangeChange(g.key, 0, from.value); });
      to.addEventListener('change', function () { onRangeChange(g.key, 1, to.value); });
      els.rangeSelects[g.key] = [from, to];
      return h('div', { class: 'field-row' },
        h('label', { for: 'f8-' + g.key + '-from', class: 'label-fixed', text: '「' + g.label + '」從' }), from,
        h('label', { for: 'f8-' + g.key + '-to', text: '到' }), to);
    });
    var cols = COLS.map(function (c) {
      var sel = h('select', { id: 'f8-' + c.key, class: 'in-select' });
      sel.addEventListener('change', function () { onColChange(c.key, sel.value); });
      els.colSelects[c.key] = sel;
      return h('div', { class: 'field-row' },
        h('label', { for: 'f8-' + c.key, class: 'label-fixed', text: c.label + '在' }), sel);
    });

    els.catSelect = h('select', { id: 'f8-catCol', class: 'in-select' });
    els.catSelect.addEventListener('change', function () { onCatChange(els.catSelect.value); });
    var catRow = h('div', { class: 'field-row' },
      h('label', { for: 'f8-catCol', class: 'label-fixed', text: '「類別」在' }), els.catSelect);
    els.items = h('div', { class: 'cat-items' });

    els.preview = h('div', { class: 'sheet-area' });
    els.check = h('div', { class: 'sheet-area' });

    els.summary = h('p', { class: 'summary', 'aria-live': 'polite' });
    els.notes = h('ul', { class: 'notes' });
    els.exportBtn = h('button', { type: 'button', class: 'btn-primary', disabled: true, text: '匯出 Excel', onclick: onExport });
    els.nextBtn = h('button', {
      type: 'button', class: 'btn-secondary', disabled: true, text: '下一步：總表',
      onclick: function () { IEApp.go('report'); }
    });
    els.exportNote = h('span', { class: 'export-note', 'aria-live': 'polite' });

    panel.appendChild(h('div', { class: 'step-head' },
      h('h1', { text: '匯入人力報表' }),
      h('p', { text: '把人力報表（現有人力 - 鍍膜）放進前面做好的 Excel，成為同一個檔案裡名為「人力」的工作表。匯入後確認 on line、off line 各是哪幾列，以及「一般鍍膜」「＿＿鍍膜」各在哪一行，之後做總表時就照這些位置抓資料。' })));

    panel.appendChild(h('ol', { class: 'tasks' },
      h('li', { class: 'task' },
        h('h2', { text: '第 5 步的 Excel' }),
        els.baseInfo, els.baseDrop.el,
        h('div', { class: 'link-row' }, els.baseSwitch, els.backBtn),
        els.baseMsg),
      h('li', { class: 'task' },
        h('h2', { text: '上傳人力報表' }),
        els.mpDrop.el, els.sheetRow, els.mpMsg),
      h('li', { class: 'task' },
        h('h2', { text: '確認位置' }),
        h('p', { class: 'hint', text: '程式會先自動帶入：寫 on line、off line 的格子（合併的範圍就是那幾列），以及表頭寫「一般鍍膜」和另一個「…鍍膜」的那兩行；類別行裡的每個項目也會列出各在第幾列。請對照下面的表格確認，不對就改。' }),
        rows, cols, catRow, els.items, els.preview, els.check),
      h('li', { class: 'task' },
        h('h2', { text: '匯出' }),
        els.summary, els.notes,
        h('div', { class: 'actions' }, els.exportBtn, els.nextBtn, els.exportNote))));

    els.sheetSelect.addEventListener('change', function () {
      pickSheet(parseInt(els.sheetSelect.value, 10));
    });

    if (S.mp) showManpowerFile();
    fillSelects();
    refresh();
  }

  /* ---------- 第 5 步的 Excel ---------- */
  function openBase(file) {
    els.baseMsg.hidden = true;
    els.baseDrop.setBusy(file.name);
    ui.readWorkbook(file)
      .then(function (res) {
        var names = res.workbook.worksheets.map(function (ws) { return ws.name; });
        var others = names.filter(function (n) { return !isExtra(n); });
        if (!others.length) throw new Error('這個檔案沒有 MPS 工作表，看起來不是第 5 步匯出的檔案。請上傳第 5 步匯出的 Excel。');
        S.base = {
          name: file.name, buffer: res.buffer, sheets: names,
          hasManpower: names.some(function (n) { return sameName(n, SHEET_NAME); })
        };
        S.preferUpload = true;
        els.exportNote.textContent = '';
        refresh();
      })
      .catch(function (err) {
        els.baseMsg.textContent = err && err.message ? err.message : String(err);
        els.baseMsg.hidden = false;
        refresh();
      });
  }

  function switchSource() {
    S.preferUpload = !S.preferUpload;
    els.exportNote.textContent = '';
    refresh();
  }

  /* ---------- 人力報表 ---------- */
  function openManpower(file) {
    els.mpMsg.hidden = true;
    els.mpDrop.setBusy(file.name);
    ui.readWorkbook(file)
      .then(function (res) {
        S.mp = { name: file.name, workbook: res.workbook, ws: null, ext: null, labels: {} };
        showManpowerFile();
        // 檔案裡有叫「人力」的工作表就先選那一張，沒有就選第一張
        var sheets = res.workbook.worksheets;
        var named = sheets.filter(function (ws) { return sameName(ws.name, SHEET_NAME); })[0];
        pickSheet(named ? sheets.indexOf(named) : 0);
      })
      .catch(function (err) {
        els.mpMsg.textContent = err && err.message ? err.message : String(err);
        els.mpMsg.hidden = false;
        if (S.mp) showManpowerFile(); else els.mpDrop.setIdle();
        refresh();
      });
  }

  function showManpowerFile() {
    els.mpDrop.setFile(S.mp.name);
    var sheets = S.mp.workbook.worksheets;
    els.sheetSelect.textContent = '';
    sheets.forEach(function (ws, i) {
      els.sheetSelect.appendChild(h('option', { value: String(i), text: ws.name }));
    });
    els.sheetRow.hidden = sheets.length < 2;
    if (S.mp.ws) els.sheetSelect.value = String(sheets.indexOf(S.mp.ws));
  }

  // 比對用：小寫、去掉空白和橫線（on line = Online = ON-LINE）
  function squash(t) { return XIO.normalizeText(t == null ? '' : String(t)).toLowerCase().replace(/[\s\-_－]/g, ''); }
  function cellText(cell) { return cell ? XIO.normalizeText(XIO.displayText(cell)).replace(/\s+/g, ' ') : ''; }

  // 找第一個文字符合 test 的儲存格（合併的只看左上角那一格）；onlyRow 有給就只找那一列
  // 回傳 { row, col, bottom, right }（合併時是合併的範圍）或 null
  function findCellBy(ws, ext, test, onlyRow) {
    var merges = XIO.listMerges(ws);
    var r1 = onlyRow || 1, r2 = onlyRow || ext.lastRow;
    for (var r = r1; r <= r2; r++) {
      var row = ws.findRow(r);
      if (!row) continue;
      for (var c = 1; c <= ext.colCount; c++) {
        var cell = row.findCell(c);
        if (!cell || (cell.isMerged && cell.master !== cell)) continue;
        if (!test(squash(cellText(cell)))) continue;
        var m = merges.filter(function (x) { return x.top === r && x.left === c; })[0];
        return { row: r, col: c, bottom: m ? m.bottom : r, right: m ? m.right : c };
      }
    }
    return null;
  }

  // 自動帶入四項：{ online: [起, 迄], offline: [起, 迄], generalCol, otherCol, headRow, catCol }（找不到的是 0）
  function detect(ws, ext) {
    var out = { online: [0, 0], offline: [0, 0], generalCol: 0, otherCol: 0, headRow: 0, catCol: 0 };
    var on = findCellBy(ws, ext, function (t) { return t === 'online'; });
    var off = findCellBy(ws, ext, function (t) { return t === 'offline'; });
    if (on) {
      out.online = [on.row, on.bottom > on.row ? on.bottom : (off && off.row > on.row ? off.row - 1 : ext.lastRow)];
      out.catCol = on.right + 1;
    }
    if (off) {
      out.offline = [off.row, off.bottom > off.row ? off.bottom : ext.lastRow];
      if (!out.catCol) out.catCol = off.right + 1;
    }
    var general = findCellBy(ws, ext, function (t) { return t.indexOf('一般鍍膜') >= 0; });
    if (general) {
      out.generalCol = general.col;
      out.headRow = general.row;
      var other = findCellBy(ws, ext, function (t) { return t.indexOf('鍍膜') >= 0 && t.indexOf('一般') < 0; }, general.row);
      if (other) out.otherCol = other.col;
      var cat = findCellBy(ws, ext, function (t) { return t === '類別'; }, general.row);
      if (cat) out.catCol = cat.col;
    }
    return out;
  }

  // 類別行裡的項目：on line、off line 範圍內，類別那一行有文字的格子（合併的只算第一列）
  // 回傳 [{ name, row, group: 'online' | 'offline' }]，依列的順序
  function detectItems(pick) {
    var mp = S.mp, out = [];
    if (!pick.catCol) return out;
    GROUPS.forEach(function (g) {
      var p = pick[g.key];
      if (!(p[0] > 0) || !(p[1] >= p[0])) return;
      for (var r = p[0]; r <= p[1]; r++) {
        var row = mp.ws.findRow(r);
        var cell = row ? row.findCell(pick.catCol) : null;
        if (!cell || (cell.isMerged && cell.master !== cell)) continue;
        var t = cellText(cell);
        if (!t || squash(t) === 'online' || squash(t) === 'offline') continue;
        out.push({ name: t, row: r, group: g.key });
      }
    });
    return out;
  }
  // 自動找到的項目，套上上次使用者改過的列（依名稱；還在報表範圍內才用）
  function applyRemembered(items) {
    items.forEach(function (it) {
      var r = parseInt(settings.items[it.name], 10);
      if (r >= 1 && r <= S.mp.ext.lastRow) it.row = r;
    });
    return items;
  }

  function pickSheet(index) {
    var mp = S.mp;
    mp.ws = mp.workbook.worksheets[index];
    mp.ext = XIO.extent(mp.ws);
    mp.labels = {};
    XIO.manpowerRows(mp.ws, mp.ext).forEach(function (x) { mp.labels[x.row] = x.text; });
    els.sheetSelect.value = String(index);
    els.exportNote.textContent = '';

    // 預設值：先自動帶入；帶不出來的，用上次記住的（還在報表範圍內才用）
    var d = detect(mp.ws, mp.ext);
    var maxR = mp.ext.lastRow, maxC = mp.ext.colCount;
    function okRange(p) { return p[0] >= 1 && p[1] >= p[0] && p[1] <= maxR; }
    S.pick = {
      online: okRange(d.online) ? d.online : okRange(settings.online) ? settings.online.slice() : [0, 0],
      offline: okRange(d.offline) ? d.offline : okRange(settings.offline) ? settings.offline.slice() : [0, 0],
      generalCol: d.generalCol || (settings.generalCol <= maxC ? settings.generalCol : 0),
      otherCol: d.otherCol || (settings.otherCol <= maxC ? settings.otherCol : 0),
      headRow: d.headRow,
      catCol: d.catCol || (settings.catCol <= maxC ? settings.catCol : 0),
      items: []
    };
    S.pick.items = applyRemembered(detectItems(S.pick));
    fillSelects();
    refresh();
  }

  function shorten(text, max) {
    return text.length > max ? text.slice(0, max - 1) + '…' : text;
  }

  // 每一行的表頭文字：有找到「一般鍍膜」那一列就用那一列，沒有就從選的第一列往上找
  function colLabels() {
    var mp = S.mp;
    if (!mp || !mp.ws) return [];
    var labels = [];
    if (S.pick.headRow) {
      var row = mp.ws.findRow(S.pick.headRow);
      for (var c = 1; c <= mp.ext.colCount; c++) labels[c] = row ? cellText(row.findCell(c)) : '';
      return labels;
    }
    var top = S.pick.online[0] || S.pick.offline[0] || mp.ext.lastRow + 1;
    return XIO.labelsAbove(mp.ws, top, mp.ext.colCount);
  }

  function fillSelects() {
    var mp = S.mp;
    var ready = !!(mp && mp.ws);
    GROUPS.forEach(function (g) {
      els.rangeSelects[g.key].forEach(function (sel, k) {
        sel.textContent = '';
        if (!ready) { sel.appendChild(h('option', { value: '', text: '先上傳人力報表' })); sel.disabled = true; return; }
        sel.disabled = false;
        sel.appendChild(h('option', { value: '', text: '請選擇' }));
        for (var r = 1; r <= mp.ext.lastRow; r++) {
          sel.appendChild(h('option', {
            value: String(r),
            text: '第 ' + r + ' 列' + (mp.labels[r] ? '　' + shorten(mp.labels[r], 30) : '')
          }));
        }
        var v = S.pick[g.key][k];
        sel.value = v ? String(v) : '';
      });
    });
    var labels = ready ? colLabels() : [];
    COLS.forEach(function (c) {
      var sel = els.colSelects[c.key];
      sel.textContent = '';
      if (!ready) { sel.appendChild(h('option', { value: '', text: '先上傳人力報表' })); sel.disabled = true; return; }
      sel.disabled = false;
      sel.appendChild(h('option', { value: '', text: '請選擇' }));
      for (var col = 1; col <= mp.ext.colCount; col++) {
        sel.appendChild(h('option', {
          value: String(col),
          text: XIO.numToCol(col) + ' 行' + (labels[col] ? '　' + shorten(labels[col], 30) : '')
        }));
      }
      sel.value = S.pick[c.key] ? String(S.pick[c.key]) : '';
    });

    // 類別在哪一行
    var cs = els.catSelect;
    cs.textContent = '';
    if (!ready) { cs.appendChild(h('option', { value: '', text: '先上傳人力報表' })); cs.disabled = true; }
    else {
      cs.disabled = false;
      cs.appendChild(h('option', { value: '', text: '請選擇' }));
      for (var cc = 1; cc <= mp.ext.colCount; cc++) {
        cs.appendChild(h('option', { value: String(cc), text: XIO.numToCol(cc) + ' 行' + (labels[cc] ? '　' + shorten(labels[cc], 30) : '') }));
      }
      cs.value = S.pick.catCol ? String(S.pick.catCol) : '';
    }
    fillItems();
  }

  // 類別項目的清單：每個項目一列，可以改在第幾列
  function fillItems() {
    var box = els.items;
    box.textContent = '';
    var mp = S.mp;
    if (!mp || !mp.ws || !S.pick.catCol) return;
    var items = S.pick.items;
    if (!items.length) {
      box.appendChild(h('p', { class: 'hint', text: '在這一行的 on line、off line 範圍裡找不到類別項目。請確認「類別」在哪一行，以及 on line、off line 的列。' }));
      return;
    }
    var body = h('tbody');
    items.forEach(function (it, i) {
      var sel = h('select', { class: 'in-select in-range', 'aria-label': it.name + ' 在第幾列' });
      sel.appendChild(h('option', { value: '', text: '請選擇' }));
      for (var r = 1; r <= mp.ext.lastRow; r++) {
        sel.appendChild(h('option', { value: String(r), text: '第 ' + r + ' 列' + (mp.labels[r] ? '　' + shorten(mp.labels[r], 24) : '') }));
      }
      sel.value = it.row ? String(it.row) : '';
      sel.addEventListener('change', function () {
        S.pick.items[i].row = parseInt(sel.value, 10) || 0;
        els.exportNote.textContent = '';
        refresh();
      });
      body.appendChild(h('tr', null,
        h('td', { text: it.group === 'online' ? 'on line' : 'off line' }),
        h('th', { scope: 'row', text: it.name }),
        h('td', null, sel)));
    });
    box.appendChild(h('p', { class: 'sheet-caption', text: '類別行裡找到的項目（' + items.length + ' 個），各在人力報表的第幾列：' }));
    box.appendChild(h('table', { class: 'grid-inputs cat-table' },
      h('thead', null, h('tr', null, h('th', { text: '大項' }), h('th', { text: '類別' }), h('th', { text: '在第幾列' }))), body));
  }

  function onCatChange(value) {
    if (!S.pick) return;
    S.pick.catCol = parseInt(value, 10) || 0;
    S.pick.items = applyRemembered(detectItems(S.pick));
    fillItems();
    els.exportNote.textContent = '';
    refresh();
  }

  function onRangeChange(key, k, value) {
    if (!S.pick) return;
    S.pick[key][k] = parseInt(value, 10) || 0;
    S.pick.items = applyRemembered(detectItems(S.pick));
    fillItems();
    els.exportNote.textContent = '';
    refresh();
  }

  function onColChange(key, value) {
    if (!S.pick) return;
    S.pick[key] = parseInt(value, 10) || 0;
    els.exportNote.textContent = '';
    refresh();
  }

  // 讀出某一格的人數：{ text, number, blank, bad }（空白當成 0）
  function readCell(r, c) {
    var row = S.mp.ws.findRow(r);
    var cell = row ? row.findCell(c) : null;
    var v = cell ? (cell.isMerged ? cell.master.value : cell.value) : null;
    if (v && typeof v === 'object' && v.formula !== undefined) v = v.result;      // 公式：用計算結果
    v = XIO.plainValue(v);
    if (v === null || v === undefined || v === '') return { text: '', number: 0, blank: true };
    var n = typeof v === 'number' ? v : Number(XIO.normalizeText(String(v)).replace(/,/g, ''));
    if (isNaN(n)) return { text: String(v), number: 0, bad: true };
    return { text: String(Math.round(n * 1000) / 1000), number: n };
  }

  // 類別：使用者選的類別那一行的文字（合併格只在第一列顯示）
  function categoryText(r) {
    if (!S.pick.catCol) return '';
    var row = S.mp.ws.findRow(r);
    var cell = row ? row.findCell(S.pick.catCol) : null;
    if (!cell || (cell.isMerged && cell.master !== cell)) return '';
    var t = cellText(cell);
    return t && squash(t) !== 'online' && squash(t) !== 'offline' ? t : '';
  }

  /* ---------- 更新畫面 ---------- */
  function refresh() {
    var src = activeSource();
    var prev = previous();
    IEApp.state.withManpower = null;
    els.exportBtn.disabled = true;
    els.nextBtn.disabled = true;
    els.notes.textContent = '';
    els.preview.textContent = '';
    els.check.textContent = '';

    // 第 5 步的 Excel
    if (!src) {
      els.baseInfo.textContent = '還沒有第 5 步的結果。回到第 5 步，或上傳之前匯出的檔案。';
      els.baseDrop.setIdle();
    } else if (src.kind === 'prev') {
      els.baseInfo.textContent = '沿用' + previousLabel() + '：' + prev.outName + '。不用先匯出；要改用之前匯出的檔案，再從下面上傳。';
      if (S.base) els.baseDrop.setFile(S.base.name, '已上傳，目前沒有使用。點這裡換一個檔案');
      else els.baseDrop.setIdle();
    } else {
      els.baseInfo.textContent = '使用上傳的檔案，裡面的工作表：' + S.base.sheets.join('、') + '。';
      els.baseDrop.setFile(S.base.name);
    }
    els.backBtn.hidden = !!src;
    els.baseSwitch.hidden = !(S.base && prev);
    if (src) els.baseSwitch.textContent = src.kind === 'file' ? '改用' + previousLabel() : '改用上傳的檔案';

    var mp = S.mp;
    var hasSheet = !!(mp && mp.ws && mp.ext.lastRow);
    var pick = S.pick;

    // 選好的列、行
    var keepRows = {};
    var missing = [];
    var badRange = [];
    if (hasSheet) {
      GROUPS.forEach(function (g) {
        var p = pick[g.key];
        if (!p[0] || !p[1]) { missing.push(g.label + ' 從第幾列到第幾列'); return; }
        if (p[1] < p[0]) { badRange.push(g.label); return; }
        for (var r = p[0]; r <= p[1]; r++) keepRows[r] = true;
      });
      COLS.forEach(function (c) { if (!pick[c.key]) missing.push(c.label + '在哪一行'); });
      if (!pick.catCol) missing.push('「類別」在哪一行');
    }

    // 人力報表預覽：選到的列 × 兩行畫螢光
    if (hasSheet) {
      var keep = {};
      COLS.forEach(function (c) { if (pick[c.key]) keep[pick[c.key]] = true; });
      var lastPicked = Math.max(pick.online[1] || 0, pick.offline[1] || 0);
      ui.renderSheet(els.preview, mp.ws, {
        headerEnd: 0,
        lastRow: mp.ext.lastRow,
        colCount: mp.ext.colCount,
        rows: Math.min(Math.max(MIN_PREVIEW_ROWS, lastPicked + 2), MAX_PREVIEW_ROWS),
        keep: keep,
        keepRows: keepRows,
        caption: '人力報表。共 ' + mp.ext.lastRow.toLocaleString('en-US') + ' 列、' + mp.ext.colCount +
          ' 行。螢光的格子是之後總表要抓的資料。'
      });
    }

    // 核對表
    var read = null;
    if (hasSheet && !badRange.length && (pick.online[0] || pick.offline[0])) read = renderCheck();

    // 匯出
    if (!src) { setSummary('先準備好第 5 步的 Excel。'); return; }
    if (!mp || !mp.ws) { setSummary('上傳人力報表。'); return; }
    if (!hasSheet) { setSummary('人力報表的這個工作表沒有資料。請確認檔案，或換一個工作表。', true); return; }
    if (badRange.length) { setSummary(badRange.join('、') + ' 的「到」比「從」還前面，請重新選。', true); return; }
    if (missing.length) { setSummary('還沒有選：' + missing.join('、') + '。'); return; }

    // 注意事項
    var warn = false;
    var on = pick.online, off = pick.offline;
    if (on[0] <= off[1] && off[0] <= on[1]) {
      warn = true;
      els.notes.appendChild(h('li', { class: 'is-warn' }, 'on line 和 off line 的列有重疊，請確認。'));
    }
    if (pick.generalCol === pick.otherCol) {
      warn = true;
      els.notes.appendChild(h('li', { class: 'is-warn' }, '「一般鍍膜」和「＿＿鍍膜」選了同一行，請確認。'));
    }
    if (!pick.items.length) {
      warn = true;
      els.notes.appendChild(h('li', { class: 'is-warn' }, '類別行裡沒有找到任何項目，請確認「類別」在哪一行。'));
    }
    var noRow = pick.items.filter(function (it) { return !it.row; }).map(function (it) { return it.name; });
    if (noRow.length) {
      warn = true;
      els.notes.appendChild(h('li', { class: 'is-warn' }, '還沒選在第幾列：' + noRow.join('、') + '。'));
    }
    var byRow = {};
    pick.items.forEach(function (it) { if (it.row) (byRow[it.row] = byRow[it.row] || []).push(it.name); });
    Object.keys(byRow).forEach(function (r) {
      if (byRow[r].length < 2) return;
      warn = true;
      els.notes.appendChild(h('li', { class: 'is-warn' }, byRow[r].join('、') + ' 選了同一列（第 ' + r + ' 列），請確認。'));
    });
    if (read && read.bad) {
      warn = true;
      els.notes.appendChild(h('li', { class: 'is-warn' },
        '有 ' + read.bad + ' 格不是數字（上面表格標成淺紅色）。請確認列和行有沒有選對。'));
    }

    var replacing = src.kind === 'file' && S.base.hasManpower;
    var text = '匯出的檔案會多一個「' + SHEET_NAME + '」工作表（' + mp.ext.lastRow.toLocaleString('en-US') + ' 列、' +
      mp.ext.colCount + ' 行）。on line 第 ' + on[0] + '–' + on[1] + ' 列、off line 第 ' + off[0] + '–' + off[1] + ' 列；' +
      '一般鍍膜在 ' + XIO.numToCol(pick.generalCol) + ' 行、＿＿鍍膜在 ' + XIO.numToCol(pick.otherCol) + ' 行；' +
      '類別在 ' + XIO.numToCol(pick.catCol) + ' 行，共 ' + pick.items.length + ' 個項目。';
    if (replacing) text += ' 檔案裡原本的「' + SHEET_NAME + '」工作表會換成這次上傳的。';
    setSummary(text, false, warn || replacing);
    els.exportBtn.disabled = false;

    // 給下一步用
    var sheets = src.kind === 'prev'
      ? [prev.mainSheet, REF_SHEET]
      : S.base.sheets.filter(function (n) { return !sameName(n, SHEET_NAME); });
    sheets = sheets.concat([SHEET_NAME]);
    var sheet = mp.ws;
    IEApp.state.withManpower = {
      outName: outputName(src),
      sheets: sheets,
      mainSheet: src.kind === 'prev' ? prev.mainSheet : S.base.sheets.filter(function (n) { return !isExtra(n); })[0],
      dataStartRow: src.kind === 'prev' ? prev.dataStartRow : 0,
      keyCol: src.kind === 'prev' && prev.keyCol ? prev.keyCol : 1,
      manpowerSheet: SHEET_NAME,
      manpower: {
        onlineRows: on.slice(), offlineRows: off.slice(),
        generalCol: pick.generalCol, otherCol: pick.otherCol,
        catCol: pick.catCol,
        categories: pick.items.map(function (it) { return { name: it.name, row: it.row, group: it.group }; })
      },
      build: function () { return buildWorkbook(src, sheet); }
    };
    saveSettings();       // 都選好了，記下來給下個月當預設值
    els.nextBtn.disabled = false;
  }

  function setSummary(text, isError, isWarn) {
    els.summary.textContent = text;
    els.summary.className = 'summary' + (isError ? ' is-error' : '') + (isWarn ? ' is-warn' : '');
  }

  // 核對表：on line、off line 每一列 × 兩行的人數，加上各自的合計。回傳 { bad }
  function renderCheck() {
    var pick = S.pick;
    var labels = colLabels();
    var cols = COLS.map(function (c) { return pick[c.key]; });
    var head = h('tr', null,
      h('th', { class: 'colhead', scope: 'col', text: '大項' }),
      h('th', { class: 'colhead', scope: 'col', text: '列' }),
      h('th', { class: 'colhead', scope: 'col', text: '類別' }));
    COLS.forEach(function (c, i) {
      var col = cols[i];
      head.appendChild(h('th', { class: 'colhead', scope: 'col',
        text: col ? XIO.numToCol(col) + ' 行' + (labels[col] ? '　' + labels[col] : '') : c.label + '（還沒選）' }));
    });

    var body = h('tbody');
    var bad = 0;
    GROUPS.forEach(function (g) {
      var p = pick[g.key];
      if (!p[0] || !p[1] || p[1] < p[0]) return;
      var sums = cols.map(function () { return 0; });
      for (var r = p[0]; r <= p[1]; r++) {
        var tr = h('tr', null,
          h('th', { class: 'rowhead', scope: 'row', text: r === p[0] ? g.label : '' }),
          h('td', { text: '第 ' + r + ' 列' }),
          h('td', { text: categoryText(r) }));
        cols.forEach(function (col, i) {
          if (!col) { tr.appendChild(h('td')); return; }
          var info = readCell(r, col);
          if (info.bad) bad++;
          sums[i] += info.number;
          tr.appendChild(h('td', { class: info.bad ? 'is-missing' : 'is-num', text: info.text }));
        });
        body.appendChild(tr);
      }
      var total = h('tr', { class: 'is-total' },
        h('th', { class: 'rowhead', scope: 'row', text: g.label + ' 合計' }), h('td'), h('td'));
      cols.forEach(function (col, i) {
        total.appendChild(h('td', { class: 'is-num', text: col ? String(Math.round(sums[i] * 1000) / 1000) : '' }));
      });
      body.appendChild(total);
    });

    els.check.appendChild(h('p', { class: 'sheet-caption', text: '照上面選的位置讀到的人數（空白當成 0）：' }));
    els.check.appendChild(h('div', { class: 'sheet-scroll', tabindex: '0', role: 'group', 'aria-label': 'on line、off line 的人數' },
      h('table', { class: 'sheet is-result is-check' }, h('thead', null, head), body)));
    return { bad: bad };
  }

  /* ---------- 產生活頁簿 ---------- */
  // 每次都從頭產生一份新的：前面步驟的內容 + 人力工作表。回傳 Promise<Workbook>
  function buildWorkbook(src, sheet) {
    var base = src.kind === 'prev'
      ? Promise.resolve(src.prev.build())
      : XIO.load(src.base.buffer.slice(0));
    return base.then(function (wb) {
      XIO.copySheet(sheet, wb, SHEET_NAME);
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
    var ready = IEApp.state.withManpower;
    if (!ready) return;
    els.exportBtn.disabled = true;
    els.exportNote.textContent = '產生檔案中…';
    ready.build()
      .then(function (wb) { return XIO.toBlob(wb); })
      .then(function (blob) {
        ui.download(blob, ready.outName);
        els.exportNote.textContent = '已匯出 ' + ready.outName + '（工作表：' + ready.sheets.join('、') + '），在瀏覽器的下載資料夾。';
      })
      .catch(function (err) {
        els.exportNote.textContent = '匯出失敗：' + (err && err.message ? err.message : err);
      })
      .then(function () {
        els.exportBtn.disabled = false;
      });
  }

  IEApp.registerStep({ id: 'manpower', render: render });
})(window);
