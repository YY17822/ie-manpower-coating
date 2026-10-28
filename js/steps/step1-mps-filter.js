/*
 * 第 1 步：篩選 MPS 報表（Coating）
 *   1. 上傳一個或多個 MPS 的 Excel（例如塑膠一個檔、玻璃一個檔）
 *   2. 每個檔案各選一種篩選方式，把列分到 Plastic、Glass 兩段：
 *        a. 依製程（Plastic）：程式在料號那一行自動找「TTL (製程名)」列，
 *           TTL (Plastic Lens-Coating) 上面那一塊算 Plastic。
 *        b. 依製程（Glass）：TTL (GMO-Coating)、TTL (Filter&UVIR-Coating) 上面那兩塊算 Glass（先 GMO、再 Filter&UVIR）。
 *           每個製程最後一列料號的下面是「x」，再下面才是 TTL 列；「x」和 TTL 列不會留下。
 *        c. 依 Status：使用者輸入 A 行等於多少代表 Plastic、等於多少代表 Glass。
 *   3. 每個檔案各自輸入要保留哪些行（位置可以不同，但個數和順序要一樣）
 *   4. 匯出新的 Excel：第 1–4 列表頭照第 1 個檔案保留，下面分兩段——
 *        第 5 列     標題「Plastic（Plastic Lens-Coating）」，跨所有保留的行置中
 *        接著        Plastic 的列（第 1 個檔案的在前，依檔案順序接下去）
 *        空 10 列    留給第 2 步的匯總表格
 *        接著        標題「Glass（GMO-Coating、Filter&UVIR-Coating）」，下面是 Glass 的列
 *
 * 做完之後把結果放在 IEApp.state.mpsFiltered（sections 是兩段各在第幾列），第 2 步直接拿來用。
 *
 * 名詞：行 = 直的（A、B、C…），列 = 橫的（1、2、3…）。
 */
(function (global) {
  'use strict';

  var IEApp = global.IEApp;
  var XIO = global.ExcelIO;
  var ui = IEApp.ui;
  var h = IEApp.h;

  var STORE_KEY = 'step1';             // 共用設定的組名（IEApp.settings）
  var PREVIEW_ROWS = 8;       // 原始報表預覽顯示幾列資料
  var OUT_PREVIEW_MAX = 40;   // 匯出預覽最多顯示幾列（要看得到兩段）
  var MAX_CHIPS = 24;

  // 輸出的兩段。製程名稱可以在畫面的「製程名稱不一樣時」改，標題會跟著變。
  var SECTIONS = [
    { key: 'plastic', label: 'Plastic', procs: ['plasticProc'], valueSetting: 'plasticValues' },
    { key: 'glass', label: 'Glass', procs: ['gmoProc', 'filterProc'], valueSetting: 'glassValues' }
  ];

  // 三種篩選方式
  var MODES = [
    { value: 'plastic', title: '依製程', sub: '自動找 TTL 列，分出 Plastic Lens-Coating' },
    { value: 'glass', title: '依製程', sub: '自動找 TTL 列，分出 GMO-Coating、Filter&UVIR-Coating' },
    { value: 'status', title: '依 Status', sub: '輸入 A 行等於多少代表 Plastic、Glass' }
  ];
  var MODE_SECTION = { plastic: 'plastic', glass: 'glass' };   // 依製程的兩種，各自分到哪一段

  // 整個步驟共用的設定
  var DEFAULTS = {
    files: [],
    plasticProc: 'Plastic Lens-Coating',
    gmoProc: 'GMO-Coating',
    filterProc: 'Filter&UVIR-Coating'
  };
  // 每個檔案自己的設定（照上傳順序：第 1 個檔案用 files[0]…）
  var FILE_DEFAULTS = {
    mode: 'plastic', plasticValues: '', glassValues: '',
    cols: '', startRow: 5, selCol: 'A', pnCol: 'B'
  };

  // 這一步的狀態（切到別的步驟再回來時還在）
  // S.files: [{ name, workbook, ws, info, infoKey, blocks, blocksKey, picked, keepCols }]
  var S = { files: [], result: null };
  var settings = loadSettings();
  var els = {};
  var timer = null;

  /* ---------- 記住上次輸入的條件（每個月做的事情差不多） ---------- */
  function fileDefaults(i) {
    var d = Object.assign({}, FILE_DEFAULTS);
    if (i === 1) d.mode = 'glass';      // 第 2 個檔案預設是玻璃
    return d;
  }
  function normFile(f, i) {
    var d = fileDefaults(i);
    var out = {};
    f = f && typeof f === 'object' ? f : {};
    Object.keys(d).forEach(function (k) { out[k] = f[k] === undefined ? d[k] : f[k]; });
    if (!MODES.some(function (m) { return m.value === out.mode; })) out.mode = d.mode;
    return out;
  }
  function loadSettings() {
    var s = {};
    try { s = IEApp.settings.read(STORE_KEY) || {}; } catch (e) { s = {}; }
    var out = {};
    Object.keys(DEFAULTS).forEach(function (k) { out[k] = s[k] === undefined ? DEFAULTS[k] : s[k]; });
    out.files = (Array.isArray(out.files) ? out.files : []).map(normFile);
    return out;
  }
  function saveSettings() {
    try { IEApp.settings.write(STORE_KEY, settings); } catch (e) { /* 存不了就算了 */ }
  }
  // 第 i 個檔案的設定（不夠就補預設值）
  function fset(i) {
    while (settings.files.length <= i) settings.files.push(normFile(null, settings.files.length));
    return settings.files[i];
  }

  /* ---------- 讀取設定值 ---------- */
  function intOr(v, fallback) {
    var n = parseInt(XIO.normalizeText(v), 10);
    return n >= 1 ? n : fallback;
  }
  function oneCol(text, fallback) {
    var p = XIO.parseColumns(text);
    return p.cols.length === 1 && !p.bad.length ? p.cols[0] : fallback;
  }
  function startRowOf(fs) { return intOr(fs.startRow, FILE_DEFAULTS.startRow); }
  function selColOf(fs) { return oneCol(fs.selCol, 1); }
  function pnColOf(fs) { return oneCol(fs.pnCol, 2); }
  function valueList(text) {
    return XIO.normalizeText(text || '').split(/[,、;]+/)
      .map(function (v) { return v.trim(); })
      .filter(function (v) { return v !== ''; });
  }
  function procName(setting) {
    return XIO.normalizeText(settings[setting] || '') || DEFAULTS[setting];
  }
  function sectionTitle(sec) {
    return sec.label + '（' + sec.procs.map(procName).join('、') + '）';
  }
  function sectionOf(key) { return SECTIONS.filter(function (s) { return s.key === key; })[0]; }
  function outputName() {
    var first = S.files[0] ? S.files[0].name : 'MPS.xlsx';
    return first.replace(/\.[^.]+$/, '') + '_篩選.xlsx';
  }

  /* ---------- 畫面 ---------- */
  function render(panel) {
    els = { files: [] };

    els.drop = ui.fileDrop({
      id: 'f-file',
      title: '選擇 MPS 的 Excel 檔（可以一次選多個）',
      sub: '或把檔案拖進這裡。支援 .xlsx',
      multiple: true,
      onFiles: openFiles
    });
    els.fileList = h('ul', { class: 'file-list' });
    els.fileMsg = h('p', { class: 'msg is-error', role: 'alert', hidden: true });
    els.fileTasks = h('div', { class: 'file-tasks' });

    els.procInputs = {};
    var procRows = [
      ['plasticProc', 'Plastic 的製程名稱'],
      ['gmoProc', 'Glass 的製程名稱 1'],
      ['filterProc', 'Glass 的製程名稱 2']
    ].map(function (p) {
      var input = h('input', {
        id: 'f-' + p[0], type: 'text', class: 'in in-short', autocomplete: 'off', spellcheck: 'false', value: settings[p[0]]
      });
      els.procInputs[p[0]] = input;
      return h('div', { class: 'field-row' }, h('label', { for: 'f-' + p[0], text: p[1] }), input);
    });

    els.summary = h('p', { class: 'summary', 'aria-live': 'polite' });
    els.fileNotes = h('ul', { class: 'notes' });
    els.outPreview = h('div', { class: 'sheet-area' });
    els.exportBtn = h('button', { type: 'button', class: 'btn-primary', disabled: true, text: '匯出 Excel', onclick: onExport });
    els.nextBtn = h('button', {
      type: 'button', class: 'btn-secondary', disabled: true, text: '下一步：匯總',
      onclick: function () { IEApp.go('mps-totals'); }
    });
    els.exportNote = h('span', { class: 'export-note', 'aria-live': 'polite' });

    panel.appendChild(h('div', { class: 'step-head' },
      h('h1', { text: '篩選 MPS 報表' }),
      h('p', { text: '可以上傳多個 MPS 報表，每個檔案各自選篩選方式和要保留的行，把列分成 Plastic、Glass 兩段，匯出成一個新的 Excel。' +
        '表頭（第 1–4 列）照第 1 個檔案保留，下面先放 Plastic、空 10 列，再放 Glass，每一段上面有標題；同一段裡依檔案順序接下去。' })));

    els.exportTask = h('li', { class: 'task' },
      h('h2', { text: '匯出' }),
      els.summary, els.fileNotes, els.outPreview,
      h('div', { class: 'actions' }, els.exportBtn, els.nextBtn, els.exportNote),
      h('details', { class: 'advanced' },
        h('summary', { text: '製程名稱不一樣時' }),
        procRows,
        h('p', { class: 'hint', text: '製程名稱就是 TTL ( ) 括號裡的字，也會用在匯出檔案的標題。比對時不分大小寫，空白和橫線有沒有都可以。' })));

    els.tasks = h('ol', { class: 'tasks' },
      h('li', { class: 'task' },
        h('h2', { text: '上傳 MPS 報表' }),
        els.drop.el, els.fileList, els.fileMsg,
        h('p', { class: 'hint', text: '可以一次選多個檔案，也可以分幾次加入。匯出時同一段裡的列照這裡的順序排，可以用「往上移」調整。' +
          '每個檔案的設定照順序記住：下個月第 1 個檔案會用上次第 1 個檔案的設定。' })));
    panel.appendChild(els.tasks);

    Object.keys(els.procInputs).forEach(function (k) {
      els.procInputs[k].addEventListener('input', function () {
        settings[k] = els.procInputs[k].value;
        S.files.forEach(function (f) { f.blocksKey = ''; });
        schedule();
      });
    });

    renderFiles();
  }

  function schedule() {
    clearTimeout(timer);
    timer = setTimeout(recompute, 200);
  }

  /* ---------- 上傳的檔案清單、每個檔案一塊設定 ---------- */
  function renderFiles() {
    els.fileList.textContent = '';
    S.files.forEach(function (f, i) {
      els.fileList.appendChild(h('li', { class: 'file-item' },
        h('span', { class: 'file-no', text: '檔案 ' + (i + 1) }),
        h('span', { class: 'file-name', text: f.name }),
        i > 0 ? h('button', { type: 'button', class: 'btn-link', text: '往上移', onclick: function () { moveFile(i); } }) : null,
        h('button', { type: 'button', class: 'btn-link', text: '移除', onclick: function () { removeFile(i); } })));
    });
    if (S.files.length) els.drop.setFile('再加入檔案', '已經上傳 ' + S.files.length + ' 個檔案。點這裡或拖檔案進來可以再加。');
    else els.drop.setIdle();

    // 每個檔案一個 task，放在「上傳」和「匯出」之間
    while (els.tasks.children.length > 1) els.tasks.removeChild(els.tasks.lastChild);
    els.files = [];
    S.files.forEach(function (f, i) {
      var task = buildFileTask(f, i);
      els.tasks.appendChild(task);
    });
    els.tasks.appendChild(els.exportTask);
    recompute();
  }

  function buildFileTask(f, i) {
    var fs = fset(i);
    var E = {};
    els.files[i] = E;
    var id = 'f' + i + '-';

    E.sheetSelect = h('select', { id: id + 'sheet', 'aria-label': '工作表' });
    f.workbook.worksheets.forEach(function (ws, k) {
      E.sheetSelect.appendChild(h('option', { value: String(k), text: ws.name }));
    });
    E.sheetSelect.value = String(f.workbook.worksheets.indexOf(f.ws));
    E.sheetSelect.addEventListener('change', function () {
      f.ws = f.workbook.worksheets[parseInt(E.sheetSelect.value, 10)];
      f.infoKey = ''; f.blocksKey = '';
      recompute();
    });
    var sheetRow = h('div', { class: 'field-row', hidden: f.workbook.worksheets.length < 2 },
      h('label', { for: id + 'sheet', text: '工作表' }), E.sheetSelect);

    // 篩選方式（三選一）
    E.radios = MODES.map(function (m) {
      var r = h('input', { type: 'radio', name: id + 'mode', id: id + 'mode-' + m.value, value: m.value });
      r.checked = fs.mode === m.value;
      r.addEventListener('change', function () { if (r.checked) { fs.mode = m.value; recompute(); } });
      return h('label', { class: 'mode-option', for: id + 'mode-' + m.value }, r,
        h('span', null, h('strong', { text: m.title }), h('span', { class: 'mode-sub', text: m.sub })));
    });

    // 依製程（兩種共用：料號在哪一行）
    E.pnCol = textInput(id + 'pncol', 'in in-tiny', fs.pnCol, fs, 'pnCol', true);
    E.procHint = h('p', { class: 'hint' });
    E.procChips = h('div', { class: 'chips' });
    E.procPane = h('div', { class: 'mode-pane' },
      h('div', { class: 'field-row' },
        h('label', { for: id + 'pncol', text: '料號（P/N）在' }), E.pnCol, h('span', { text: '行' })),
      E.procHint, E.procChips);

    // 依 Status
    E.selLabels = [];
    var statusRows = SECTIONS.map(function (sec) {
      var lab = h('span', { text: 'A' });
      E.selLabels.push(lab);
      var input = textInput(id + 'val-' + sec.key, 'in in-short', fs[sec.valueSetting], fs, sec.valueSetting, false);
      input.placeholder = '例如 1';
      return h('div', { class: 'field-row' },
        h('label', { for: id + 'val-' + sec.key, class: 'label-fixed' }, sec.label + '：', lab, ' 行等於'), input);
    });
    E.statusChips = h('div', { class: 'chips' });
    E.statusPane = h('div', { class: 'mode-pane' },
      statusRows,
      h('p', { class: 'hint', text: '一種有多個數值時用逗號分開，例如 1, 3。同一個數值不能同時算 Plastic 和 Glass。這個檔案只有其中一種時，另一個留空白。' }),
      E.statusChips);

    E.cols = textInput(id + 'cols', 'in in-wide', fs.cols, fs, 'cols', false);
    E.cols.placeholder = '例如 B, C, U, Z, AE';
    E.cols.spellcheck = false;
    E.srcPreview = h('div', { class: 'sheet-area' });
    E.msg = h('p', { class: 'summary', 'aria-live': 'polite' });

    var startRow = textInput(id + 'start', 'in in-tiny', fs.startRow, fs, 'startRow', true);
    startRow.inputMode = 'numeric';
    var selCol = textInput(id + 'selcol', 'in in-tiny', fs.selCol, fs, 'selCol', true);

    return h('li', { class: 'task' },
      h('h2', null, '檔案 ' + (i + 1) + '：', h('span', { class: 'file-name', text: f.name })),
      sheetRow,
      h('h3', { class: 'sub-head', text: '篩選方式' }),
      h('div', { class: 'mode-pick', role: 'radiogroup', 'aria-label': '檔案 ' + (i + 1) + ' 的篩選方式' }, E.radios),
      E.procPane, E.statusPane,
      h('h3', { class: 'sub-head', text: '要保留的行' }),
      h('div', { class: 'field-row' }, h('label', { for: id + 'cols', text: '行號' }), E.cols),
      h('p', { class: 'hint', text: '連續的行可以寫成 U:Z，也可以直接點下面表格最上方的行號來選。' +
        '每個檔案保留的行數要一樣，順序也要對得上（第 1 個保留的行都放料號…），位置不同沒關係。' }),
      E.srcPreview, E.msg,
      h('details', { class: 'advanced' },
        h('summary', { text: '報表版面不一樣時' }),
        h('div', { class: 'field-row' },
          h('label', { for: id + 'start', text: '資料從第' }), startRow, h('span', { text: '列開始（之前的列都當作表頭；匯出的表頭用第 1 個檔案的）' })),
        h('div', { class: 'field-row' },
          h('label', { for: id + 'selcol', text: 'Status 在' }), selCol, h('span', { text: '行（依 Status 篩選時用）' }))));
  }

  // 輸入框：打字就存進這個檔案的設定，稍後重算。resetCache = 會影響讀檔結果（要重新掃描）
  function textInput(idv, cls, value, fs, key, resetCache) {
    var input = h('input', { id: idv, type: 'text', class: cls, autocomplete: 'off', value: String(value) });
    input.addEventListener('input', function () {
      fs[key] = input.value;
      if (resetCache) S.files.forEach(function (f) { f.infoKey = ''; f.blocksKey = ''; });
      schedule();
    });
    return input;
  }

  /* ---------- 開檔、移除、排順序 ---------- */
  function openFiles(files) {
    els.fileMsg.hidden = true;
    els.drop.setBusy(files.map(function (f) { return f.name; }).join('、'));
    var errors = [];
    files.reduce(function (p, file) {
      return p.then(function () {
        return ui.readWorkbook(file).then(function (res) {
          S.files.push({ name: file.name, workbook: res.workbook, ws: res.workbook.worksheets[0], infoKey: '', blocksKey: '' });
        }).catch(function (err) {
          errors.push(file.name + '：' + (err && err.message ? err.message : String(err)));
        });
      });
    }, Promise.resolve()).then(function () {
      els.exportNote.textContent = '';
      if (errors.length) {
        els.fileMsg.textContent = errors.join(' ');
        els.fileMsg.hidden = false;
      }
      renderFiles();
    });
  }

  function removeFile(i) {
    S.files.splice(i, 1);
    // 這個檔案的設定移到最後面（不丟掉），後面的檔案往前補，各自的設定跟著走
    fset(i);
    settings.files.push(settings.files.splice(i, 1)[0]);
    els.exportNote.textContent = '';
    renderFiles();
  }

  function moveFile(i) {
    if (i < 1) return;
    var f = S.files.splice(i, 1)[0];
    S.files.splice(i - 1, 0, f);
    fset(i);
    var s = settings.files.splice(i, 1)[0];
    settings.files.splice(i - 1, 0, s);
    els.exportNote.textContent = '';
    renderFiles();
  }

  /* ---------- 一個檔案：找出 Plastic、Glass 各是哪幾列 ---------- */

  // 依製程：mode 是 plastic 或 glass。回傳 { rows: { plastic: [...], glass: [...] }, warn } 或 { error }
  function rowsByProcess(f, fs, E, sr) {
    var pc = pnColOf(fs);
    var key = sr + ':' + pc + ':' + f.info.lastRow;
    if (f.blocksKey !== key) {
      f.blocks = XIO.detectProcessBlocks(f.ws, { dataStartRow: sr, keyCol: pc, lastRow: f.info.lastRow });
      f.blocksKey = key;
    }
    var blocks = f.blocks.blocks;
    var sec = sectionOf(MODE_SECTION[fs.mode]);
    var rows = { plastic: [], glass: [] };
    var used = {};
    var missing = [];
    sec.procs.forEach(function (setting) {
      var want = XIO.processKey(procName(setting));
      var hit = false;
      blocks.forEach(function (b, i) {
        if (b.key !== want) return;
        hit = true;
        used[i] = sec.label;
        rows[sec.key] = rows[sec.key].concat(b.rows);
      });
      if (!hit) missing.push('TTL (' + procName(setting) + ')');
    });
    renderProcessChips(E, blocks, used);

    if (!blocks.length) {
      return { error: XIO.numToCol(pc) + ' 行找不到「TTL (製程名)」的列。請確認料號在哪一行（上面的「料號（P/N）在」），或改用別的篩選方式。' };
    }
    if (missing.length === sec.procs.length) {
      return {
        error: '找不到 ' + missing.join('、') +
          '。這份報表有的製程列在上面；名稱不一樣的話，到「匯出」下面的「製程名稱不一樣時」改製程名稱。'
      };
    }
    return { rows: rows, warn: missing.length ? '找不到 ' + missing.join('、') + '，只用找到的那一塊。' : '' };
  }

  // 依 Status：回傳 { rows, wanted } 或 { error, wanted }
  function rowsByStatus(f, fs, E, sr) {
    var sc = selColOf(fs);
    var wanted = {};
    var conflict = [];
    var lists = {};
    SECTIONS.forEach(function (sec) {
      lists[sec.key] = valueList(fs[sec.valueSetting]);
      lists[sec.key].forEach(function (v) {
        var k = XIO.matchKey(v);
        if (wanted[k] && wanted[k] !== sec.key) conflict.push(v);
        wanted[k] = sec.key;
      });
    });
    renderStatusChips(E, f.info, wanted, sc);

    var colName = XIO.numToCol(sc);
    if (!lists.plastic.length && !lists.glass.length) {
      return { wanted: wanted, prompt: true, error: '輸入 ' + colName + ' 行等於多少代表 Plastic、Glass（這個檔案只有一種時，另一個留空白）。' };
    }
    if (conflict.length) {
      return { wanted: wanted, error: conflict.join('、') + ' 同時寫在 Plastic 和 Glass，請只留一邊。' };
    }
    var rows = { plastic: [], glass: [] };
    for (var r = sr; r <= f.info.lastRow; r++) {
      var row = f.ws.findRow(r);
      if (!row) continue;
      var cell = row.findCell(sc);
      var k = XIO.matchKey(cell ? (cell.isMerged ? cell.master.value : cell.value) : null);
      if (k !== '' && wanted[k]) rows[wanted[k]].push(r);
    }
    return { rows: rows, wanted: wanted };
  }

  // 算一個檔案，更新它那一塊的畫面。回傳 { ok, rows, keepCols, count } 或 { ok: false }
  function computeFile(f, i) {
    var fs = fset(i);
    var E = els.files[i];
    var byStatus = fs.mode === 'status';
    E.procPane.hidden = byStatus;
    E.statusPane.hidden = !byStatus;
    var sc = selColOf(fs);
    var scName = XIO.numToCol(sc);
    E.selLabels.forEach(function (l) { l.textContent = scName; });
    E.procChips.textContent = '';
    E.statusChips.textContent = '';
    if (!byStatus) {
      var sec = sectionOf(MODE_SECTION[fs.mode]);
      E.procHint.textContent = '程式會在這一行找「TTL (製程名)」：' +
        sec.procs.map(function (p) { return 'TTL (' + procName(p) + ')'; }).join('、') +
        ' 上面那' + (sec.procs.length > 1 ? '幾塊' : '一塊') + '算 ' + sec.label + '。「x」和 TTL 這兩列不會留下。';
    }

    var sr = startRowOf(fs);
    var key = sr + ':' + sc;
    if (f.infoKey !== key) {
      f.info = XIO.inspect(f.ws, { dataStartRow: sr, selectorCol: sc });
      f.infoKey = key;
    }
    var info = f.info;
    var parsed = XIO.parseColumns(fs.cols);
    var keep = {};
    parsed.cols.forEach(function (c) { keep[c] = true; });

    var picked = !info.dataRows ? {} : byStatus ? rowsByStatus(f, fs, E, sr) : rowsByProcess(f, fs, E, sr);

    var scroller = E.srcPreview.querySelector('.sheet-scroll');
    var left = scroller ? scroller.scrollLeft : 0;
    ui.renderSheet(E.srcPreview, f.ws, {
      headerEnd: sr - 1,
      lastRow: info.lastRow,
      colCount: Math.max(info.colCount, 1),
      rows: PREVIEW_ROWS,
      keep: keep,
      onToggle: function (c) { toggleColumn(i, c); },
      selectorCol: sc,
      wanted: byStatus && picked.rows ? picked.wanted : null,
      caption: '原始報表（前 ' + PREVIEW_ROWS + ' 列資料）。共 ' + info.dataRows.toLocaleString('en-US') +
        ' 列資料、' + info.colCount + ' 行（到 ' + XIO.numToCol(info.colCount || 1) + ' 行）。'
    });
    scroller = E.srcPreview.querySelector('.sheet-scroll');
    if (scroller) scroller.scrollLeft = left;

    function msg(text, isError, isWarn) {
      E.msg.textContent = text;
      E.msg.className = 'summary' + (isError ? ' is-error' : '') + (isWarn ? ' is-warn' : '');
    }

    if (!info.dataRows) { msg('第 ' + sr + ' 列之後沒有資料。請確認上傳的檔案和工作表。', true); return { ok: false }; }
    if (picked.error) { msg(picked.error, !picked.prompt); return { ok: false }; }
    if (parsed.bad.length) {
      msg('看不懂「' + parsed.bad.join('、') + '」。行號請用英文字母，例如 B 或 U:Z。', true);
      return { ok: false };
    }
    if (!parsed.cols.length) { msg('輸入要保留的行。'); return { ok: false }; }

    var counts = { plastic: picked.rows.plastic.length, glass: picked.rows.glass.length };
    if (!counts.plastic && !counts.glass) {
      var seen = info.values.slice(0, 8).map(function (v) { return v.label; }).join('、');
      msg(!byStatus ? sectionOf(MODE_SECTION[fs.mode]).label + ' 的製程底下沒有料號。' :
        scName + ' 行沒有符合的列。' + (seen ? '這份報表的 ' + scName + ' 行有：' + seen + '。' : scName + ' 行從第 ' + sr + ' 列起都是空白。'), true);
      return { ok: false };
    }

    var warn = [];
    if (picked.warn) warn.push(picked.warn);
    var beyond = parsed.cols.filter(function (c) { return c > info.colCount; });
    if (beyond.length) warn.push(XIO.formatColumns(beyond) + ' 行在報表裡沒有資料。');
    msg(SECTIONS.filter(function (s) { return counts[s.key]; })
      .map(function (s) { return s.label + ' ' + counts[s.key].toLocaleString('en-US') + ' 列'; }).join('、') +
      '，保留 ' + parsed.cols.length + ' 行（' + XIO.formatColumns(parsed.cols) + '）。' +
      (warn.length ? ' 注意：' + warn.join(' ') : ''), false, warn.length > 0);

    return { ok: true, rows: picked.rows, keepCols: parsed.cols, counts: counts, dataRows: info.dataRows, startRow: sr };
  }

  /* ---------- 重新計算並更新畫面 ---------- */
  function recompute() {
    saveSettings();
    S.result = null;
    IEApp.state.mpsFiltered = null;
    els.exportBtn.disabled = true;
    els.nextBtn.disabled = true;
    els.outPreview.textContent = '';
    els.fileNotes.textContent = '';

    if (!S.files.length) { setSummary('先上傳 MPS 報表。'); return; }

    var res = S.files.map(computeFile);
    var bad = [];
    res.forEach(function (r, i) { if (!r.ok) bad.push(i + 1); });
    if (bad.length) {
      setSummary('檔案 ' + bad.join('、') + ' 還沒設定好，請看那個檔案下面的說明。', false, true);
      return;
    }

    // 每個檔案保留的行數要一樣（第 1 個檔案的表頭、行寬套用到全部）
    var n = res[0].keepCols.length;
    var diff = [];
    res.forEach(function (r, i) { if (r.keepCols.length !== n) diff.push('檔案 ' + (i + 1) + ' 保留 ' + r.keepCols.length + ' 行'); });
    if (diff.length) {
      setSummary('每個檔案保留的行數要一樣：檔案 1 保留 ' + n + ' 行，' + diff.join('、') + '。', true);
      return;
    }

    // 兩段的列：依檔案順序接下去
    var first = S.files[0];
    var sections = SECTIONS.map(function (sec) {
      var rows = [];
      res.forEach(function (r, i) {
        var f = S.files[i];
        r.rows[sec.key].forEach(function (row) {
          rows.push(i === 0 ? row : { ws: f.ws, r: row, cols: r.keepCols });
        });
      });
      return { key: sec.key, title: sectionTitle(sec), rows: rows };
    });
    var opts = {
      dataStartRow: res[0].startRow, keepCols: res[0].keepCols, lastRow: first.info.lastRow,
      gap: XIO.SECTION_GAP, sections: sections
    };
    var ws = first.ws;
    var result = XIO.filter(ws, opts);
    S.result = result;
    var totalData = res.reduce(function (a, r) { return a + r.dataRows; }, 0);

    var text = SECTIONS.map(function (sec, k) {
      return sec.label + ' ' + sections[k].rows.length.toLocaleString('en-US') + ' 列';
    }).join('、') + '（' + S.files.length + ' 個檔案共 ' + totalData.toLocaleString('en-US') + ' 列資料），保留 ' + n + ' 行。' +
      result.sections.map(function (s) {
        return s.title + '：標題在第 ' + s.titleRow + ' 列' + (s.count ? '，資料在第 ' + s.firstRow + '–' + s.lastRow + ' 列' : '');
      }).join('；') + '。';
    var warn = [];
    sections.forEach(function (s, k) { if (!s.rows.length) warn.push(SECTIONS[k].label + ' 沒有任何列，只會有標題。'); });
    if (warn.length) text += ' 注意：' + warn.join(' ');
    setSummary(text, false, warn.length > 0);

    if (S.files.length > 1) {
      res.forEach(function (r, i) {
        els.fileNotes.appendChild(h('li', null, '檔案 ' + (i + 1) + '（' + S.files[i].name + '）：' +
          SECTIONS.map(function (s) { return s.label + ' ' + r.counts[s.key] + ' 列'; }).join('、')));
      });
    }

    ui.renderSheet(els.outPreview, result.worksheet, {
      headerEnd: result.headerEnd,
      lastRow: result.lastRow,
      colCount: result.colCount,
      rows: Math.min(result.lastRow - result.headerEnd, OUT_PREVIEW_MAX),
      caption: '匯出後的樣子（表頭是檔案 1 的）。'
    });
    els.exportBtn.disabled = false;
    els.nextBtn.disabled = false;

    // 給第 2 步用。build() 每次都產生一份全新的活頁簿，後面的步驟可以放心往裡面加工作表。
    IEApp.state.mpsFiltered = {
      sourceName: S.files.map(function (f) { return f.name; }).join('、'),
      outName: outputName(),
      sheetName: result.worksheet.name,
      matchedRows: result.matchedRows,
      headerEnd: result.headerEnd,
      sections: result.sections,      // [{ key, title, titleRow, firstRow, lastRow, count }]
      build: function () { return XIO.filter(ws, opts).workbook; }
    };
  }

  function setSummary(text, isError, isWarn) {
    els.summary.textContent = text;
    els.summary.className = 'summary' + (isError ? ' is-error' : '') + (isWarn ? ' is-warn' : '');
  }

  /* ---------- 報表裡找到的製程（標出用的是哪幾塊） ---------- */
  function renderProcessChips(E, blocks, used) {
    E.procChips.textContent = '';
    if (!blocks.length) return;
    E.procChips.appendChild(h('span', { class: 'chips-label', text: '報表裡的製程' }));
    blocks.forEach(function (b, i) {
      var tag = used[i];
      E.procChips.appendChild(h('span', {
        class: 'chip is-static' + (tag ? ' is-on' : ''),
        title: 'TTL 在第 ' + b.ttlRow + ' 列'
      }, tag ? h('span', { class: 'chip-tag', text: tag }) : null,
         h('span', { class: 'chip-value', text: b.name }),
         h('span', { class: 'chip-count', text: b.rows.length.toLocaleString('en-US') + ' 列' })));
    });
  }

  /* ---------- Status 行出現過的值（標出算 Plastic 還是 Glass） ---------- */
  function renderStatusChips(E, info, wanted, sc) {
    E.statusChips.textContent = '';
    if (!info.values.length && !info.blankCount) return;
    var col = XIO.numToCol(sc);
    var labelOf = {};
    SECTIONS.forEach(function (sec) { labelOf[sec.key] = sec.label; });
    E.statusChips.appendChild(h('span', { class: 'chips-label', text: '報表裡 ' + col + ' 行的值' }));
    info.values.slice(0, MAX_CHIPS).forEach(function (v) {
      var tag = wanted[v.key] ? labelOf[wanted[v.key]] : '';
      E.statusChips.appendChild(h('span', { class: 'chip is-static' + (tag ? ' is-on' : '') },
        tag ? h('span', { class: 'chip-tag', text: tag }) : null,
        h('span', { class: 'chip-value', text: v.label }),
        h('span', { class: 'chip-count', text: v.count.toLocaleString('en-US') + ' 列' })));
    });
    if (info.values.length > MAX_CHIPS) {
      E.statusChips.appendChild(h('span', { class: 'chips-more', text: '另有 ' + (info.values.length - MAX_CHIPS) + ' 種' }));
    }
    if (info.blankCount) {
      E.statusChips.appendChild(h('span', { class: 'chips-more', text: '空白 ' + info.blankCount.toLocaleString('en-US') + ' 列' }));
    }
  }

  function toggleColumn(i, c) {
    var fs = fset(i);
    var cols = XIO.parseColumns(fs.cols).cols;
    var k = cols.indexOf(c);
    if (k >= 0) cols.splice(k, 1); else cols.push(c);
    cols.sort(function (a, b) { return a - b; });
    fs.cols = XIO.formatColumns(cols);
    els.files[i].cols.value = fs.cols;
    recompute();
    var btn = els.files[i].srcPreview.querySelector('button[data-col="' + c + '"]');
    if (btn) btn.focus({ preventScroll: true });
  }

  /* ---------- 匯出 ---------- */
  function onExport() {
    var ready = IEApp.state.mpsFiltered;
    if (!S.result || !ready) return;
    var rows = ready.matchedRows;
    var name = ready.outName;
    els.exportBtn.disabled = true;
    els.exportNote.textContent = '產生檔案中…';
    XIO.toBlob(ready.build()).then(function (blob) {
      ui.download(blob, name);
      els.exportNote.textContent = '已匯出 ' + name + '（' + rows.toLocaleString('en-US') + ' 列），在瀏覽器的下載資料夾。';
    }).catch(function (err) {
      els.exportNote.textContent = '匯出失敗：' + (err && err.message ? err.message : err);
    }).then(function () {
      els.exportBtn.disabled = false;
    });
  }

  IEApp.registerStep({ id: 'mps-filter', render: render });
})(window);
