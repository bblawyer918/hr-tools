/* 人資工具共用函式：讀 Excel、日期、匯出、複製。需要先載入 xlsx.full.min.js */
var HR = (function () {
  // 國定假日（紀念日及節日實施條例第 4、6 條；實際以人事行政總處公告為準）
  var HOLIDAYS = {
    2026: { '1-1': '開國紀念日', '2-15': '小年夜', '2-16': '除夕', '2-17': '春節', '2-18': '春節', '2-19': '春節',
      '2-28': '和平紀念日', '4-4': '兒童節', '4-5': '民族掃墓節', '5-1': '勞動節', '6-19': '端午節',
      '9-25': '中秋節', '9-28': '教師節', '10-10': '國慶日', '10-25': '光復紀念日', '12-25': '行憲紀念日' },
    2027: { '1-1': '開國紀念日', '2-4': '小年夜', '2-5': '除夕', '2-6': '春節', '2-7': '春節', '2-8': '春節',
      '2-28': '和平紀念日', '4-4': '兒童節', '4-5': '民族掃墓節', '5-1': '勞動節', '6-9': '端午節',
      '9-15': '中秋節', '9-28': '教師節', '10-10': '國慶日', '10-25': '光復紀念日', '12-25': '行憲紀念日' }
  };
  var WK = '日一二三四五六';

  function $(id) { return document.getElementById(id); }
  function txt(v) { return v === null || v === undefined ? '' : String(v).trim(); }
  function num(v) {
    if (v === null || v === undefined || v === '') return null;
    var n = typeof v === 'number' ? v : parseFloat(String(v).replace(/[,\s元]/g, ''));
    return isNaN(n) ? null : n;
  }
  function esc(s) { return txt(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
  function fmt(n, digits) { var p = Math.pow(10, digits === undefined ? 1 : digits); return (Math.round(n * p) / p).toLocaleString('zh-TW'); }
  function pct(a, b) { return b ? fmt(a / b * 100) + '%' : '—'; }

  // 日期：Excel 日期序號、Date、「2026/3/5」「2026-03-05」「115/03/05」「115年3月5日」都可以
  function toDate(v) {
    if (v === null || v === undefined || v === '') return null;
    if (v instanceof Date) return isNaN(v) ? null : new Date(v.getFullYear(), v.getMonth(), v.getDate());
    if (typeof v === 'number') {
      if (v < 1000 || v > 80000) return null;
      var u = new Date(Math.round((v - 25569) * 86400000));
      return new Date(u.getUTCFullYear(), u.getUTCMonth(), u.getUTCDate());
    }
    var m = String(v).trim().match(/^(\d{2,4})\s*[\/\-.年]\s*(\d{1,2})\s*[\/\-.月]\s*(\d{1,2})/);
    if (!m) return null;
    var y = +m[1]; if (y < 1911) y += 1911;
    var d = new Date(y, +m[2] - 1, +m[3]);
    return d.getMonth() === +m[2] - 1 ? d : null;
  }
  function pad(n) { return ('0' + n).slice(-2); }
  function ymd(d) { return d ? d.getFullYear() + '/' + pad(d.getMonth() + 1) + '/' + pad(d.getDate()) : ''; }
  function roc(d) { return d ? (d.getFullYear() - 1911) + '/' + pad(d.getMonth() + 1) + '/' + pad(d.getDate()) : ''; }
  function iso(d) { return d ? d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()) : ''; }
  function addDays(d, n) { return new Date(d.getFullYear(), d.getMonth(), d.getDate() + n); }
  function addMonths(d, n) { var r = new Date(d.getFullYear(), d.getMonth() + n, d.getDate()); if (r.getDate() !== d.getDate()) r = new Date(r.getFullYear(), r.getMonth(), 0); return r; }
  function daysBetween(a, b) { return Math.round((new Date(b.getFullYear(), b.getMonth(), b.getDate()) - new Date(a.getFullYear(), a.getMonth(), a.getDate())) / 86400000); }
  function today() { var t = new Date(); return new Date(t.getFullYear(), t.getMonth(), t.getDate()); }
  function holiday(d) { return (HOLIDAYS[d.getFullYear()] || {})[(d.getMonth() + 1) + '-' + d.getDate()] || ''; }
  function isWorkday(d) { return d.getDay() !== 0 && d.getDay() !== 6 && !holiday(d); }
  function workdaysBetween(a, b) { var n = 0, d = addDays(a, 1); while (d <= b) { if (isWorkday(d)) n++; d = addDays(d, 1); } return n; }

  // 讀檔
  function readWorkbook(file, cb) {
    var fr = new FileReader();
    fr.onload = function () {
      try { cb(null, XLSX.read(new Uint8Array(fr.result), { type: 'array' })); }
      catch (e) { cb(new Error('讀不到這個檔案，請確認是 Excel（.xlsx／.xls）。')); }
    };
    fr.readAsArrayBuffer(file);
  }
  function rows(wb, name) { return XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: '' }); }
  // 找有指定欄位的標題列；aliases：{ key: ['欄名', '別名'] }
  function findTable(wb, aliases, required, preferSheets) {
    var names = (preferSheets || []).filter(function (n) { return wb.SheetNames.indexOf(n) >= 0; }).concat(wb.SheetNames);
    for (var s = 0; s < names.length; s++) {
      var rs = rows(wb, names[s]);
      for (var r = 0; r < Math.min(rs.length, 20); r++) {
        var head = rs[r].map(txt), col = {};
        Object.keys(aliases).forEach(function (k) {
          for (var i = 0; i < aliases[k].length; i++) { var c = head.indexOf(aliases[k][i]); if (c >= 0) { col[k] = c; break; } }
        });
        if (required.every(function (k) { return col[k] !== undefined; })) {
          var out = rs.slice(r + 1).map(function (row) { var o = {}; Object.keys(col).forEach(function (k) { o[k] = row[col[k]]; }); return o; })
            .filter(function (o) { return required.some(function (k) { return txt(o[k]) !== ''; }); });
          return { sheet: names[s], rows: out, cols: col };
        }
      }
    }
    return null;
  }
  function bindDrop(dropId, inputId, onFile, multiple) {
    var drop = $(dropId), input = $(inputId);
    if (multiple) input.multiple = true;
    drop.onclick = function () { input.click(); };
    drop.ondragover = function (e) { e.preventDefault(); drop.classList.add('over'); };
    drop.ondragleave = function () { drop.classList.remove('over'); };
    drop.ondrop = function (e) { e.preventDefault(); drop.classList.remove('over'); if (e.dataTransfer.files.length) onFile(multiple ? Array.from(e.dataTransfer.files) : e.dataTransfer.files[0]); };
    input.onchange = function (e) { if (e.target.files.length) onFile(multiple ? Array.from(e.target.files) : e.target.files[0]); e.target.value = ''; };
  }
  function exportXlsx(sheets, filename) {
    var wb = XLSX.utils.book_new();
    Object.keys(sheets).forEach(function (n) { XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(sheets[n]), n.slice(0, 31)); });
    XLSX.writeFile(wb, filename);
  }
  function copyText(text, done) {
    function fallback() { var ta = document.createElement('textarea'); ta.value = text; document.body.appendChild(ta); ta.select(); document.execCommand('copy'); ta.remove(); done && done(); }
    if (navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(text).then(function () { done && done(); }, fallback);
    else fallback();
  }
  function msg(el, text, err) { $(el).innerHTML = text ? '<div class="msg' + (err ? ' err' : '') + '">' + text + '</div>' : ''; }
  function store(key, val) {
    try { if (val === undefined) return JSON.parse(localStorage.getItem(key) || 'null'); localStorage.setItem(key, JSON.stringify(val)); } catch (e) { return null; }
  }

  return { HOLIDAYS: HOLIDAYS, WK: WK, $: $, txt: txt, num: num, esc: esc, fmt: fmt, pct: pct, toDate: toDate, ymd: ymd, roc: roc, iso: iso,
    addDays: addDays, addMonths: addMonths, daysBetween: daysBetween, today: today, holiday: holiday, isWorkday: isWorkday, workdaysBetween: workdaysBetween,
    readWorkbook: readWorkbook, rows: rows, findTable: findTable, bindDrop: bindDrop, exportXlsx: exportXlsx, copyText: copyText, msg: msg, store: store };
})();
