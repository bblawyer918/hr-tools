/* 出勤紀錄標準格式的解析與判斷（與 attendance/出勤管理儀表板.html 同一套規則）。需要 xlsx 與 common.js */
var ATT = (function () {
  var RULES = { '84-1': { dailyNormal: 10, dailyMax: 12, monthNormal: 240 }, '一般': { dailyNormal: 8, dailyMax: 12, monthNormal: null } };
  var DEFAULT_SHIFTS = {
    '日班12H': { start: 420, end: 1140, rest: 0, normal: 10, rule: '84-1', grace: 5 },
    '夜班12H': { start: 1140, end: 420, rest: 0, normal: 10, rule: '84-1', grace: 5 },
    '機動12H': { start: 420, end: 1140, rest: 0, normal: 10, rule: '84-1', grace: 5 },
    '辦公室9H': { start: 480, end: 1020, rest: 60, normal: 8, rule: '一般', grace: 5 }
  };
  var LEAVE_NAMES = ['特休', '事假', '病假', '婚假', '喪假', '公假', '公傷病假', '產假', '陪產檢及陪產假', '生理假', '家庭照顧假', '產檢假', '安胎休養假', '補休'];
  var SKIP = { '填寫說明': 1, '班別設定': 1, '假別': 1, '班表': 1, '人力檢核': 1, '問題清單': 1 };
  var txt = HR.txt, num = HR.num;

  function toMin(v) {
    if (v === null || v === undefined || v === '') return null;
    if (typeof v === 'number') {
      if (v >= 0 && v < 1) return Math.round(v * 1440);
      if (v >= 100 && v < 2400 && v % 100 < 60) return Math.floor(v / 100) * 60 + v % 100;
      return null;
    }
    var s = String(v).trim().replace('：', ':');
    var m = s.match(/^(\d{1,2}):(\d{2})/) || s.match(/^(\d{2})(\d{2})$/);
    if (!m) return null;
    var h = +m[1], mi = +m[2];
    return h < 24 && mi < 60 ? h * 60 + mi : null;
  }
  function hhmm(min) { min = ((min % 1440) + 1440) % 1440; return ('0' + Math.floor(min / 60)).slice(-2) + ':' + ('0' + min % 60).slice(-2); }
  function parseLeave(v) {
    var s = txt(v); if (!s) return [];
    if (num(v) !== null && /^[\d.]+$/.test(s)) return [{ type: '未寫假別', h: num(v) }];
    return s.split(/[、,，;；\/\s]+(?=\D)/).map(function (part) {
      var m = part.match(/^([^\d.]+?)\s*([\d.]+)\s*(h|小時|hr)?$/i);
      if (!m) return { type: part, h: null };
      var t = m[1].trim(), known = LEAVE_NAMES.filter(function (n) { return n === t || n.indexOf(t) === 0 || t.indexOf(n) === 0; })[0];
      return { type: known || t, h: +m[2] };
    });
  }
  function readShifts(wb) {
    var shifts = JSON.parse(JSON.stringify(DEFAULT_SHIFTS));
    if (wb.SheetNames.indexOf('班別設定') < 0) return { shifts: shifts, source: '內建預設' };
    var got = {};
    HR.rows(wb, '班別設定').slice(1).forEach(function (r) {
      var name = txt(r[0]), st = toMin(r[1]), en = toMin(r[2]);
      if (!name || st === null || en === null) return;
      got[name] = { start: st, end: en, rest: num(r[3]) || 0, normal: num(r[4]) || 8, rule: RULES[txt(r[5])] ? txt(r[5]) : (/84/.test(txt(r[5])) ? '84-1' : '一般'), grace: num(r[6]) || 0, code: txt(r[7]) };
    });
    return Object.keys(got).length ? { shifts: got, source: '檔案內「班別設定」' } : { shifts: shifts, source: '內建預設' };
  }
  function monthSheets(wb) {
    var ms = wb.SheetNames.filter(function (n) { return !SKIP[n]; });
    ms.sort(function (a, b) { return (/\d+年\d+月/.test(b) ? 1 : 0) - (/\d+年\d+月/.test(a) ? 1 : 0); });
    return ms;
  }
  // 每人一組列：第一列標「社區」，之後是 上班／下班／加班／請假（舊格式是 出勤／請假）
  function parse(rows) {
    var hdr = -1, cTitle = -1;
    for (var r = 0; r < Math.min(rows.length, 30) && hdr < 0; r++) {
      var row = rows[r] || [];
      for (var c = 0; c < row.length; c++) if (txt(row[c]) === '職稱') { hdr = r; cTitle = c; break; }
    }
    if (hdr < 0) throw new Error('找不到「職稱」標題列，這份可能不是出勤紀錄。');
    var head = rows[hdr], cShift = cTitle + 1, cName = cTitle + 2, cLabel = cTitle + 3, dayCol = {};
    head.forEach(function (v, c) { var n = num(v); if (c > cLabel && n !== null && n >= 1 && n <= 31 && n === Math.floor(n)) dayCol[n] = c; });
    var ym = null;
    for (var r2 = 0; r2 <= hdr; r2++) (rows[r2] || []).forEach(function (v) { var m = txt(v).match(/^(\d{2,3})\s*[\/年]\s*(\d{1,2})/); if (m && !ym) ym = { y: +m[1] + 1911, m: +m[2] }; });
    if (!ym) throw new Error('找不到月份（例如「115/9」）。');
    var nDays = new Date(ym.y, ym.m, 0).getDate(), people = [], cur = null, mode = 'old';
    for (var i = hdr + 1; i < rows.length; i++) {
      var rw = rows[i] || [], label = txt(rw[cLabel]);
      if (/實際每日上班時數/.test(txt(rw[cTitle]))) break;
      if (label === '社區' && txt(rw[cName])) {
        cur = { title: txt(rw[cTitle]), shift: txt(rw[cShift]), name: txt(rw[cName]), days: {} };
        for (var d = 1; d <= nDays; d++) cur.days[d] = { site: dayCol[d] !== undefined ? txt(rw[dayCol[d]]) : '', in: null, out: null, ot: null, hours: null, code: '', leave: [] };
        people.push(cur); continue;
      }
      if (!cur) continue;
      for (var d2 = 1; d2 <= nDays; d2++) {
        if (dayCol[d2] === undefined) continue;
        var v = rw[dayCol[d2]], x = cur.days[d2];
        if (label === '上班') { mode = 'new'; x.in = toMin(v); if (x.in === null && txt(v)) x.code = txt(v); }
        else if (label === '下班') { x.out = toMin(v); if (x.out === null && txt(v)) x.code = txt(v); }
        else if (label === '加班') x.ot = num(v);
        else if (label === '請假') x.leave = parseLeave(v);
        else if (label === '出勤') { var n = num(v); if (n !== null) x.hours = n; else if (txt(v)) x.code = txt(v); }
      }
    }
    if (!people.length) throw new Error('找到標題列，但沒有讀到任何人員。');
    return { ym: ym, nDays: nDays, people: people, mode: mode };
  }
  // 單日判斷（NUEIP 式：遲到／早退／缺卡／曠職／超時未報加班）。回傳 [[類型, 說明], ...]
  function judgeDay(data, d, x, sh, opts) {
    var issues = [], gap = opts && opts.otGap !== undefined ? opts.otGap : 30;
    var leaveH = x.leave.reduce(function (a, l) { return a + (l.h || 0); }, 0);
    x.worked = 0; x.isWork = false; x.late = 0; x.early = 0;
    if (x.hours !== null || (x.in === null && x.out === null && x.code && data.mode === 'old')) {
      x.worked = x.hours || 0; x.isWork = x.worked > 0 || !!x.code;
      if (x.code) issues.push(['無法判讀', d + ' 日出勤欄寫「' + x.code + '」']);
    } else if (x.in !== null || x.out !== null) {
      x.isWork = true;
      if (x.in === null || x.out === null) issues.push(['缺卡', d + ' 日缺' + (x.in === null ? '上班' : '下班') + '紀錄']);
      if (sh && x.in !== null) {
        var late = x.in - sh.start; if (late < -720) late += 1440; if (late > 720) late -= 1440;
        if (late > sh.grace && leaveH === 0) { x.late = late; issues.push(['遲到', d + ' 日遲到 ' + late + ' 分（' + hhmm(x.in) + '）']); }
      }
      if (x.in !== null && x.out !== null) {
        var span = x.out - x.in; if (span <= 0) span += 1440;
        if (sh) {
          var L = ((sh.end - sh.start) + 1440) % 1440 || 1440, inRel = ((x.in - sh.start) + 1440) % 1440; if (inRel > 720) inRel -= 1440;
          span = Math.max(0, Math.min(inRel + span, L) - Math.max(inRel, 0));
        }
        x.worked = Math.max(0, span - (sh ? sh.rest : 0)) / 60 + (x.ot || 0);
        if (sh) {
          var early = (sh.end - x.out + 1440) % 1440; if (early > 720) early -= 1440;
          if (early > 0 && leaveH === 0) { x.early = early; issues.push(['早退', d + ' 日早退 ' + early + ' 分（' + hhmm(x.out) + '）']); }
          if (-early > gap && !(x.ot > 0)) issues.push(['超時未報加班', d + ' 日晚 ' + (-early) + ' 分下班，加班欄空白']);
        }
      }
    } else if (x.code) {
      x.isWork = true; issues.push(['無法判讀', d + ' 日寫「' + x.code + '」']);
    } else if (x.site && leaveH === 0) {
      issues.push(['曠職', d + ' 日有排班（' + x.site + '）但沒有上下班紀錄，也沒有請假']);
    }
    return issues;
  }
  return { RULES: RULES, LEAVE_NAMES: LEAVE_NAMES, toMin: toMin, hhmm: hhmm, parseLeave: parseLeave, readShifts: readShifts, monthSheets: monthSheets, parse: parse, judgeDay: judgeDay };
})();
