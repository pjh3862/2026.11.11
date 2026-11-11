/* 공정 재고·C/T 도구 — 순수 계산 로직 (화면·저장소 코드 없음)
   3차(2026-10-27) 범위 재설정 — 기획서 12장
   · 기준 자료 = 설정의 「관리 아이템」(고객사·그룹·상세그룹·품목·품번·비고 + 아이템마다 공정 순서)
   · 현재고(공정별 완료재고) = 기초재고(기준날짜 아침) + 기준날짜부터의 공정 실적 누계
       공정 P 를 수량 q 만큼 끝내면 → P 의 완료재고 + q, 바로 앞 공정의 완료재고 − q
       첫 공정은 앞 공정이 없음(원자재는 관리하지 않음) → 첫 공정 완료재고만 + q
       「출하」는 모든 아이템의 마지막 공정(자동) → 출하 실적은 마지막 공정 완료재고만 − q, 출하 누계 + q
   · C/T(개당 걸린 시간) = 실적 시간(start~end 에서 stop~restart 사이를 뺀 순작업 시간) ÷ 수량
   · 공정별 일별 생산수량 — 실적이 있는 날의 평균, 평균에서 30% 이상 벗어난 날
   브라우저(window.PPLogic)와 node(require) 양쪽에서 씁니다. */
(function (root, factory) {
  var api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.PPLogic = api;
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';
  /* 판(版) — index.html 의 ?v= · app.js 의 APP_VERSION 과 같아야 합니다(scripts/bump-version.mjs 로 함께 올림). */
  var VERSION = '2026-11-11.8';

  /* 저장 형식 판. 1 = 1·2차(발주서·재고 체크·과부족), 2 = 3차(아이템·기초재고·공정 실적·타이머) */
  var SCHEMA_VERSION = 2;
  var SHIP = '출하';
  var DEV_PCT = 0.3; // 평균에서 벗어났다고 볼 기준(30%)

  // ── 1. 날짜 (모두 'YYYY-MM-DD' 문자열, UTC 로 계산해 시간대 영향 없음) ──────
  var DAY = 86400000;
  var WEEKDAY_KO = ['일', '월', '화', '수', '목', '금', '토'];
  function pad(n) { return (n < 10 ? '0' : '') + n; }
  function isoFromUtc(ms) { var d = new Date(ms); return d.getUTCFullYear() + '-' + pad(d.getUTCMonth() + 1) + '-' + pad(d.getUTCDate()); }
  function utc(iso) { var p = iso.split('-'); return Date.UTC(+p[0], +p[1] - 1, +p[2]); }
  function validYmd(y, m, d) {
    if (!(y >= 1900 && y <= 2200 && m >= 1 && m <= 12 && d >= 1 && d <= 31)) return null;
    var ms = Date.UTC(y, m - 1, d), back = new Date(ms);
    if (back.getUTCMonth() !== m - 1 || back.getUTCDate() !== d) return null; // 2월 30일 같은 값 거부
    return isoFromUtc(ms);
  }
  /** 날짜 칸 값 → 'YYYY-MM-DD' 또는 null. Date · 엑셀 일련번호 · 2026-10-20 / 2026.10.20 / 2026/10/20 / 20261020 / 2026년 10월 20일 / 26.10.20 */
  function parseDate(v) {
    if (v == null || v === '') return null;
    if (v instanceof Date) {
      if (isNaN(v.getTime())) return null;
      return validYmd(v.getFullYear(), v.getMonth() + 1, v.getDate());
    }
    if (typeof v === 'number') {
      if (v >= 20000101 && v <= 22001231 && Math.floor(v) === v) return parseDate(String(v));
      if (v > 20000 && v < 100000) return isoFromUtc(Date.UTC(1899, 11, 30) + Math.round(v) * DAY); // 엑셀 일련번호
      return null;
    }
    var s = String(v).trim(), m;
    if ((m = s.match(/^(\d{4})[-./ ]\s*(\d{1,2})[-./ ]\s*(\d{1,2})\.?(?:\s.*)?$/))) return validYmd(+m[1], +m[2], +m[3]);
    if ((m = s.match(/^(\d{4})(\d{2})(\d{2})$/))) return validYmd(+m[1], +m[2], +m[3]);
    if ((m = s.match(/^(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일/))) return validYmd(+m[1], +m[2], +m[3]);
    if ((m = s.match(/^(\d{2})[-./](\d{1,2})[-./](\d{1,2})\.?$/))) return validYmd(2000 + +m[1], +m[2], +m[3]);
    return null;
  }
  function addDays(iso, n) { return isoFromUtc(utc(iso) + n * DAY); }
  function weekday(iso) { return new Date(utc(iso)).getUTCDay(); }
  function daysBetween(a, b) { return Math.round((utc(b) - utc(a)) / DAY); }
  function fmtDate(iso) { return iso ? iso + ' (' + WEEKDAY_KO[weekday(iso)] + ')' : ''; }
  function shortDate(iso) { return iso ? +iso.slice(5, 7) + '/' + +iso.slice(8, 10) : ''; }
  /** 이 PC 시계 기준 날짜(현지) — ms 또는 Date */
  function todayIso(now) { var d = now == null ? new Date() : new Date(+now); return validYmd(d.getFullYear(), d.getMonth() + 1, d.getDate()); }
  function hhmm(ms) { if (ms == null) return ''; var d = new Date(+ms); return pad(d.getHours()) + ':' + pad(d.getMinutes()); }
  /** 'YYYY-MM-DD' + 'HH:MM' → 현지 시각 ms */
  function localMs(iso, time) { var p = iso.split('-'), t = (time || '00:00').split(':'); return new Date(+p[0], +p[1] - 1, +p[2], +t[0] || 0, +t[1] || 0, 0, 0).getTime(); }

  // ── 2. 숫자 ─────────────────────────────────────────
  /** 수량·시간 칸 → 0 이상의 수 또는 null. 1,200 / 1200EA / 7.5h / ' 300 ' 허용 */
  function parseNum(v) {
    if (v == null || v === '') return null;
    if (typeof v === 'number') return isFinite(v) && v >= 0 ? v : null;
    var s = String(v).replace(/,/g, '').replace(/\s*(EA|ea|개|대|pcs|PCS|분|min|h|H|시간)$/, '').trim();
    if (!/^\d+(\.\d+)?$/.test(s)) return null;
    return Number(s);
  }
  /** 재고 칸 — 음수도 허용(기초재고를 틀리게 적은 것도 그대로 보이게). 빈 칸 = 0 */
  function parseQty(v) {
    if (v == null || v === '') return 0;
    if (typeof v === 'number') return isFinite(v) ? v : null;
    var s = String(v).replace(/,/g, '').replace(/\s*(EA|ea|개|대)$/, '').trim();
    if (s === '' || s === '-') return 0;
    if (!/^-?\d+(\.\d+)?$/.test(s)) return null;
    return Number(s);
  }
  function round(v, d) { if (v == null || !isFinite(v)) return null; var p = Math.pow(10, d || 0); return Math.round(v * p) / p; }
  function fmtN(v, d) {
    if (v == null || !isFinite(v)) return '-';
    var r = round(v, d == null ? 0 : d);
    var parts = String(Math.abs(r)).split('.');
    parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    return (r < 0 && r !== 0 ? '-' : '') + parts.join('.');
  }
  function fmtPct(v) { return v == null || !isFinite(v) ? '-' : round(v * 100, 1) + '%'; }
  /** 초 → 「1시간 02분 05초」 「3분 20초」 「45초」 */
  function fmtDur(sec) {
    if (sec == null || !isFinite(sec)) return '-';
    sec = Math.max(0, Math.round(sec));
    var h = Math.floor(sec / 3600), m = Math.floor(sec % 3600 / 60), s = sec % 60;
    if (h) return h + '시간 ' + pad(m) + '분 ' + pad(s) + '초';
    if (m) return m + '분 ' + pad(s) + '초';
    return s + '초';
  }
  /** 초 → 「01:02:05」 (타이머 표시) */
  function fmtClock(sec) { sec = Math.max(0, Math.floor(sec || 0)); return pad(Math.floor(sec / 3600)) + ':' + pad(Math.floor(sec % 3600 / 60)) + ':' + pad(sec % 60); }
  function str(v) { return v == null ? '' : String(v).trim(); }
  function normPart(v) { return str(v).replace(/\s+/g, ' ').toUpperCase(); }
  function normName(v) { return str(v).replace(/\s+/g, ' '); }
  function uniq(a) { var s = {}, o = []; a.forEach(function (x) { if (!s[x]) { s[x] = 1; o.push(x); } }); return o; }
  function sum(a, f) { var t = 0; a.forEach(function (x) { t += f ? f(x) : x; }); return t; }
  var seq = 0;
  function newId(prefix) { seq = (seq + 1) % 1000; return (prefix || 'x') + Date.now().toString(36) + pad(seq) + Math.floor(Math.random() * 1e4).toString(36); }

  // ── 3. 관리 아이템 · 공정 (설정) ─────────────────────
  var ITEM_FIELDS = [
    { key: 'customer', label: '고객사' },
    { key: 'group', label: '그룹' },
    { key: 'subgroup', label: '상세그룹' },
    { key: 'product', label: '품목' },
    { key: 'part', label: '품번' },
    { key: 'note', label: '비고' }
  ];
  var FILTER_KEYS = ['customer', 'group', 'subgroup', 'product'];
  /** 공정 순서 글 「가접 > 로봇 > 사상」 「가접, 로봇」 → { route:[], unknown:[], dup:[] }. 「출하」는 자동으로 붙으므로 적어도 뺍니다 */
  function parseRoute(text, processes) {
    var names = Array.isArray(text) ? text.map(normName) : String(text == null ? '' : text).split(/\s*(?:>|→|,|，|\/|\n|\t)\s*/).map(normName);
    var known = {}; (processes || []).forEach(function (p) { known[p] = 1; });
    var route = [], unknown = [], dup = [];
    names.forEach(function (n) {
      if (!n || n === SHIP) return;
      if (route.indexOf(n) >= 0) { dup.push(n); return; }
      if (processes && !known[n]) unknown.push(n);
      route.push(n);
    });
    return { route: route, unknown: unknown, dup: dup };
  }
  /** 아이템 하나 검사 → { ok, item, errors:{key:msg} } */
  function validateItem(input, processes, items, selfId) {
    var it = { id: input.id || selfId || newId('i') }, errors = {};
    ITEM_FIELDS.forEach(function (f) { it[f.key] = f.key === 'part' ? normPart(input[f.key]) : normName(input[f.key]); });
    if (!it.part) errors.part = '품번을 적어 주세요';
    else if ((items || []).some(function (x) { return x.part === it.part && x.id !== it.id; })) errors.part = '같은 품번이 이미 있습니다';
    var r = parseRoute(input.route, processes);
    it.route = r.route;
    if (!r.route.length) errors.route = '공정을 하나 이상 적어 주세요(출하는 자동으로 붙습니다)';
    else if (r.unknown.length) errors.route = '공정 목록에 없는 이름: ' + r.unknown.join(', ');
    else if (r.dup.length) errors.route = '같은 공정이 두 번: ' + r.dup.join(', ');
    return { ok: !Object.keys(errors).length, item: it, errors: errors };
  }
  /** 아이템의 전체 흐름(마지막에 출하) */
  function flowOf(item) { return (item.route || []).concat([SHIP]); }
  function routeText(item) { return (item.route || []).join(' > '); }
  /** 거르기 — f = { customer, group, subgroup, product } (빈 값은 모두) */
  function filterItems(items, f) {
    f = f || {};
    return items.filter(function (it) { return FILTER_KEYS.every(function (k) { return !f[k] || it[k] === f[k]; }); });
  }
  function filterOptions(items) {
    var o = {};
    FILTER_KEYS.forEach(function (k) { o[k] = uniq(items.map(function (it) { return it[k]; }).filter(Boolean)).sort(); });
    return o;
  }
  /** 공정 이름 바꾸기 — 아이템 공정 순서·기초재고·실적·타이머를 함께 바꿉니다 */
  function renameProcess(db, from, to) {
    to = normName(to);
    if (!to || to === from || to === SHIP) return false;
    if (db.processes.indexOf(to) >= 0) return false;
    db.processes = db.processes.map(function (p) { return p === from ? to : p; });
    db.items.forEach(function (it) { it.route = it.route.map(function (p) { return p === from ? to : p; }); });
    db.opening.rows.forEach(function (r) { if (r.process === from) r.process = to; });
    db.results.forEach(function (r) { if (r.process === from) r.process = to; });
    db.timers.forEach(function (t) { if (t.process === from) t.process = to; });
    return true;
  }
  /** 공정을 쓰는 곳 수 — 빼기 전에 묻기 */
  function processUse(db, name) {
    return {
      items: db.items.filter(function (it) { return it.route.indexOf(name) >= 0; }).length,
      results: db.results.filter(function (r) { return r.process === name; }).length,
      opening: db.opening.rows.filter(function (r) { return r.process === name && r.qty; }).length
    };
  }

  // ── 4. 엑셀 양식 — 관리 아이템 · 기초재고 ────────────
  /* 관리 아이템 양식은 두 모양을 다 읽습니다.
     · 한 줄에 공정 하나(기본 · 10-30 받은 실제 목록 모양) — 「고객사 | 그룹 | 상세그룹 | 품목 | 품번 | 비고 | 공정」,
       같은 품번을 공정 수만큼 줄로 적고 위에서 아래 순서가 곧 공정 순서. 「출하」 줄은 있어도 되고 없어도 됩니다(자동으로 마지막).
     · 한 줄에 아이템 하나(3차 첫 판) — 마지막 칸 「공정 순서」에 「가접 > 로봇 > 사상」. */
  var ITEM_HEAD = ['고객사', '그룹', '상세그룹', '품목', '품번', '비고', '공정'];
  var ITEM_HEAD_WIDE = ['고객사', '그룹', '상세그룹', '품목', '품번', '비고', '공정 순서'];
  var HEAD_SYN = {
    customer: ['고객사', '고객', '거래처', '납품처'], group: ['그룹', '제품군', '차급'], subgroup: ['상세그룹', '세부그룹', '하위그룹'],
    product: ['품목', '품명', '제품'], part: ['품번', '품목코드', '부품번호', '자재코드', 'partno'], note: ['비고', '메모'],
    route: ['공정순서', '공정', '공정명', '공정경로', 'route']
  };
  function normHeader(s) { return str(s).toLowerCase().replace(/[\s_()\[\]\-·.:/#]/g, ''); }
  function headerMap(row, keys) {
    var map = {};
    (row || []).forEach(function (h, i) {
      var n = normHeader(h);
      keys.forEach(function (k) { if (map[k] == null && HEAD_SYN[k].some(function (s) { return normHeader(s) === n; })) map[k] = i; });
    });
    return map;
  }
  function itemRowHead(it) { return [it.customer, it.group, it.subgroup, it.product, it.part, it.note]; }
  /** 지금 아이템 → 한 줄에 공정 하나 모양(출하 줄 포함). 다시 넣으면 그대로 읽힙니다 */
  function itemSheet(items) {
    var out = [ITEM_HEAD];
    items.forEach(function (it) { flowOf(it).forEach(function (p) { out.push(itemRowHead(it).concat([p])); }); });
    return out;
  }
  function itemTemplate(processes) {
    var procs = (processes && processes.length ? processes : ['공정1', '공정2']).slice(0, 3);
    var it = { customer: '고객사A', group: '중형', subgroup: '하부', product: '프레임', part: 'SMP-1001', note: '3PCS 세트', route: procs };
    return itemSheet([it]);
  }
  function rowItemFields(row, map) { var o = {}; ITEM_FIELDS.forEach(function (f) { o[f.key] = map[f.key] == null ? '' : row[map[f.key]]; }); return o; }
  /** 관리 아이템 엑셀 → { rows:[item], errors:[{row, msg}], newProcesses:[], format:'long'|'wide' }.
      공정 목록에 없는 공정 이름은 newProcesses 로 알려 줌(반영하면 공정 목록에 추가) */
  function parseItemSheet(aoa, processes) {
    var out = { rows: [], errors: [], newProcesses: [], format: null };
    var hr = -1, map = null;
    for (var i = 0; i < Math.min(aoa.length, 15); i++) {
      var m = headerMap(aoa[i], Object.keys(HEAD_SYN));
      if (m.part != null && m.route != null) { hr = i; map = m; break; }
    }
    if (hr < 0) { out.errors.push({ row: 0, msg: '머리행(「품번」「공정」 열)을 찾지 못했습니다 — 양식 받기로 받은 모양에 맞춰 주세요.' }); return out; }
    var known = (processes || []).slice();
    function learn(names) { parseRoute(names, known).unknown.forEach(function (n) { if (known.indexOf(n) < 0) { known.push(n); out.newProcesses.push(n); } }); }
    function routeNames(v) { return String(v == null ? '' : v).split(/\s*(?:>|→|,|，|\/|\n|\t)\s*/).map(normName).filter(Boolean); }
    var data = [];
    for (var r = hr + 1; r < aoa.length; r++) {
      var row = aoa[r] || [];
      if (!row.some(function (v) { return str(v) !== ''; })) continue;
      data.push({ r: r + 1, row: row, names: routeNames(row[map.route]), part: normPart(row[map.part]) });
    }
    // 모양 판별 — 모든 줄의 공정 칸이 이름 하나 이하이고, 품번이 두 줄 이상 나오거나 품번이 빈(위 칸을 이어 쓰는) 줄이 있으면 「한 줄에 공정 하나」
    var cnt = {}; data.forEach(function (d) { if (d.part) cnt[d.part] = (cnt[d.part] || 0) + 1; });
    var long = data.length && data.every(function (d) { return d.names.length <= 1; }) &&
      (Object.keys(cnt).some(function (k) { return cnt[k] > 1; }) || data.some(function (d) { return !d.part && d.names.length; }));
    out.format = long ? 'long' : 'wide';
    if (!long) {
      var seen = {};
      data.forEach(function (d) {
        var input = rowItemFields(d.row, map); input.route = d.row[map.route];
        learn(d.names);
        var v = validateItem(input, known, out.rows);
        if (v.item.part && seen[v.item.part]) { out.errors.push({ row: d.r, msg: v.item.part + ' — 같은 품번이 위에 또 있습니다(아래 줄은 건너뜀)' }); return; }
        if (!v.ok) { out.errors.push({ row: d.r, msg: (v.item.part || '(품번 없음)') + ' — ' + Object.keys(v.errors).map(function (k) { return v.errors[k]; }).join(' · ') }); return; }
        seen[v.item.part] = 1;
        out.rows.push(v.item);
      });
      return out;
    }
    // 한 줄에 공정 하나 — 품번별로 모아 위에서 아래 순서로 공정 순서를 만듦. 품번 칸이 비면(셀 병합 등) 바로 위 품번을 이어 씀
    var groups = [], byPart = {}, last = null;
    data.forEach(function (d) {
      var part = d.part || (last && last.part);
      if (!part) { out.errors.push({ row: d.r, msg: '(품번 없음) — 품번을 적어 주세요' }); return; }
      var g = byPart[part];
      var f = rowItemFields(d.row, map);
      if (!g) { g = byPart[part] = { part: part, row: d.r, fields: f, names: [], diff: [] }; groups.push(g); }
      else if (d.part) {
        ITEM_FIELDS.forEach(function (x) {
          if (x.key === 'part') return;
          var a = normName(g.fields[x.key]), b = normName(f[x.key]);
          if (!a && b) g.fields[x.key] = f[x.key];
          else if (b && a !== b && g.diff.indexOf(x.label) < 0) g.diff.push(x.label);
        });
      }
      if (d.names.length) g.names.push(d.names[0]);
      last = g;
    });
    groups.forEach(function (g) {
      learn(g.names);
      var input = Object.assign({}, g.fields, { part: g.part, route: g.names });
      var v = validateItem(input, known, out.rows);
      if (g.diff.length) out.errors.push({ row: g.row, msg: g.part + ' — 줄마다 ' + g.diff.join('·') + ' 값이 다릅니다(첫 줄 값으로 넣음)' });
      if (!v.ok) { out.errors.push({ row: g.row, msg: g.part + ' — ' + Object.keys(v.errors).map(function (k) { return v.errors[k]; }).join(' · ') }); return; }
      out.rows.push(v.item);
    });
    return out;
  }
  /** 가져온 아이템을 지금 목록에 합침 — 같은 품번은 내용을 바꾸고(id 유지 → 실적·기초재고 그대로), 새 품번은 추가 */
  function mergeItems(items, add) {
    var byPart = {}; items.forEach(function (it) { byPart[it.part] = it; });
    var added = 0, updated = 0;
    add.forEach(function (n) {
      var o = byPart[n.part];
      if (o) { ITEM_FIELDS.forEach(function (f) { o[f.key] = n[f.key]; }); o.route = n.route; updated++; }
      else { items.push(n); byPart[n.part] = n; added++; }
    });
    return { added: added, updated: updated };
  }
  /** 기초재고 양식: 1행 「기준날짜 | 2026-10-12」, 3행 머리행 「품번 | 품목 | 공정1 | 공정2 …」(공정 목록 순서). 그 아이템이 거치지 않는 공정 칸은 「-」 */
  function openingSheet(db) {
    var procs = db.processes, by = openingIndex(db.opening);
    var head = ['품번', '품목'].concat(procs);
    return [['기준날짜', db.opening.date || ''], ['(기준날짜 아침 = 그날 작업 시작 전 공정별 완료재고. 그 아이템이 거치지 않는 공정은 「-」)'], head].concat(db.items.map(function (it) {
      return [it.part, it.product].concat(procs.map(function (p) { return it.route.indexOf(p) < 0 ? '-' : (by[it.id + '|' + p] || 0); }));
    }));
  }
  function openingIndex(op) { var by = {}; (op.rows || []).forEach(function (r) { by[r.itemId + '|' + r.process] = r.qty; }); return by; }
  /** 기초재고 엑셀 → { date, rows:[{itemId, process, qty}], errors:[], unknownParts:[], offRoute:[] } */
  function parseOpeningSheet(aoa, db) {
    var out = { date: null, rows: [], errors: [], unknownParts: [], offRoute: [] };
    var hr = -1, partCol = -1;
    for (var i = 0; i < Math.min(aoa.length, 20); i++) {
      var row = aoa[i] || [];
      for (var c = 0; c < row.length; c++) {
        var n = normHeader(row[c]);
        if (n === '기준날짜' || n === '기준일') { var d = parseDate(row[c + 1]); if (d) out.date = d; }
        if (hr < 0 && HEAD_SYN.part.some(function (s) { return normHeader(s) === n; })) { hr = i; partCol = c; }
      }
      if (hr >= 0) break;
    }
    if (hr < 0) { out.errors.push('머리행(「품번」 열)을 찾지 못했습니다 — 「기초재고 양식 받기」로 받은 모양에 맞춰 주세요.'); return out; }
    if (!out.date) out.errors.push('「기준날짜」 칸을 찾지 못했습니다 — 화면에서 기준날짜를 적어 주세요.');
    var head = aoa[hr], cols = [];
    head.forEach(function (h, c) { var nm = normName(h); if (db.processes.indexOf(nm) >= 0) cols.push({ c: c, p: nm }); });
    if (!cols.length) out.errors.push('공정 열이 없습니다 — 머리행의 공정 이름이 설정의 공정 목록과 같아야 합니다.');
    var byPart = {}; db.items.forEach(function (it) { byPart[it.part] = it; });
    for (var r = hr + 1; r < aoa.length; r++) {
      var rw = aoa[r] || [], part = normPart(rw[partCol]);
      if (!part) continue;
      if (/^(합계|계|total)$/i.test(part)) continue;
      var it = byPart[part];
      if (!it) { out.unknownParts.push(part); continue; }
      cols.forEach(function (x) {
        var raw = rw[x.c];
        if (str(raw) === '-') return;
        var q = parseQty(raw);
        if (q == null) { out.errors.push((r + 1) + '행 ' + part + ' · ' + x.p + ' — 숫자가 아닙니다(' + str(raw) + ')'); return; }
        if (it.route.indexOf(x.p) < 0) { if (q) out.offRoute.push(part + ' · ' + x.p + ' ' + q); return; }
        out.rows.push({ itemId: it.id, process: x.p, qty: q });
      });
    }
    return out;
  }

  // ── 5. 타이머 (start · stop · restart · end) ────────
  /* 타이머 = { id, itemId, process, startedAt, segs:[{from, to|null}] }. 마지막 칸 to 가 null 이면 「작업 중」.
     stop 은 멈춤(쉬는 시간·설비 정지), restart 는 다시 시작, end 는 수량을 받아 실적 한 건으로 닫습니다.
     시각은 ms(Date.now()) — 브라우저에 저장되므로 새로고침·화면을 꺼도 이어집니다. */
  function timerStart(itemId, process, now) { return { id: newId('t'), itemId: itemId, process: process, startedAt: now, segs: [{ from: now, to: null }] }; }
  function timerRunning(t) { var s = t.segs[t.segs.length - 1]; return !!s && s.to == null; }
  function timerStop(t, now) { var s = t.segs[t.segs.length - 1]; if (s && s.to == null) s.to = Math.max(now, s.from); return t; }
  function timerRestart(t, now) { if (!timerRunning(t)) t.segs.push({ from: now, to: null }); return t; }
  /** 순작업 시간(ms) — 멈춘 사이는 빼고 */
  function timerNetMs(t, now) { return sum(t.segs, function (s) { return Math.max(0, (s.to == null ? now : s.to) - s.from); }); }
  /** 멈춘 시간(ms) — 지금 멈춰 있으면 멈춘 뒤 지금까지도 셈 */
  function timerPausedMs(t, now) { return Math.max(0, now - t.startedAt - timerNetMs(t, now)); }
  /** end — 수량을 받아 실적으로. → { ok, result, error } */
  function timerFinish(t, qty, now, note) {
    var q = parseNum(qty);
    if (q == null || q <= 0) return { ok: false, error: '수량을 1 이상으로 적어 주세요' };
    var stopped = JSON.parse(JSON.stringify(t)); timerStop(stopped, now);
    var endAt = stopped.segs[stopped.segs.length - 1].to;
    var net = timerNetMs(stopped, endAt);
    return { ok: true, result: {
      id: newId('r'), itemId: t.itemId, process: t.process, qty: q, date: todayIso(endAt),
      start: t.startedAt, end: endAt, runSec: Math.round(net / 1000), stops: stopped.segs.length - 1, source: '타이머', note: str(note)
    } };
  }
  /** 타이머 없이 바로 적기 — { date, itemId, process, qty, minutes(선택), note } */
  function validateManual(input, db) {
    var e = {}, it = db.items.filter(function (x) { return x.id === input.itemId; })[0];
    var date = parseDate(input.date); if (!date) e.date = '일자를 골라 주세요';
    if (!it) e.itemId = '품번을 골라 주세요';
    var proc = normName(input.process);
    if (it && flowOf(it).indexOf(proc) < 0) e.process = '이 품번의 공정 순서에 없는 공정입니다';
    var q = parseNum(input.qty); if (q == null || q <= 0) e.qty = '수량을 1 이상으로 적어 주세요';
    var min = str(input.minutes) === '' ? null : parseNum(input.minutes);
    if (str(input.minutes) !== '' && min == null) e.minutes = '분을 숫자로 적어 주세요(비워도 됨)';
    if (Object.keys(e).length) return { ok: false, errors: e };
    return { ok: true, errors: {}, result: { id: newId('r'), itemId: it.id, process: proc, qty: q, date: date, start: null, end: null, runSec: min == null ? null : Math.round(min * 60), stops: 0, source: '직접', note: str(input.note) } };
  }

  // ── 6. 재고 흐름 ─────────────────────────────────────
  /** 기준날짜 이후(그날 포함) 실적만 — 기준날짜가 없으면 모든 실적 */
  function resultsAfter(results, date) { return date ? results.filter(function (r) { return r.date >= date; }) : results.slice(); }
  /** 아이템별 공정 완료재고. → [{ item, flow, opening:{p:q}, done:{p:q}, used:{p:q}, stock:{p:q}, shipped, wip, negatives:[p], offRoute }] */
  function computeStock(db, items) {
    var by = openingIndex(db.opening), list = items || db.items, rs = resultsAfter(db.results, db.opening.date);
    var idx = {};
    var rows = list.map(function (it) {
      var o = { item: it, flow: flowOf(it), opening: {}, done: {}, used: {}, stock: {}, shipped: 0, wip: 0, negatives: [], offRoute: 0 };
      it.route.forEach(function (p) { o.opening[p] = by[it.id + '|' + p] || 0; o.done[p] = 0; o.used[p] = 0; });
      idx[it.id] = o; return o;
    });
    rs.forEach(function (r) {
      var o = idx[r.itemId]; if (!o) return;
      var i = o.flow.indexOf(r.process);
      if (i < 0) { o.offRoute++; return; }
      if (r.process === SHIP) o.shipped += r.qty; else o.done[r.process] += r.qty;
      if (i > 0) o.used[o.flow[i - 1]] += r.qty; // 앞 공정 완료재고를 가져다 씀
    });
    rows.forEach(function (o) {
      o.item.route.forEach(function (p) {
        o.stock[p] = o.opening[p] + o.done[p] - o.used[p];
        o.wip += o.stock[p];
        if (o.stock[p] < 0) o.negatives.push(p);
      });
    });
    return rows;
  }

  // ── 7. C/T ───────────────────────────────────────────
  /** 시간이 있는 실적으로 아이템 × 공정 C/T. → [{ itemId, process, n, qty, sec, ct }] (ct = 초/개, 수량 가중 = 시간 합 ÷ 수량 합) */
  function ctTable(results, itemIds) {
    var ok = itemIds ? {} : null; if (ok) itemIds.forEach(function (id) { ok[id] = 1; });
    var m = {};
    results.forEach(function (r) {
      if (r.runSec == null || !(r.qty > 0) || (ok && !ok[r.itemId])) return;
      var k = r.itemId + '|' + r.process;
      var o = m[k] || (m[k] = { itemId: r.itemId, process: r.process, n: 0, qty: 0, sec: 0, min: null, max: null });
      var ct = r.runSec / r.qty;
      o.n++; o.qty += r.qty; o.sec += r.runSec;
      o.min = o.min == null ? ct : Math.min(o.min, ct); o.max = o.max == null ? ct : Math.max(o.max, ct);
    });
    return Object.keys(m).map(function (k) { var o = m[k]; o.ct = o.sec / o.qty; return o; });
  }
  /** 공정별로 묶은 C/T(아이템을 섞은 수량 가중 평균) */
  function ctByProcess(ct, processes) {
    return processes.map(function (p) {
      var rows = ct.filter(function (x) { return x.process === p; });
      var q = sum(rows, function (x) { return x.qty; }), s = sum(rows, function (x) { return x.sec; });
      return { process: p, n: sum(rows, function (x) { return x.n; }), qty: q, sec: s, ct: q ? s / q : null };
    });
  }

  // ── 8. 공정별 일별 생산수량 ──────────────────────────
  /** 한 공정의 날짜별 수량. 평균 = 실적이 있는 날의 평균. 평균에서 30% 이상 벗어난 날 표시
      → { process, days:[{date, qty, n, dev, flag:'많음'|'적음'|''}], avg, max, min, flagged } */
  function dailyProduction(results, process, itemIds, from) {
    var ok = itemIds ? {} : null; if (ok) itemIds.forEach(function (id) { ok[id] = 1; });
    var m = {};
    results.forEach(function (r) {
      if (r.process !== process || (ok && !ok[r.itemId]) || (from && r.date < from)) return;
      var d = m[r.date] || (m[r.date] = { date: r.date, qty: 0, n: 0, rows: [] });
      d.qty += r.qty; d.n++; d.rows.push(r);
    });
    var days = Object.keys(m).sort().map(function (k) { return m[k]; });
    var avg = days.length ? sum(days, function (d) { return d.qty; }) / days.length : null;
    days.forEach(function (d) {
      d.dev = avg ? (d.qty - avg) / avg : null;
      d.flag = d.dev != null && Math.abs(d.dev) >= DEV_PCT - 1e-9 ? (d.dev > 0 ? '많음' : '적음') : '';
    });
    return {
      process: process, days: days, avg: avg,
      max: days.length ? Math.max.apply(null, days.map(function (d) { return d.qty; })) : null,
      min: days.length ? Math.min.apply(null, days.map(function (d) { return d.qty; })) : null,
      flagged: days.filter(function (d) { return d.flag; })
    };
  }

  // ── 9. 한 번에 계산 ──────────────────────────────────
  /** 화면에 쓰는 결과. f = 거르기 { customer, group, subgroup, product } */
  function analyze(db, f) {
    var items = filterItems(db.items, f), ids = items.map(function (it) { return it.id; });
    var procs = db.processes.filter(function (p) { return items.some(function (it) { return it.route.indexOf(p) >= 0; }); });
    var stock = computeStock(db, items);
    var after = resultsAfter(db.results, db.opening.date);
    var ct = ctTable(db.results, ids);
    var flowProcs = procs.concat([SHIP]);
    // 일별 생산수량은 기본으로 기초재고 기준날짜부터(f.from 으로 바꿈, '' = 모든 날)
    var from = f && f.from != null ? f.from : (db.opening.date || '');
    var daily = flowProcs.map(function (p) { return dailyProduction(db.results, p, ids, from); });
    var procStock = procs.map(function (p) { return { process: p, qty: sum(stock, function (o) { return o.stock[p] || 0; }), items: stock.filter(function (o) { return o.item.route.indexOf(p) >= 0; }).length }; });
    var checks = [];
    if (!db.items.length) checks.push({ level: '정보', text: '관리 아이템이 없습니다 — 설정에서 아이템(품번·공정 순서)을 먼저 넣어 주세요.', href: '#/settings' });
    if (db.items.length && !db.opening.date) checks.push({ level: '주의', text: '기초재고 기준날짜가 없습니다 — 모든 실적을 0 에서부터 더하고 있습니다.', href: '#/import' });
    stock.forEach(function (o) {
      if (o.negatives.length) checks.push({ level: '경고', text: o.item.part + ' — ' + o.negatives.map(function (p) { return p + ' ' + fmtN(o.stock[p]); }).join(', ') + ' (완료재고가 0 아래: 기초재고나 앞 공정 실적이 빠졌는지 확인)', href: '#/import', itemId: o.item.id });
      if (o.offRoute) checks.push({ level: '주의', text: o.item.part + ' — 공정 순서에 없는 공정 실적 ' + o.offRoute + '건은 재고에 넣지 않았습니다(설정에서 공정 순서 확인).', href: '#/settings' });
    });
    var before = db.opening.date ? db.results.filter(function (r) { return r.date < db.opening.date; }).length : 0;
    if (before) checks.push({ level: '정보', text: '기준날짜(' + db.opening.date + ') 전 실적 ' + before + '건은 기초재고에 이미 들어 있다고 보고 재고에서 뺐습니다(C/T·일별 생산에는 씀).', href: '#/import' });
    return {
      items: items, processes: procs, flowProcs: flowProcs, stock: stock, from: from, ct: ct, ctProc: ctByProcess(ct, flowProcs), daily: daily, procStock: procStock,
      openingDate: db.opening.date, resultsAfter: after.filter(function (r) { return ids.indexOf(r.itemId) >= 0; }).length, checks: checks,
      kpi: {
        items: items.length,
        wip: sum(stock, function (o) { return o.wip; }),
        shipped: sum(stock, function (o) { return o.shipped; }),
        negatives: sum(stock, function (o) { return o.negatives.length; }),
        timers: db.timers.length,
        flagged: sum(daily, function (d) { return d.flagged.length; })
      }
    };
  }
  function itemById(db) { var m = {}; db.items.forEach(function (it) { m[it.id] = it; }); return m; }

  // ── 10. AI 프롬프트 (계산된 사실만, 품번·고객사는 별칭) ─────
  var ITEM_ALIAS = '품목', CUST_ALIAS = '거래처';
  function assignAliases(map, keys, prefix) {
    var out = {}, used = 0, re = new RegExp('^' + prefix + '(\\d+)$');
    Object.keys(map || {}).forEach(function (k) { var m = re.exec(map[k]); if (m) { out[k] = map[k]; used = Math.max(used, +m[1]); } });
    uniq((keys || []).map(str).filter(Boolean)).sort().forEach(function (k) { if (!out[k]) out[k] = prefix + (++used); });
    return out;
  }
  /** opts: { maskItems, maskCustomers, aliases:{items, customers} } → { prompt, used, leaks } */
  function buildAiPrompt(res, db, opts) {
    opts = opts || {};
    var ia = opts.maskItems !== false ? (opts.aliases && opts.aliases.items) || {} : null;
    var ca = opts.maskCustomers !== false ? (opts.aliases && opts.aliases.customers) || {} : null;
    var used = { items: {}, customers: {} }, byId = itemById(db);
    function P(code) { if (ia && ia[code]) { used.items[code] = ia[code]; return ia[code]; } return code; }
    function C(c) { if (!c) return '고객사 미기재'; if (ca && ca[c]) { used.customers[c] = ca[c]; return ca[c]; } return c; }
    var L = [];
    L.push('너는 제조업 계획운영팀의 생산관리 보조야. 아래 [사실]은 현장 공정 실적(시작·종료 단추로 잰 시간과 수량)과 기초재고로 도구가 규칙으로 계산한 값이야.');
    L.push('[사실]에 있는 숫자만 근거로 삼아 줘. 숫자를 새로 만들거나 바꾸지 말고, 기록에 없는 원인은 반드시 「추정」이라고 표시해 줘.');
    if (ia || ca) L.push('보안상 ' + [ia ? '품번은 「' + ITEM_ALIAS + '1」' : '', ca ? '고객사는 「' + CUST_ALIAS + '1」' : ''].filter(Boolean).join(', ') + ' 같은 별칭으로 바꿔 두었어. 답에서도 별칭을 그대로 써 줘.');
    L.push('');
    L.push('[사실]');
    L.push('- 계산식: 공정 완료재고 = 기초재고(기준날짜 ' + (res.openingDate || '없음') + ' 아침) + 그 공정 실적 − 다음 공정 실적. 출하 실적은 마지막 공정 완료재고에서 빠짐. C/T = 순작업 시간(멈춘 시간 제외) ÷ 수량.');
    L.push('- 공정별 완료재고 합계 (공정 | 완료재고 | 거치는 품번 수)');
    res.procStock.forEach(function (p) { L.push('  ' + [p.process, p.qty, p.items].join(' | ')); });
    L.push('- 품번별 완료재고 (고객사 | 품번 | 공정:재고 … | 출하 누계)');
    res.stock.forEach(function (o) {
      L.push('  ' + [C(o.item.customer), P(o.item.part), o.item.route.map(function (p) { return p + ':' + o.stock[p]; }).join(' '), o.shipped].join(' | '));
    });
    L.push('- 공정별 평균 C/T (공정 | 실적 건수 | 수량 | 초/개)');
    res.ctProc.forEach(function (c) { if (c.n) L.push('  ' + [c.process, c.n, c.qty, round(c.ct, 1)].join(' | ')); });
    L.push('- 품번 × 공정 C/T 중 같은 공정 평균보다 30% 이상 긴 것 (품번 | 공정 | 초/개 | 공정 평균 초/개)');
    var slow = 0;
    res.ct.forEach(function (c) {
      var pa = res.ctProc.filter(function (x) { return x.process === c.process; })[0];
      if (pa && pa.ct && c.ct >= pa.ct * (1 + DEV_PCT)) { slow++; L.push('  ' + [P((byId[c.itemId] || {}).part || '?'), c.process, round(c.ct, 1), round(pa.ct, 1)].join(' | ')); }
    });
    if (!slow) L.push('  (없음)');
    L.push('- 공정별 일별 생산수량 — 평균에서 30% 이상 벗어난 날 (공정 | 일 평균 | 일자:수량 …)');
    res.daily.forEach(function (d) { if (d.flagged.length) L.push('  ' + [d.process, round(d.avg, 1), d.flagged.map(function (x) { return x.date + ':' + x.qty; }).join(' ')].join(' | ')); });
    if (!res.daily.some(function (d) { return d.flagged.length; })) L.push('  (없음)');
    L.push('');
    L.push('[요청] 한국어로 짧게 정리해 줘.');
    L.push('1. 공정 재고 흐름 한 줄 요약 — 어느 공정 앞에 재고가 쌓이고 어디가 비는지');
    L.push('2. C/T 가 긴 공정·품번과 그 숫자');
    L.push('3. 생산수량이 평균에서 크게 벗어난 날 — 현장에서 확인할 것(공정마다 2~3개)');
    L.push('확실하지 않은 것은 「확인 필요」라고 적어 줘.');
    var text = L.join('\n'), leaks = [];
    if (ia) Object.keys(ia).forEach(function (k) { if (k.length >= 3 && text.indexOf(k) >= 0) leaks.push(k); });
    if (ca) Object.keys(ca).forEach(function (k) { if (k.length >= 2 && text.indexOf(k) >= 0) leaks.push(k); });
    return { prompt: text, used: used, leaks: leaks };
  }
  function unmaskText(text, aliases) {
    var back = {};
    [(aliases && aliases.items) || {}, (aliases && aliases.customers) || {}].forEach(function (m) { Object.keys(m).forEach(function (k) { back[m[k]] = k; }); });
    return String(text == null ? '' : text).replace(new RegExp('(' + ITEM_ALIAS + '|' + CUST_ALIAS + ')(\\d+)(?![0-9개건종가곳])', 'g'), function (all) { return back[all] || all; });
  }

  // ── 11. 내보내기 ─────────────────────────────────────
  var RESULT_HEAD = ['일자', '고객사', '품번', '품목', '공정', '수량', '시작', '종료', '순작업(분)', '멈춤 횟수', 'C/T(초/개)', '입력경로', '비고'];
  function dt(ms) { return ms == null ? '' : todayIso(ms) + ' ' + hhmm(ms); }
  function resultSheet(db) {
    var by = itemById(db);
    return [RESULT_HEAD].concat(db.results.slice().sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : (a.end || 0) - (b.end || 0); }).map(function (r) {
      var it = by[r.itemId] || {};
      return [r.date, it.customer || '', it.part || '(지운 품번)', it.product || '', r.process, r.qty, dt(r.start), dt(r.end), r.runSec == null ? '' : round(r.runSec / 60, 2), r.stops || 0, r.runSec == null ? '' : round(r.runSec / r.qty, 1), r.source || '', r.note || ''];
    }));
  }
  function exportSheets(res, db) {
    var byId = itemById(db), procs = res.processes;
    var sh = {};
    sh['공정별 재고현황'] = [['고객사', '그룹', '상세그룹', '품목', '품번'].concat(procs, ['출하 누계', '재공 합계'])].concat(res.stock.map(function (o) {
      var it = o.item;
      return [it.customer, it.group, it.subgroup, it.product, it.part].concat(procs.map(function (p) { return it.route.indexOf(p) < 0 ? '' : o.stock[p]; }), [o.shipped, o.wip]);
    }));
    sh['C/T'] = [['품번', '공정', '실적 건수', '수량', '순작업(분)', 'C/T(초/개)', '가장 짧은', '가장 긴']].concat(res.ct.map(function (c) {
      return [(byId[c.itemId] || {}).part || '', c.process, c.n, c.qty, round(c.sec / 60, 1), round(c.ct, 1), round(c.min, 1), round(c.max, 1)];
    }));
    sh['일별 생산수량'] = [['공정', '일자', '수량', '평균', '평균 대비', '판정']];
    res.daily.forEach(function (d) { d.days.forEach(function (x) { sh['일별 생산수량'].push([d.process, x.date, x.qty, round(d.avg, 1), x.dev == null ? '' : round(x.dev * 100, 1) + '%', x.flag]); }); });
    sh['공정 실적'] = resultSheet(db);
    return sh;
  }

  // ── 12. 저장 형식 ────────────────────────────────────
  function defaultProcesses() { return ['가접', '로봇', '사상', 'MT', '가공', '도킹']; }
  function emptyDb() {
    return { schemaVersion: SCHEMA_VERSION, processes: defaultProcesses(), items: [], opening: { date: '', rows: [] }, results: [], timers: [],
      aliases: { items: {}, customers: {} }, maskItems: true, maskCustomers: true, aiNote: '', sample: false, notice: null };
  }
  function cleanItem(o, procs) {
    var it = { id: str(o.id) || newId('i') };
    ITEM_FIELDS.forEach(function (f) { it[f.key] = f.key === 'part' ? normPart(o[f.key]) : normName(o[f.key]); });
    it.route = parseRoute(Array.isArray(o.route) ? o.route : [], null).route.filter(function (p) { return procs.indexOf(p) >= 0; });
    return it;
  }
  /** 저장된 것 → 지금 형식. 1·2차(schemaVersion 1) 자료는 공정 이름과 모델(품번)만 옮기고, notice 로 알립니다 */
  function restoreDb(o) {
    var db = emptyDb();
    if (!o || typeof o !== 'object') return db;
    if (o.schemaVersion !== SCHEMA_VERSION) {
      var procs = o.settings && Array.isArray(o.settings.processes) ? uniq(o.settings.processes.map(function (p) { return normName(p && p.name); }).filter(function (n) { return n && n !== SHIP; })) : [];
      if (procs.length) db.processes = procs;
      var models = [];
      (o.checks || []).forEach(function (c) { (c.rows || []).forEach(function (r) { models.push(normPart(r.item)); }); });
      ((o.stock && o.stock.rows) || []).forEach(function (r) { models.push(normPart(r.item)); });
      models = uniq(models.filter(Boolean));
      if (!o.sample) db.items = models.map(function (m) { return { id: newId('i'), customer: '', group: '', subgroup: '', product: '', part: m, note: '', route: db.processes.slice() }; });
      var had = ['results', 'orders', 'deliveries', 'checks'].reduce(function (t, k) { return t + (Array.isArray(o[k]) ? o[k].length : 0); }, 0);
      db.notice = { from: o.schemaVersion || 1, had: had, items: db.items.length, sample: !!o.sample };
      return db;
    }
    db.processes = Array.isArray(o.processes) ? uniq(o.processes.map(normName).filter(function (n) { return n && n !== SHIP; })) : db.processes;
    db.items = Array.isArray(o.items) ? o.items.filter(function (x) { return x && typeof x === 'object'; }).map(function (x) { return cleanItem(x, db.processes); }) : [];
    var ids = {}; db.items.forEach(function (it) { ids[it.id] = 1; });
    if (o.opening && typeof o.opening === 'object') db.opening = { date: parseDate(o.opening.date) || '', rows: (o.opening.rows || []).filter(function (r) { return r && ids[r.itemId] && typeof r.qty === 'number'; }) };
    db.results = (o.results || []).filter(function (r) { return r && ids[r.itemId] && r.process && r.qty > 0 && parseDate(r.date); });
    db.timers = (o.timers || []).filter(function (t) { return t && ids[t.itemId] && t.process && Array.isArray(t.segs) && t.segs.length; });
    if (o.aliases && typeof o.aliases === 'object') db.aliases = { items: o.aliases.items || {}, customers: o.aliases.customers || {} };
    db.maskItems = o.maskItems !== false; db.maskCustomers = o.maskCustomers !== false;
    db.aiNote = str(o.aiNote); db.sample = !!o.sample;
    db.notice = o.notice && typeof o.notice === 'object' ? o.notice : null;
    return db;
  }

  return {
    VERSION: VERSION, SCHEMA_VERSION: SCHEMA_VERSION, SHIP: SHIP, DEV_PCT: DEV_PCT, WEEKDAY_KO: WEEKDAY_KO,
    ITEM_FIELDS: ITEM_FIELDS, FILTER_KEYS: FILTER_KEYS, ITEM_HEAD: ITEM_HEAD, ITEM_HEAD_WIDE: ITEM_HEAD_WIDE, RESULT_HEAD: RESULT_HEAD, ITEM_ALIAS: ITEM_ALIAS, CUST_ALIAS: CUST_ALIAS,
    parseDate: parseDate, addDays: addDays, weekday: weekday, daysBetween: daysBetween, fmtDate: fmtDate, shortDate: shortDate, todayIso: todayIso, hhmm: hhmm, localMs: localMs,
    parseNum: parseNum, parseQty: parseQty, round: round, fmtN: fmtN, fmtPct: fmtPct, fmtDur: fmtDur, fmtClock: fmtClock, normPart: normPart, newId: newId,
    parseRoute: parseRoute, validateItem: validateItem, flowOf: flowOf, routeText: routeText, filterItems: filterItems, filterOptions: filterOptions,
    renameProcess: renameProcess, processUse: processUse,
    itemSheet: itemSheet, itemTemplate: itemTemplate, parseItemSheet: parseItemSheet, mergeItems: mergeItems, openingSheet: openingSheet, openingIndex: openingIndex, parseOpeningSheet: parseOpeningSheet,
    timerStart: timerStart, timerRunning: timerRunning, timerStop: timerStop, timerRestart: timerRestart, timerNetMs: timerNetMs, timerPausedMs: timerPausedMs, timerFinish: timerFinish, validateManual: validateManual,
    resultsAfter: resultsAfter, computeStock: computeStock, ctTable: ctTable, ctByProcess: ctByProcess, dailyProduction: dailyProduction, analyze: analyze, itemById: itemById,
    assignAliases: assignAliases, buildAiPrompt: buildAiPrompt, unmaskText: unmaskText,
    resultSheet: resultSheet, exportSheets: exportSheets,
    defaultProcesses: defaultProcesses, emptyDb: emptyDb, restoreDb: restoreDb
  };
});
