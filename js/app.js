/* 공정 재고·C/T — 화면 (3차 2026-10-28 범위 재설정)
   계산은 js/logic.js(PPLogic), 그래프는 js/charts.js, 저장은 js/store.js, AI 연결은 js/ai-endpoint.js·ai-panel.js.
   화면: 0 대시보드 · 1 실적 입력(타이머) · 2 기초재고 · 3 공정 분석 · 4 AI 분석 · 설정 */
(function () {
  'use strict';
  var L = window.PPLogic, S = window.PPStore, C = window.PPCharts, AIP = window.AIPanel;
  /* 판 맞추기 — 브라우저가 옛 logic.js 와 새 app.js 를 섞어 쓰면(Pages 캐시) 계산하지 않고 새로고침을 안내합니다.
     판은 scripts/bump-version.mjs 로 index.html(?v=) · logic.js 와 함께 올립니다. */
  var APP_VERSION = '2026-11-11.8';
  if (!L || L.VERSION !== APP_VERSION) {
    var m0 = document.getElementById('main') || document.body;
    m0.innerHTML = '<section class="card" id="versionMismatch"><h1>새 판이 나왔습니다 — 새로고침(Ctrl+F5)해 주세요</h1><p class="alert warn">이 브라우저가 옛 파일과 새 파일을 섞어 불러왔습니다(js/logic.js ' +
      (L && L.VERSION ? L.VERSION : '옛 판') + ' · 화면 ' + APP_VERSION + '). 이대로 계산하면 버튼과 결과가 어긋날 수 있어 멈췄습니다.</p>' +
      '<p>Windows 는 <b>Ctrl+F5</b>, Mac 은 <b>Cmd+Shift+R</b> 을 한 번 눌러 주세요. 쌓아 둔 자료는 지워지지 않습니다.</p></section>';
    return;
  }
  (function () { var ft = document.querySelector('.site-footer'); if (!ft) return; var s = document.createElement('span'); s.className = 'ver'; s.textContent = ' · 판 ' + APP_VERSION; ft.appendChild(s); })();
  var db = S.loadDb();
  var pref = S.loadUi();
  if (!pref.filter) pref.filter = {};
  var ui = { tab: 'timer', manualErr: {}, proc: '', from: null, itemErr: {}, calcItem: '' };
  var main = document.getElementById('main');

  /* ── 작은 도구 ───────────────────────────── */
  function h(tag, attrs) {
    var e = document.createElement(tag);
    if (attrs) Object.keys(attrs).forEach(function (k) {
      var v = attrs[k];
      if (v == null || v === false) return;
      if (k === 'class') e.className = v;
      else if (k.slice(0, 2) === 'on' && typeof v === 'function') e.addEventListener(k.slice(2), v);
      else if (v === true) e.setAttribute(k, '');
      else e.setAttribute(k, v);
    });
    for (var i = 2; i < arguments.length; i++) add(e, arguments[i]);
    return e;
  }
  function add(e, c) {
    if (c == null || c === false) return;
    if (Array.isArray(c)) { c.forEach(function (x) { add(e, x); }); return; }
    e.appendChild(typeof c === 'object' ? c : document.createTextNode(String(c)));
  }
  function fld(label, control, hint, err) {
    return h('label', { class: 'field' + (err ? ' invalid' : '') }, h('span', null, label), control, hint ? h('small', { class: 'hint' }, hint) : null, err ? h('small', { class: 'err' }, err) : null);
  }
  var toastTimer = null;
  function toast(msg, err) {
    var t = document.getElementById('toast');
    t.textContent = msg; t.className = 'toast' + (err ? ' error' : ''); t.hidden = false;
    clearTimeout(toastTimer); toastTimer = setTimeout(function () { t.hidden = true; }, 4200);
  }
  function save() {
    var ok = S.saveDb(db);
    document.getElementById('storeNote').textContent = ok ? '이 브라우저에 저장됨 · ' + S.sizeKb() + 'KB' : '저장 안 됨(브라우저 저장소가 막혔거나 가득 참) — 설정에서 JSON 백업을 받아 두세요';
    return ok;
  }
  function download(name, blob) {
    var a = h('a', { href: URL.createObjectURL(blob), download: name });
    document.body.appendChild(a); a.click(); setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 500);
  }
  function xlsxBlob(sheets) {
    var wb = XLSX.utils.book_new();
    Object.keys(sheets).forEach(function (name) {
      var ws = XLSX.utils.aoa_to_sheet(sheets[name]);
      ws['!cols'] = (sheets[name][0] || []).map(function (x) { return { wch: Math.max(10, String(x == null ? '' : x).length * 2 + 2) }; });
      XLSX.utils.book_append_sheet(wb, ws, name.slice(0, 31));
    });
    return new Blob([XLSX.write(wb, { bookType: 'xlsx', type: 'array' })], { type: 'application/octet-stream' });
  }
  /* 엑셀 파일 → [{name, aoa}] */
  function readSheets(file) {
    return new Promise(function (resolve, reject) {
      var fr = new FileReader();
      fr.onload = function () {
        try {
          var wb = XLSX.read(new Uint8Array(fr.result), { type: 'array', cellDates: true });
          resolve(wb.SheetNames.map(function (n) { return { name: n, aoa: XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: null }) }; }));
        } catch (e) { reject(new Error(file.name + ' — 엑셀로 읽지 못했습니다.')); }
      };
      fr.onerror = function () { reject(new Error(file.name + ' — 파일을 열지 못했습니다.')); };
      fr.readAsArrayBuffer(file);
    });
  }
  function fileBtn(label, accept, onFile, cls) {
    var f = h('input', { type: 'file', accept: accept });
    f.addEventListener('change', function () { var file = f.files[0]; f.value = ''; if (file) onFile(file); });
    return h('label', { class: 'btn' + (cls ? ' ' + cls : '') }, f, label);
  }
  function copyText(text, okMsg) {
    function fallback() {
      var ta = h('textarea', { style: 'position:fixed;left:-9999px' }); ta.value = text; document.body.appendChild(ta); ta.select();
      try { document.execCommand('copy'); toast(okMsg || '복사했습니다.'); } catch (e) { toast('복사하지 못했습니다. 글을 직접 골라 복사해 주세요.', true); }
      ta.remove();
    }
    if (navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(text).then(function () { toast(okMsg || '복사했습니다.'); }, fallback);
    else fallback();
  }
  function n0(v) { return L.fmtN(v); }
  function ctText(sec) { return sec == null ? '-' : L.fmtDur(sec); }
  function pageHead(stage, title, lead, right) {
    return h('div', { class: 'page-head' }, h('div', { class: 'titles' }, h('div', { class: 'stage' }, stage), h('h1', null, title), lead ? h('p', { class: 'lead' }, lead) : null), right || null);
  }
  function tile(k, v, sub, cls) { return h('div', { class: 'tile' + (cls ? ' ' + cls : '') }, h('div', { class: 'k' }, k), h('div', { class: 'v' }, v), sub ? h('div', { class: 'note' }, sub) : null); }
  function table(head, rows, opts) {
    opts = opts || {};
    var t = h('table', { class: 'list' + (opts.cls ? ' ' + opts.cls : '') });
    var tr = h('tr');
    head.forEach(function (x, i) { tr.appendChild(h('th', { class: opts.num && opts.num[i] ? 'num' : null, scope: 'col' }, x)); });
    t.appendChild(h('thead', null, tr));
    var tb = h('tbody');
    rows.forEach(function (r) {
      var row = h('tr', { class: r.cls || null });
      (r.cells || r).forEach(function (x, i) { row.appendChild(h('td', { class: (opts.num && opts.num[i] ? 'num' : '') || null }, x)); });
      tb.appendChild(row);
    });
    if (!rows.length) tb.appendChild(h('tr', null, h('td', { colspan: head.length, class: 'note' }, opts.empty || '없음')));
    t.appendChild(tb);
    return h('div', { class: 'table-wrap' }, t);
  }
  var TAG = { '경고': 'bad', '주의': 'warn', '정보': 'info', '많음': 'warn', '적음': 'bad', '작업 중': 'ok', '멈춤': 'warn' };
  function tag(text) { return h('span', { class: 'tag ' + (TAG[text] || 'info') }, text); }
  function itemLabel(it) { return it ? it.part + (it.product ? ' · ' + it.product : '') + (it.customer ? ' (' + it.customer + ')' : '') : '(지운 품번)'; }
  function byId() { return L.itemById(db); }
  function analyze() { return L.analyze(db, Object.assign({}, pref.filter, ui.from != null ? { from: ui.from } : {})); }
  function needItems() {
    return h('section', { class: 'card empty' },
      h('p', null, '아직 관리 아이템이 없습니다.'),
      h('p', null, '설정에서 아이템(고객사·품번·공정 순서)을 넣고, 기초재고를 한 번 적으면 실적이 쌓일 때마다 공정별 재고가 바뀝니다. 먼저 가상 샘플로 둘러봐도 됩니다.'),
      h('div', { class: 'btn-row', style: 'justify-content:center' },
        h('button', { type: 'button', class: 'btn btn-primary', onclick: loadSample }, '가상 샘플로 시작'),
        h('a', { class: 'btn', href: '#/settings' }, '설정에서 아이템 넣기')));
  }
  /* 옛 판(1·2차) 자료를 열었을 때 한 번 보이는 안내 */
  function oldNotice() {
    var n = db.notice; if (!n) return null;
    var old = S.loadOld();
    return h('section', { class: 'card', id: 'oldNotice' }, h('h2', null, n.broken ? '저장된 자료를 읽지 못했습니다' : '저장 형식이 바뀌었습니다 (3차 범위 재설정)'),
      n.broken ? h('p', null, '이 브라우저에 저장된 자료가 깨져 있어 빈 상태로 시작했습니다. 원래 글은 아래 단추로 받아 둘 수 있습니다.') :
        h('p', null, '이 도구는 이제 「기초재고 + 공정 실적 → 공정별 재고현황」만 계산합니다. 전에 넣은 발주서·납품·재고 체크·실적(' + n.had + '건)은 새 계산에 쓰지 않습니다' +
          (n.sample ? ' — 옛 가상 샘플이었으므로 새 가상 샘플로 다시 시작하시면 됩니다.' : '. 공정 이름과 모델 ' + n.items + '개는 관리 아이템으로 옮겼습니다(공정 순서는 모든 공정으로 넣었으니 설정에서 고쳐 주세요).')),
      h('div', { class: 'btn-row' },
        old ? h('button', { type: 'button', class: 'btn', onclick: function () { download('data10-05_옛자료_' + L.todayIso() + '.json', new Blob([old], { type: 'application/json' })); } }, '옛 자료 JSON 받기') : null,
        n.sample || !db.items.length ? h('button', { type: 'button', class: 'btn btn-primary', onclick: loadSample }, '새 가상 샘플로 시작') : null,
        h('button', { type: 'button', class: 'btn', onclick: function () { db.notice = null; save(); render(); } }, '알겠습니다(안내 닫기)')));
  }
  function filterBar() {
    var opts = L.filterOptions(db.items), labels = { customer: '고객사', group: '그룹', subgroup: '상세그룹', product: '품목' };
    return h('div', { class: 'filters card no-print' }, L.FILTER_KEYS.map(function (k) {
      var sel = h('select', { 'aria-label': labels[k] }, h('option', { value: '' }, '모두'), opts[k].map(function (v) { var o = h('option', { value: v }, v); if (pref.filter[k] === v) o.selected = true; return o; }));
      sel.addEventListener('change', function () { pref.filter[k] = sel.value; S.saveUi(pref); render(); });
      return fld(labels[k], sel);
    }));
  }

  /* ── 0. 대시보드 ───────────────────────────── */
  function pageDash() {
    var head = pageHead('대시보드', '공정별 재고현황', null);
    if (!db.items.length) return [head, oldNotice(), needItems()];
    var r = analyze(), k = r.kpi, by = byId();
    head = pageHead('대시보드', '공정별 재고현황',
      '기초재고 ' + (r.openingDate ? L.fmtDate(r.openingDate) + ' 아침' : '(기준날짜 없음)') + ' + 그 뒤 공정 실적 ' + n0(r.resultsAfter) + '건으로 계산한 지금 재고입니다.',
      h('div', { class: 'btn-row no-print' }, h('button', { type: 'button', class: 'btn btn-primary', onclick: function () { download('공정재고_' + L.todayIso() + '.xlsx', xlsxBlob(L.exportSheets(r, db))); } }, '엑셀로 내보내기')));
    var out = [head, oldNotice(), filterBar()];
    out.push(h('div', { class: 'tiles' },
      tile('관리 아이템', k.items + '개', '거르기 결과'),
      tile('공정 완료재고 합계', n0(k.wip) + '개', '출하 전 모든 공정'),
      tile('출하 누계', n0(k.shipped) + '개', r.openingDate ? r.openingDate + '부터' : '모든 실적'),
      tile('평균에서 ±30% 벗어난 날', k.flagged + '건', '작업 중 타이머 ' + k.timers + '개', k.flagged ? 'bad' : '')));
    if (r.checks.length) out.push(h('section', { class: 'card' }, h('h2', null, '확인할 것 ' + r.checks.length + '건'),
      h('div', { class: 'an-list' }, r.checks.map(function (c) { return h('div', { class: 'an lv-' + (TAG[c.level] || 'info') }, h('div', { class: 'an-top' }, tag(c.level), h('a', { class: 'an-title', href: c.href }, c.text))); }))));
    // 공정별 · 모델별 재고
    var procs = r.processes;
    var rows = r.stock.map(function (o) {
      return { cls: o.negatives.length ? 'neg' : null, cells: [o.item.customer, o.item.part, o.item.product].concat(procs.map(function (p) { return o.item.route.indexOf(p) < 0 ? '' : n0(o.stock[p]); }), [n0(o.wip), n0(o.shipped)]) };
    });
    rows.push({ cls: 'sum', cells: ['합계', '', ''].concat(r.procStock.map(function (p) { return n0(p.qty); }), [n0(k.wip), n0(k.shipped)]) });
    var numCols = [0, 0, 0].concat(procs.map(function () { return 1; }), [1, 1]);
    out.push(h('section', { class: 'card' }, h('h2', null, '공정별 · 모델별 완료재고'),
      h('div', { class: 'charts' },
        C.figure('공정별 완료재고 합계', { type: 'bar', unit: '개', labels: r.procStock.map(function (p) { return p.process; }), series: [{ name: '완료재고', values: r.procStock.map(function (p) { return p.qty; }), color: 'var(--series-1)' }] }),
        C.figure('모델(품번)별 공정 재고 합계', { type: 'bar', unit: '개', labels: r.stock.map(function (o) { return o.item.part; }), series: [{ name: '재공 합계', values: r.stock.map(function (o) { return o.wip; }), color: 'var(--series-3)' }] })),
      table(['고객사', '품번', '품목'].concat(procs, ['재공 합계', '출하 누계']), rows, { num: numCols, empty: '거르기에 맞는 아이템이 없습니다.' }),
      h('p', { class: 'note' }, '빈 칸 = 그 품번이 거치지 않는 공정. 공정 P 실적이 들어오면 P 완료재고는 늘고 바로 앞 공정 완료재고는 줄어듭니다. 「출하」는 모든 품번의 마지막 공정이라 출하 실적은 마지막 공정 재고에서 빠집니다.'),
      calcCard(r)));
    // C/T
    var ctRows = r.ct.slice().sort(function (a, b) {
      var A = by[a.itemId] || { part: '', route: [] }, B = by[b.itemId] || { part: '', route: [] };
      return A.part !== B.part ? (A.part < B.part ? -1 : 1) : L.flowOf(A).indexOf(a.process) - L.flowOf(B).indexOf(b.process);
    }).map(function (c) {
      var pa = r.ctProc.filter(function (x) { return x.process === c.process; })[0], slow = pa && pa.ct && c.ct >= pa.ct * (1 + L.DEV_PCT);
      return { cls: slow ? 'differs' : null, cells: [(by[c.itemId] || {}).part || '?', c.process, c.n, n0(c.qty), ctText(c.sec), ctText(c.ct), ctText(c.min) + ' ~ ' + ctText(c.max), slow ? tag('평균보다 30%↑') : ''] };
    });
    var ctp = r.ctProc.filter(function (c) { return c.n; });
    out.push(h('section', { class: 'card' }, h('h2', null, '실제 걸린 시간 (C/T)'),
      h('p', { class: 'note' }, 'C/T(개당 걸린 시간) = 순작업 시간(시작~종료에서 멈춘 사이를 뺀 시간) ÷ 수량. 작업 인원 수는 따지지 않습니다(여러 명이 함께 해도 그 작업에 걸린 시간 그대로). 시간이 있는 실적만 씁니다(기초재고 기준날짜 전 실적 포함).'),
      h('div', { class: 'charts one' }, C.figure('공정별 평균 C/T (초/개)', { type: 'bar', unit: '초', labels: ctp.map(function (c) { return c.process; }), series: [{ name: 'C/T', values: ctp.map(function (c) { return L.round(c.ct, 1); }), color: 'var(--series-1)' }] })),
      table(['품번', '공정', '실적 건수', '수량', '순작업 합계', 'C/T(개당)', '가장 짧은 ~ 긴', ''], ctRows, { num: [0, 0, 1, 1, 1, 1, 0, 0], empty: '시간이 있는 실적이 없습니다 — 실적 입력의 시작·종료 단추로 재면 쌓입니다.' })));
    return out;
  }
  /* 계산 보기 — 한 품번의 공정마다 기초재고 + 이 공정 실적 − 다음 공정이 가져감 = 완료재고 */
  function calcCard(r) {
    if (!r.stock.length) return null;
    var o = r.stock.filter(function (x) { return x.item.id === ui.calcItem; })[0] || r.stock[0];
    var sel = h('select', { 'aria-label': '계산을 볼 품번' }, r.stock.map(function (x) { var op = h('option', { value: x.item.id }, itemLabel(x.item)); if (x === o) op.selected = true; return op; }));
    sel.addEventListener('change', function () { ui.calcItem = sel.value; render(); });
    var flow = o.flow, rows = o.item.route.map(function (p, i) {
      return { cls: o.stock[p] < 0 ? 'neg' : null, cells: [p, n0(o.opening[p]), '+ ' + n0(o.done[p]), '− ' + n0(o.used[p]) + ' (' + flow[i + 1] + ')', n0(o.stock[p])] };
    });
    rows.push({ cls: 'sum', cells: ['출하', '', '+ ' + n0(o.shipped) + ' (출하 누계)', '', ''] });
    return h('details', { class: 'block', open: true }, h('summary', null, '계산 보기 — 숫자가 어떻게 바뀌었는지'),
      fld('품번', sel),
      table(['공정', '기초재고', '이 공정 실적', '다음 공정이 가져감', '= 완료재고'], rows, { num: [0, 1, 1, 1, 1] }),
      h('p', { class: 'note' }, '첫 공정(' + o.item.route[0] + ')은 앞 공정이 없으므로 원자재는 빼지 않습니다(자재 재고는 이 도구에서 관리하지 않음).'));
  }

  /* ── 1. 실적 입력 (타이머 · 직접 적기) ───────── */
  function pageEntry() {
    var head = pageHead('① 실적 입력', '공정 실적', '아이템과 공정을 고르고 시작 → (멈춤 · 다시 시작) → 종료 때 수량을 적으면 실적 한 건이 됩니다. 출하도 같은 방법으로 「출하」 공정 실적으로 적습니다. 새로고침하거나 화면을 꺼도 타이머는 이어집니다.');
    if (!db.items.length) return [head, needItems()];
    var tabs = h('div', { class: 'seg', role: 'tablist', style: 'margin-bottom:14px' }, [['timer', '타이머로 재기'], ['manual', '직접 적기']].map(function (x) {
      var r = h('input', { type: 'radio', name: 'tab', value: x[0] }); if (ui.tab === x[0]) r.checked = true;
      r.addEventListener('change', function () { ui.tab = x[0]; render(); });
      return h('label', { class: 'pill' }, r, h('span', null, x[1]));
    }));
    return [head, h('div', { class: 'entry' }, tabs, ui.tab === 'manual' ? manualCard() : startCard(), ui.tab === 'manual' ? null : timersCard(), recentCard())];
  }
  function itemSelect(val, label) {
    var sel = h('select', { 'aria-label': label || '아이템' }, h('option', { value: '' }, '— 골라 주세요 —'), db.items.map(function (it) { var o = h('option', { value: it.id }, itemLabel(it)); if (it.id === val) o.selected = true; return o; }));
    return sel;
  }
  function procPills(item, val, name) {
    var flow = item ? L.flowOf(item) : [];
    if (!item) return h('p', { class: 'note' }, '아이템을 고르면 그 아이템의 공정 순서가 나옵니다.');
    return h('div', { class: 'seg', role: 'radiogroup', 'aria-label': '공정' }, flow.map(function (p, i) {
      var r = h('input', { type: 'radio', name: name, value: p }); if (p === val) r.checked = true;
      return h('label', { class: 'pill' }, r, h('span', null, (i + 1) + '. ' + p));
    }));
  }
  function startCard() {
    var it = byId()[pref.itemId];
    var sel = itemSelect(it ? it.id : '');
    sel.addEventListener('change', function () { pref.itemId = sel.value; S.saveUi(pref); render(); });
    var pv = it && L.flowOf(it).indexOf(pref.process) >= 0 ? pref.process : '';
    var pills = procPills(it, pv, 'tproc');
    function start() {
      var chk = pills.querySelector && pills.querySelector('input:checked');
      if (!it) { toast('아이템을 골라 주세요.', true); return; }
      if (!chk) { toast('공정을 골라 주세요.', true); return; }
      pref.process = chk.value; S.saveUi(pref);
      db.timers.push(L.timerStart(it.id, chk.value, Date.now()));
      save(); render(); toast(it.part + ' · ' + chk.value + ' 시작했습니다.');
    }
    return h('section', { class: 'card' }, h('h2', null, '새 작업 시작'),
      fld('아이템', sel), h('div', { class: 'field', style: 'margin:12px 0' }, h('span', null, '공정'), pills),
      h('button', { type: 'button', class: 'btn btn-primary btn-big btn-block', onclick: start }, '▶ 시작 (start)'));
  }
  function timersCard() {
    var by = byId(), now = Date.now();
    return h('section', { class: 'card' }, h('h2', null, '진행 중인 작업 ' + db.timers.length + '개'),
      db.timers.length ? h('div', { class: 'entry-list' }, db.timers.map(function (t) {
        var it = by[t.itemId], run = L.timerRunning(t);
        var qty = h('input', { inputmode: 'numeric', placeholder: '수량', 'aria-label': '완료 수량', style: 'max-width:110px' });
        function act(fn, msg) { return function () { fn(); save(); render(); if (msg) toast(msg); }; }
        function finish() {
          var f = L.timerFinish(t, qty.value, Date.now());
          if (!f.ok) { toast(f.error, true); qty.focus(); return; }
          db.results.push(f.result); db.timers = db.timers.filter(function (x) { return x !== t; });
          save(); render(); toast('실적 저장 — ' + (it ? it.part : '') + ' ' + t.process + ' ' + f.result.qty + '개 · 순작업 ' + L.fmtDur(f.result.runSec) + ' · C/T ' + L.fmtDur(f.result.runSec / f.result.qty));
        }
        return h('div', { class: 'timer-card' + (run ? ' run' : ' paused') },
          h('div', { class: 'an-top' }, tag(run ? '작업 중' : '멈춤'), h('b', null, itemLabel(it)), h('span', { class: 'note' }, '· ' + t.process)),
          h('div', { class: 'timer-clock', 'data-timer': t.id }, L.fmtClock(L.timerNetMs(t, now) / 1000)),
          h('div', { class: 'note' }, '시작 ' + L.hhmm(t.startedAt) + ' · 멈춘 시간 ', h('span', { 'data-paused': t.id }, L.fmtDur(L.timerPausedMs(t, now) / 1000)), ' · 멈춤 ' + (t.segs.length - (run ? 1 : 0)) + '회'),
          h('div', { class: 'btn-row', style: 'margin-top:8px' },
            run ? h('button', { type: 'button', class: 'btn', onclick: act(function () { L.timerStop(t, Date.now()); }) }, '❚❚ 멈춤 (stop)')
              : h('button', { type: 'button', class: 'btn', onclick: act(function () { L.timerRestart(t, Date.now()); }) }, '▶ 다시 시작 (restart)'),
            qty, h('button', { type: 'button', class: 'btn btn-primary', onclick: finish }, '■ 종료 (end)'),
            h('button', { type: 'button', class: 'btn btn-sm', onclick: function () { if (!confirm('이 작업을 실적 없이 지울까요?')) return; db.timers = db.timers.filter(function (x) { return x !== t; }); save(); render(); } }, '취소')));
      })) : h('p', { class: 'note' }, '진행 중인 작업이 없습니다.'));
  }
  function manualCard() {
    var E = ui.manualErr || {}, it = byId()[pref.itemId];
    var dateIn = h('input', { type: 'date', value: pref.manualDate || L.todayIso() });
    var sel = itemSelect(it ? it.id : '');
    sel.addEventListener('change', function () { pref.itemId = sel.value; pref.manualDate = dateIn.value; S.saveUi(pref); render(); });
    var pills = procPills(it, it && L.flowOf(it).indexOf(pref.process) >= 0 ? pref.process : '', 'mproc');
    var qty = h('input', { inputmode: 'numeric', placeholder: '개' }), min = h('input', { inputmode: 'decimal', placeholder: '비워도 됨' }), note = h('input', { placeholder: '설비 정지·자재 대기 등(선택)' });
    function submit() {
      var chk = pills.querySelector && pills.querySelector('input:checked');
      var v = L.validateManual({ date: dateIn.value, itemId: sel.value, process: chk ? chk.value : '', qty: qty.value, minutes: min.value, note: note.value }, db);
      if (!v.ok) { ui.manualErr = v.errors; render(); toast('빈 칸이나 잘못된 값이 있습니다.', true); return; }
      ui.manualErr = {}; db.results.push(v.result);
      pref.process = v.result.process; pref.manualDate = v.result.date; S.saveUi(pref);
      save(); render(); toast('실적을 저장했습니다.');
    }
    return h('section', { class: 'card' }, h('h2', null, '직접 적기 (타이머 없이)'),
      h('p', { class: 'note' }, '끝난 작업을 나중에 적을 때 씁니다. 시간(분)을 적으면 C/T 에도 들어갑니다.'),
      h('div', { class: 'entry-grid' },
        fld('일자', dateIn, null, E.date), h('div', { class: 'span-all' }, fld('아이템', sel, null, E.itemId)),
        h('div', { class: 'field span-all' + (E.process ? ' invalid' : '') }, h('span', null, '공정'), pills, E.process ? h('small', { class: 'err' }, E.process) : null),
        fld('수량', qty, null, E.qty), fld('순작업 시간(분)', min, null, E.minutes), h('div', { class: 'span-all' }, fld('비고', note))),
      h('button', { type: 'button', class: 'btn btn-primary btn-big btn-block', onclick: submit }, '실적 저장'));
  }
  function recentCard() {
    var by = byId(), list = db.results.slice().sort(function (a, b) { return (b.end || 0) - (a.end || 0) || (b.date > a.date ? 1 : -1); }).slice(0, 12);
    return h('section', { class: 'card' }, h('h2', null, '최근 실적'),
      table(['일자', '품번', '공정', '수량', '순작업', 'C/T', '경로', ''], list.map(function (r) {
        return [L.shortDate(r.date), (by[r.itemId] || {}).part || '?', r.process, n0(r.qty), r.runSec == null ? '-' : L.fmtDur(r.runSec), r.runSec == null ? '-' : L.fmtDur(r.runSec / r.qty), r.source || '',
          h('button', { type: 'button', class: 'btn btn-sm', onclick: function () { if (!confirm(r.date + ' ' + r.process + ' ' + r.qty + '개 실적을 지울까요?')) return; db.results = db.results.filter(function (x) { return x !== r; }); save(); render(); } }, '지우기')];
      }), { num: [0, 0, 0, 1, 1, 1, 0, 0], empty: '아직 실적이 없습니다.' }),
      h('div', { class: 'btn-row', style: 'margin-top:8px' }, h('button', { type: 'button', class: 'btn btn-sm', onclick: function () { download('공정실적_' + L.todayIso() + '.xlsx', xlsxBlob({ '공정 실적': L.resultSheet(db) })); } }, '모든 실적 엑셀로 받기')));
  }
  /* 화면을 다시 그리지 않고 타이머 숫자만 1초마다 */
  setInterval(function () {
    var now = Date.now();
    db.timers.forEach(function (t) {
      var a = main.querySelector('[data-timer="' + t.id + '"]'); if (a) a.textContent = L.fmtClock(L.timerNetMs(t, now) / 1000);
      var b = main.querySelector('[data-paused="' + t.id + '"]'); if (b) b.textContent = L.fmtDur(L.timerPausedMs(t, now) / 1000);
    });
  }, 1000);

  /* ── 2. 기초재고 ─────────────────────────────── */
  function pageOpening() {
    var head = pageHead('② 데이터 가져오기', '기초재고', '한 번만 적으면 됩니다. 기준날짜 아침(그날 작업 시작 전)의 공정별 완료재고를 적으면, 그날부터의 실적 누계로 지금 재고를 계산합니다. 언제든 고칠 수 있습니다.');
    if (!db.items.length) return [head, oldNotice(), needItems()];
    var by = L.openingIndex(db.opening), procs = db.processes;
    var dateIn = h('input', { type: 'date', value: db.opening.date || '' });
    var inputs = [];
    var rows = db.items.map(function (it) {
      return [it.part, it.product].concat(procs.map(function (p) {
        if (it.route.indexOf(p) < 0) return h('span', { class: 'note' }, '-');
        var inp = h('input', { inputmode: 'numeric', value: by[it.id + '|' + p] || '', placeholder: '0', 'aria-label': it.part + ' ' + p, style: 'width:72px;text-align:right' });
        inputs.push({ it: it, p: p, inp: inp });
        return inp;
      }));
    });
    function saveOpening() {
      var d = L.parseDate(dateIn.value);
      if (!d) { toast('기준날짜를 골라 주세요.', true); dateIn.focus(); return; }
      var out = [], bad = [];
      inputs.forEach(function (x) { var q = L.parseQty(x.inp.value); if (q == null) bad.push(x.it.part + ' ' + x.p); else if (q) out.push({ itemId: x.it.id, process: x.p, qty: q }); });
      if (bad.length) { toast('숫자가 아닌 칸: ' + bad.slice(0, 3).join(', '), true); return; }
      db.opening = { date: d, rows: out }; save(); render(); toast('기초재고를 저장했습니다 — ' + d + ' 아침 기준.');
    }
    function importOpening(file) {
      readSheets(file).then(function (sheets) {
        var best = null;
        sheets.forEach(function (s) { var o = L.parseOpeningSheet(s.aoa, db); if (!best || o.rows.length > best.rows.length) best = o; });
        if (!best || !best.rows.length && best.errors.length) { toast((best && best.errors[0]) || '읽을 시트가 없습니다.', true); return; }
        var d = best.date || L.parseDate(dateIn.value);
        if (!d) { toast('파일에 「기준날짜」가 없습니다 — 화면에서 기준날짜를 먼저 고른 뒤 다시 넣어 주세요.', true); return; }
        var msg = [file.name + ' — ' + best.rows.length + '칸, 기준날짜 ' + d];
        if (best.unknownParts.length) msg.push('설정에 없는 품번(건너뜀): ' + best.unknownParts.slice(0, 5).join(', ') + (best.unknownParts.length > 5 ? ' 외' : ''));
        if (best.offRoute.length) msg.push('공정 순서에 없는 칸(건너뜀): ' + best.offRoute.slice(0, 5).join(', '));
        best.errors.forEach(function (e) { msg.push(e); });
        if (!confirm(msg.join('\n') + '\n\n지금 기초재고를 이 파일 내용으로 바꿀까요?')) return;
        db.opening = { date: d, rows: best.rows.filter(function (r) { return r.qty; }) }; db.sample = false; save(); render(); toast('기초재고를 가져왔습니다.');
      }, function (e) { toast(e.message, true); });
    }
    var after = L.resultsAfter(db.results, db.opening.date).length, before = db.results.length - after;
    return [head, oldNotice(),
      h('section', { class: 'card' }, h('h2', null, '기준날짜와 공정별 완료재고'),
        h('div', { class: 'form-grid' }, fld('기준날짜', dateIn, '이 날 아침 재고. 이 날부터의 실적이 재고에 더해집니다.')),
        h('div', { style: 'margin-top:12px' }, table(['품번', '품목'].concat(procs), rows, { num: [0, 0].concat(procs.map(function () { return 1; })) })),
        h('p', { class: 'note' }, '「-」 = 그 품번이 거치지 않는 공정. 출하는 재고가 아니라 누계라 적지 않습니다.'),
        h('div', { class: 'btn-row', style: 'margin-top:10px' }, h('button', { type: 'button', class: 'btn btn-primary', onclick: saveOpening }, '기초재고 저장')),
        db.opening.date ? h('p', { class: 'note' }, '지금: ' + L.fmtDate(db.opening.date) + ' 아침 기준 · 그 뒤 실적 ' + n0(after) + '건이 재고에 들어감' + (before ? ' · 그 전 실적 ' + before + '건은 기초재고에 이미 들어 있다고 보고 뺌(C/T 에는 씀)' : '')) : h('p', { class: 'alert info' }, '기준날짜가 없으면 모든 실적을 0 에서부터 더합니다.')),
      h('section', { class: 'card' }, h('h2', null, '엑셀로 넣기 · 받기'),
        h('p', null, '양식을 받아 칸을 채운 뒤 다시 넣으면 됩니다. 1행에 「기준날짜」, 3행이 머리행(품번 · 품목 · 공정 이름들)입니다.'),
        h('div', { class: 'btn-row' },
          h('button', { type: 'button', class: 'btn', onclick: function () { download('기초재고_양식_' + L.todayIso() + '.xlsx', xlsxBlob({ '기초재고': L.openingSheet(db) })); } }, '기초재고 양식 받기(지금 값 포함)'),
          fileBtn('기초재고 엑셀 넣기', '.xlsx,.xls,.xlsm', importOpening, 'btn-primary'))),
      h('section', { class: 'card' }, h('h2', null, '이 도구가 받지 않는 자료'),
        h('p', { class: 'note' }, '3차(10-27 요청)부터 고객사 발주서·납품이력·실적 엑셀 가져오기는 뺐습니다. 실적은 ① 실적 입력에서 쌓고, 아이템은 설정에서 엑셀로 넣습니다.'))];
  }

  /* ── 3. 공정 분석 — 공정별 일별 생산수량 ───────── */
  function pageProcess() {
    var head = pageHead('③ 공정 분석', '공정별 일별 생산수량', '실적이 있는 날의 평균과, 평균에서 30% 이상 벗어난 날을 보여 줍니다.');
    if (!db.items.length) return [head, needItems()];
    var r = analyze();
    var fromIn = h('input', { type: 'date', value: r.from || '' });
    fromIn.addEventListener('change', function () { ui.from = fromIn.value || ''; render(); });
    var cur = r.daily.filter(function (d) { return d.process === ui.proc; })[0] || r.daily.filter(function (d) { return d.flagged.length; })[0] || r.daily.filter(function (d) { return d.days.length; })[0] || r.daily[0];
    var pills = h('div', { class: 'seg', role: 'radiogroup', 'aria-label': '공정' }, r.daily.map(function (d) {
      var x = h('input', { type: 'radio', name: 'aproc', value: d.process }); if (d === cur) x.checked = true;
      x.addEventListener('change', function () { ui.proc = d.process; render(); });
      return h('label', { class: 'pill' }, x, h('span', null, d.process + (d.flagged.length ? ' ⚑' + d.flagged.length : '')));
    }));
    var out = [head, filterBar(),
      h('section', { class: 'card' }, h('h2', null, '공정마다 한눈에'),
        h('div', { class: 'form-grid' }, fld('이 날부터', fromIn, '기본 = 기초재고 기준날짜. 비우면 모든 날')),
        table(['공정', '실적 있는 날', '일 평균', '가장 적은 날', '가장 많은 날', '±30% 벗어난 날'], r.daily.map(function (d) {
          return { cls: d.flagged.length ? 'differs' : null, cells: [d.process, d.days.length + '일', L.fmtN(d.avg, 1), n0(d.min), n0(d.max), d.flagged.length ? d.flagged.map(function (x) { return L.shortDate(x.date) + ' ' + x.flag; }).join(', ') : '없음'] };
        }), { num: [0, 1, 1, 1, 1, 0] }))];
    if (cur) {
      var by = byId();
      out.push(h('section', { class: 'card' }, h('h2', null, cur.process + ' — 일별 생산수량'), pills,
        h('div', { class: 'charts one', style: 'margin-top:12px' }, C.figure(cur.process + ' 일별 생산수량 (빨강·주황 = 평균에서 30% 이상 벗어남)', {
          type: 'bar', unit: '개', labels: cur.days.map(function (d) { return L.shortDate(d.date); }),
          series: [{ name: '생산수량', values: cur.days.map(function (d) { return d.qty; }), color: 'var(--series-1)', colors: cur.days.map(function (d) { return d.flag === '적음' ? 'var(--danger)' : d.flag === '많음' ? 'var(--series-2)' : null; }) }],
          refs: cur.avg ? [{ value: L.round(cur.avg, 1), name: '평균', color: 'var(--series-plan)' }] : []
        })),
        table(['일자', '수량', '실적 건수', '평균 대비', '판정'], cur.days.map(function (d) {
          return { cls: d.flag ? 'differs' : null, cells: [L.fmtDate(d.date), n0(d.qty), d.n, d.dev == null ? '-' : (d.dev > 0 ? '+' : '') + L.round(d.dev * 100, 1) + '%', d.flag ? tag(d.flag) : ''] };
        }), { num: [0, 1, 1, 1, 0], empty: '이 기간에 실적이 없습니다.' }),
        cur.flagged.length ? h('details', { class: 'block', open: true }, h('summary', null, '벗어난 날의 실적 ' + cur.flagged.length + '일'),
          table(['일자', '품번', '수량', '시작~종료', '순작업', '비고'], [].concat.apply([], cur.flagged.map(function (d) {
            return d.rows.map(function (x) { return [x.date, (by[x.itemId] || {}).part || '?', n0(x.qty), x.start ? L.hhmm(x.start) + '~' + L.hhmm(x.end) : '-', x.runSec == null ? '-' : L.fmtDur(x.runSec), x.note || '']; });
          })), { num: [0, 0, 1, 0, 1, 0] })) : null));
    }
    return out;
  }

  /* ── 4. AI 분석 ─────────────────────────────── */
  function pageAi() {
    var head = pageHead('④ AI', 'AI 분석 (반자동)', '도구가 계산한 공정 재고 · C/T · 벗어난 날만 담은 프롬프트를 만듭니다. 복사해 회사가 허용한 AI 에 붙여 넣고, 답을 아래에 붙여 넣어 주세요.');
    if (!db.items.length) return [head, needItems()];
    var r = analyze();
    db.aliases.items = L.assignAliases(db.aliases.items, r.items.map(function (x) { return x.part; }), L.ITEM_ALIAS);
    db.aliases.customers = L.assignAliases(db.aliases.customers, r.items.map(function (x) { return x.customer; }).filter(Boolean), L.CUST_ALIAS);
    var info = L.buildAiPrompt(r, db, { maskItems: db.maskItems, maskCustomers: db.maskCustomers, aliases: db.aliases });
    function opt(label, sub, checked, on) {
      var b = h('input', { type: 'checkbox' }); b.checked = checked;
      b.addEventListener('change', function () { on(b.checked); save(); render(); });
      return h('label', { class: 'opt' }, b, h('span', null, h('span', { class: 't' }, label), h('span', { class: 's' }, sub)));
    }
    var ta = h('textarea', { readonly: true, rows: 14, class: 'mono', 'aria-label': 'AI 프롬프트' }); ta.value = info.prompt;
    var ans = h('textarea', { rows: 8, placeholder: 'AI 답을 여기에 붙여 넣어 주세요. 별칭(품목1·거래처1)은 실제 품번·고객사로 되돌려 보여 줍니다.', 'aria-label': 'AI 답' }); ans.value = db.aiNote || '';
    function back(t) { return L.unmaskText(t, db.aliases); }
    var usedItems = Object.keys(info.used.items), usedCust = Object.keys(info.used.customers);
    return [head,
      h('section', { class: 'card' }, h('h2', null, '1. 보안 선택'),
        h('div', { class: 'mask-box' },
          opt('품번 가리기(「품목1」 같은 별칭)', '기본값. 대응표는 이 브라우저에만 있습니다.', db.maskItems, function (v) { db.maskItems = v; }),
          opt('고객사 가리기(「거래처1」 같은 별칭)', '기본값. 고객사 이름을 외부 AI 에 보내지 않습니다.', db.maskCustomers, function (v) { db.maskCustomers = v; })),
        info.leaks.length ? h('p', { class: 'alert warn' }, '가리지 못한 이름이 프롬프트에 남아 있습니다: ' + info.leaks.join(', ') + ' — 보내기 전에 지워 주세요.') : null,
        usedItems.length || usedCust.length ? h('details', null, h('summary', null, '이 프롬프트의 별칭 대응표 (이 PC 에만 있음)'),
          table(['실제', '별칭'], usedItems.map(function (k) { return [k, info.used.items[k]]; }).concat(usedCust.map(function (k) { return [k, info.used.customers[k]]; })))) : null),
      h('section', { class: 'card' }, h('h2', null, '2. 프롬프트'), ta,
        h('div', { class: 'btn-row', style: 'margin-top:8px' },
          h('button', { type: 'button', class: 'btn btn-primary', onclick: function () { copyText(info.prompt, '프롬프트를 복사했습니다.'); } }, '프롬프트 복사'),
          AIP.sendButton('설정한 AI 서버로 보내기', function () { return info.prompt; }, function (text) { db.aiNote = text; save(); render(); toast('AI 답을 받았습니다.'); }, { toast: toast, settingsHref: '#/settings' }))),
      h('section', { class: 'card' }, h('h2', null, '3. AI 답'), ans,
        h('div', { class: 'btn-row', style: 'margin-top:8px' },
          h('button', { type: 'button', class: 'btn', onclick: function () { db.aiNote = ans.value; save(); render(); toast('답을 저장했습니다.'); } }, '답 저장'),
          db.aiNote ? h('button', { type: 'button', class: 'btn', onclick: function () { copyText(back(db.aiNote), '실제 이름으로 되돌린 답을 복사했습니다.'); } }, '되돌린 답 복사') : null),
        db.aiNote ? h('div', { class: 'rp' }, h('h3', { style: 'margin-top:14px' }, '실제 품번·고객사로 되돌린 답'), h('div', { class: 'pre' }, back(db.aiNote))) : null,
        h('p', { class: 'note' }, 'AI 답은 초안입니다. 숫자는 대시보드·공정 분석 표와 대조해 확인해 주세요.'))];
  }

  /* ── 설정 ─────────────────────────────────── */
  function pageSettings() {
    return [pageHead('설정', '관리 아이템 · 공정', '아이템(고객사 · 그룹 · 상세그룹 · 품목 · 품번 · 비고)과 아이템마다 공정 순서를 정합니다. 「출하」는 모든 아이템의 마지막 공정으로 자동으로 붙습니다.'),
      oldNotice(), processCard(), itemsCard(),
      AIP.settingsCard({ toast: toast }),
      h('section', { class: 'card' }, h('h2', null, '백업'),
        h('div', { class: 'btn-row' },
          h('button', { type: 'button', class: 'btn', onclick: function () { download('data10-05_백업_' + L.todayIso() + '.json', new Blob([JSON.stringify(db)], { type: 'application/json' })); } }, 'JSON 백업 받기'),
          fileBtn('JSON 백업 불러오기', '.json', function (file) {
            var fr = new FileReader();
            fr.onload = function () { try { var o = JSON.parse(fr.result); if (!confirm('지금 데이터를 백업 파일 내용으로 바꿀까요?')) return; db = L.restoreDb(o); save(); render(); toast('백업을 불러왔습니다.'); } catch (e) { toast('백업 파일을 읽지 못했습니다.', true); } };
            fr.readAsText(file);
          }),
          S.loadOld() ? h('button', { type: 'button', class: 'btn', onclick: function () { download('data10-05_옛자료_' + L.todayIso() + '.json', new Blob([S.loadOld()], { type: 'application/json' })); } }, '옛 판(1·2차) 자료 받기') : null,
          h('button', { type: 'button', class: 'btn btn-danger', onclick: function () { if (!confirm('이 브라우저의 아이템·기초재고·실적·타이머를 모두 지울까요? (먼저 JSON 백업을 받아 두면 안전합니다)')) return; db = L.emptyDb(); save(); render(); toast('모두 지웠습니다.'); } }, '모두 지우기')),
        h('p', { class: 'note' }, '자료는 이 브라우저에만 있습니다. 다른 PC 에서 쓰거나 브라우저 기록을 지우기 전에 JSON 백업을 받아 두세요.'))];
  }
  function processCard() {
    var newIn = h('input', { placeholder: '새 공정 이름' });
    function addProc() {
      var n = String(newIn.value || '').trim().replace(/\s+/g, ' ');
      if (!n || n === L.SHIP) { toast('공정 이름을 적어 주세요(「출하」는 자동).', true); return; }
      if (db.processes.indexOf(n) >= 0) { toast('이미 있는 공정입니다.', true); return; }
      db.processes.push(n); save(); render();
    }
    return h('section', { class: 'card' }, h('h2', null, '공정 목록 (이름만)'),
      h('p', { class: 'note' }, '여기 순서는 표의 열 순서일 뿐이고, 실제 흐름은 아이템마다 적는 공정 순서를 따릅니다. 이름을 바꾸면 아이템·기초재고·실적에서도 함께 바뀝니다.'),
      table(['순서', '공정 이름', '쓰는 곳', ''], db.processes.map(function (p, i) {
        var nm = h('input', { value: p, 'aria-label': '공정 이름 ' + p, style: 'max-width:160px' });
        nm.addEventListener('change', function () { if (L.renameProcess(db, p, nm.value)) { save(); render(); toast(p + ' → ' + nm.value.trim()); } else { nm.value = p; toast('바꿀 수 없는 이름입니다(빈 칸·출하·이미 있는 이름).', true); } });
        var u = L.processUse(db, p);
        return [i + 1, nm, '아이템 ' + u.items + ' · 실적 ' + u.results,
          h('span', { class: 'btn-row nowrap' },
            h('button', { type: 'button', class: 'btn btn-sm', disabled: i === 0, 'aria-label': p + ' 위로', onclick: function () { db.processes.splice(i - 1, 0, db.processes.splice(i, 1)[0]); save(); render(); } }, '↑'),
            h('button', { type: 'button', class: 'btn btn-sm', onclick: function () {
              if (u.items || u.results) { toast(p + ' 공정을 쓰는 아이템 ' + u.items + '개 · 실적 ' + u.results + '건이 있어 뺄 수 없습니다. 아이템 공정 순서에서 먼저 빼 주세요.', true); return; }
              db.processes.splice(i, 1); save(); render();
            } }, '빼기'))];
      })),
      h('div', { class: 'btn-row', style: 'margin-top:8px' }, newIn, h('button', { type: 'button', class: 'btn', onclick: addProc }, '공정 추가')));
  }
  function itemsCard() {
    var E = ui.itemErr || {};
    var ins = {}; L.ITEM_FIELDS.forEach(function (f) { ins[f.key] = h('input', { placeholder: f.key === 'customer' ? '예: 고객사A' : f.key === 'part' ? '예: SMP-1001' : '' }); });
    var routeIn = h('input', { placeholder: db.processes.slice(0, 3).join(' > ') });
    function addItem() {
      var input = { route: routeIn.value }; L.ITEM_FIELDS.forEach(function (f) { input[f.key] = ins[f.key].value; });
      var v = L.validateItem(input, db.processes, db.items);
      if (!v.ok) { ui.itemErr = v.errors; render(); toast(Object.keys(v.errors).map(function (k) { return v.errors[k]; }).join(' · '), true); return; }
      ui.itemErr = {}; db.items.push(v.item); save(); render(); toast(v.item.part + ' 추가');
    }
    function importItems(file) {
      readSheets(file).then(function (sheets) {
        var best = null;
        sheets.forEach(function (s) { var o = L.parseItemSheet(s.aoa, db.processes); if (!best || o.rows.length > best.rows.length) best = o; });
        if (!best || !best.rows.length) { toast((best && best.errors[0] && best.errors[0].msg) || '읽을 아이템이 없습니다.', true); return; }
        var msg = [file.name + ' — 아이템 ' + best.rows.length + '개' + (best.format === 'long' ? ' (한 줄에 공정 하나 모양 — 같은 품번의 줄을 위에서 아래 순서로 공정 순서로 읽음)' : '')];
        if (best.newProcesses.length) msg.push('공정 목록에 새로 넣을 공정: ' + best.newProcesses.join(', '));
        best.errors.slice(0, 6).forEach(function (e) { msg.push(e.row + '행: ' + e.msg); });
        msg.push('같은 품번은 내용을 바꾸고(실적·기초재고 유지), 새 품번은 추가합니다.');
        if (!confirm(msg.join('\n') + '\n\n반영할까요?')) return;
        best.newProcesses.forEach(function (p) { if (db.processes.indexOf(p) < 0) db.processes.push(p); });
        var m = L.mergeItems(db.items, best.rows); save(); render(); toast('아이템 ' + m.added + '개 추가 · ' + m.updated + '개 바꿈');
      }, function (e) { toast(e.message, true); });
    }
    var rows = db.items.map(function (it) {
      var cells = L.ITEM_FIELDS.map(function (f) {
        var inp = h('input', { value: it[f.key] || '', 'aria-label': it.part + ' ' + f.label });
        inp.addEventListener('change', function () {
          var input = Object.assign({}, it, { route: it.route }); input[f.key] = inp.value;
          var v = L.validateItem(input, db.processes, db.items, it.id);
          if (!v.ok) { toast(Object.keys(v.errors).map(function (k) { return v.errors[k]; }).join(' · '), true); inp.value = it[f.key]; return; }
          it[f.key] = v.item[f.key]; save(); render();
        });
        return inp;
      });
      var rt = h('input', { value: L.routeText(it), 'aria-label': it.part + ' 공정 순서', style: 'min-width:220px' });
      rt.addEventListener('change', function () {
        var v = L.validateItem(Object.assign({}, it, { route: rt.value }), db.processes, db.items, it.id);
        if (!v.ok) { toast(v.errors.route || '공정 순서를 확인해 주세요.', true); rt.value = L.routeText(it); return; }
        it.route = v.item.route; save(); render(); toast(it.part + ' 공정 순서: ' + L.routeText(it) + ' > 출하');
      });
      return cells.concat([rt, h('button', { type: 'button', class: 'btn btn-sm', onclick: function () {
        var nr = db.results.filter(function (r) { return r.itemId === it.id; }).length;
        if (!confirm(it.part + ' 를 지울까요?' + (nr ? ' 이 품번의 실적 ' + nr + '건·기초재고·진행 중 타이머도 함께 지워집니다.' : ''))) return;
        db.items = db.items.filter(function (x) { return x !== it; });
        db.results = db.results.filter(function (r) { return r.itemId !== it.id; });
        db.opening.rows = db.opening.rows.filter(function (r) { return r.itemId !== it.id; });
        db.timers = db.timers.filter(function (t) { return t.itemId !== it.id; });
        save(); render();
      } }, '지우기')]);
    });
    return h('section', { class: 'card' }, h('h2', null, '관리 아이템 ' + db.items.length + '개'),
      h('p', { class: 'note' }, '공정 순서는 「가접 > 로봇 > 사상」처럼 「>」나 쉼표로 적습니다. 칸을 고치고 다른 곳을 누르면 저장됩니다. 엑셀은 한 줄에 공정 하나(고객사 · 그룹 · 상세그룹 · 품목 · 품번 · 비고 · 공정)로 적고, 같은 품번의 줄을 위에서 아래 순서로 공정 순서로 읽습니다(「출하」 줄은 있어도 되고 없어도 됨). 예전의 「공정 순서」 한 칸 양식도 읽습니다.'),
      table(L.ITEM_FIELDS.map(function (f) { return f.label; }).concat(['공정 순서 (> 출하 자동)', '']), rows, { cls: 'items', empty: '아이템이 없습니다.' }),
      h('h3', { style: 'margin-top:16px' }, '아이템 추가'),
      h('div', { class: 'form-grid' }, L.ITEM_FIELDS.map(function (f) { return fld(f.label, ins[f.key], null, E[f.key]); }),
        h('div', { class: 'span-all' }, fld('공정 순서', routeIn, '공정 목록: ' + db.processes.join(', '), E.route))),
      h('div', { class: 'btn-row', style: 'margin-top:10px' },
        h('button', { type: 'button', class: 'btn btn-primary', onclick: addItem }, '아이템 추가'),
        h('button', { type: 'button', class: 'btn', onclick: function () { download('관리아이템_양식.xlsx', xlsxBlob({ '관리 아이템': L.itemTemplate(db.processes) })); } }, '엑셀 양식 받기'),
        h('button', { type: 'button', class: 'btn', disabled: !db.items.length, onclick: function () { download('관리아이템_' + L.todayIso() + '.xlsx', xlsxBlob({ '관리 아이템': L.itemSheet(db.items) })); } }, '지금 아이템 엑셀로 받기'),
        fileBtn('아이템 엑셀 넣기', '.xlsx,.xls,.xlsm', importItems)));
  }

  /* ── 샘플 ─────────────────────────────────── */
  function loadSample() {
    if ((db.items.length || db.results.length) && !db.sample && !confirm('지금 데이터를 지우고 가상 샘플을 불러올까요? (먼저 설정 화면에서 JSON 백업을 받아 두면 안전합니다)')) return;
    db = window.PPSample.build(Date.now());
    ui.calcItem = ''; ui.proc = ''; ui.from = null; pref.filter = {}; S.saveUi(pref);
    save();
    toast('가상 샘플을 불러왔습니다 — 기초재고 ' + window.PPSample.OPEN + ' 아침 기준. 심어 둔 것: 10-20 로봇 설비 정지(생산 적음), SMP-4001 사상 C/T 가 김, 진행 중 타이머 2개.');
    if (location.hash === '#/dash') render(); else location.hash = '#/dash';
  }

  /* ── 틀 ─────────────────────────────────── */
  var NAV = [
    { hash: '#/dash', no: '0', t: '대시보드', s: '공정별 재고 · C/T', page: pageDash },
    { hash: '#/entry', no: '1', t: '실적 입력', s: '시작 · 멈춤 · 종료', page: pageEntry },
    { hash: '#/import', no: '2', t: '데이터 가져오기', s: '기초재고(1회)', page: pageOpening },
    { hash: '#/process', no: '3', t: '공정 분석', s: '일별 생산수량 ±30%', page: pageProcess },
    { hash: '#/ai', no: '4', t: 'AI 분석', s: '반자동 프롬프트', page: pageAi },
    { hash: '#/settings', no: '', t: '설정', s: '아이템 · 공정 · 백업', page: pageSettings }
  ];
  function current() {
    var hsh = location.hash || '#/dash';
    return NAV.filter(function (n) { return n.hash === hsh; })[0] || NAV[0];
  }
  function render() {
    var cur = current();
    var nav = document.getElementById('nav'); nav.innerHTML = '';
    NAV.forEach(function (n) {
      nav.appendChild(h('a', { href: n.hash, 'aria-current': n === cur ? 'page' : null }, h('span', { class: 'no' }, n.no), h('span', { class: 't' }, n.t), h('span', { class: 's' }, n.s)));
    });
    var chip = document.getElementById('dataChip'); chip.innerHTML = '';
    add(chip, ['기초재고 ' + (db.opening.date ? db.opening.date + ' 기준' : '없음'), h('span', { class: 'dim' }, '아이템 ' + db.items.length + ' · 실적 ' + L.fmtN(db.results.length) + '건 · 작업 중 ' + db.timers.length)]);
    var bn = document.getElementById('sampleBanner');
    bn.hidden = !db.sample;
    bn.textContent = db.sample ? '가상 샘플 데이터를 보고 있습니다(고객사·품번·수량·시간 모두 지어낸 값). 설정에서 「모두 지우기」 후 실제 아이템을 넣어 주세요.' : '';
    main.innerHTML = '';
    try { add(main, cur.page()); }
    catch (e) { console.error(e); main.appendChild(h('div', { class: 'alert warn' }, '화면을 그리지 못했습니다: ' + e.message)); }
    save();
  }
  window.addEventListener('hashchange', function () { render(); main.focus({ preventScroll: true }); window.scrollTo(0, 0); });
  /* 다른 탭에서 타이머를 누르면 이 탭도 따라가게 */
  window.addEventListener('storage', function (e) { if (e.key === 'data10-05.db') { db = S.loadDb(); render(); } });
  render();
})();
