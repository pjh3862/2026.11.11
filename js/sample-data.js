/* 가상 샘플 — 3차(2026-10-27) 범위: 관리 아이템 · 기초재고 · 공정 실적(시작~종료 시간) · 진행 중 타이머
   · 고객사·품번·품목·수량·시간은 모두 지어낸 값입니다(고객사A·B, 품번 SMP-…).
   · 기초재고 기준날짜 2026-10-12(월) 아침, 실적 10-12 ~ 10-26 근무일(월~금).
   · 아이템마다 공정 순서가 다릅니다(SMP-3001 은 가공을 MT 보다 먼저, SMP-2001 은 MT·가공 없음).
   · 일부러 심어 둔 것:
       ① 10-20(화) 로봇 설비 정지 — 그날 로봇 생산수량이 평균보다 크게 적음(±30% 밖), 다음 날 잔업으로 많음
          → 10-21 은 사상이 가져다 쓸 로봇 완료재고가 모자라 사상도 적음(뒤 공정으로 번지는 모습)
       ② SMP-4001 사상 C/T 가 다른 품번보다 김(대형)
       ③ 기준날짜 전(10-09) 실적 2건 — 재고에는 안 들어가고 C/T 에만 쓰임을 보여 줌
   · 「가상 샘플로 시작」을 누른 시각 기준으로 진행 중 타이머 2개(작업 중 1 · 멈춤 1)를 함께 넣습니다.
   브라우저에서는 window.PPSample, node 에서는 require('./sample-data.js'). */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('./logic.js'));
  else root.PPSample = factory(root.PPLogic);
})(typeof self !== 'undefined' ? self : this, function (L) {
  'use strict';
  var OPEN = '2026-10-12', LAST = '2026-10-26';
  var PROCS = ['가접', '로봇', '사상', 'MT', '가공', '도킹'];
  var CT = { '가접': 540, '로봇': 420, '사상': 330, 'MT': 280, '가공': 390, '도킹': 480, '출하': 45 }; // 초/개(기준)
  var ITEMS = [
    { id: 'smp1001', customer: '고객사A', group: '중형', subgroup: '하부', product: '프레임', part: 'SMP-1001', note: '3PCS 세트', route: ['가접', '로봇', '사상', 'MT', '도킹'], d: 6, f: 1 },
    { id: 'smp1002', customer: '고객사A', group: '중형', subgroup: '하부', product: '프레임', part: 'SMP-1002', note: '', route: ['가접', '로봇', '사상', 'MT', '가공', '도킹'], d: 4, f: 1.1 },
    { id: 'smp2001', customer: '고객사A', group: '중형', subgroup: '상부', product: '커버', part: 'SMP-2001', note: '', route: ['가접', '로봇', '사상', '도킹'], d: 5, f: 0.8 },
    { id: 'smp3001', customer: '고객사A', group: '소형', subgroup: '상부', product: '브래킷', part: 'SMP-3001', note: '가공 먼저', route: ['가접', '사상', '가공', 'MT'], d: 8, f: 0.6 },
    { id: 'smp4001', customer: '고객사B', group: '대형', subgroup: '하부', product: '프레임', part: 'SMP-4001', note: '2PCS 세트', route: ['가접', '로봇', '사상', 'MT', '가공', '도킹'], d: 3, f: 1.5 },
    { id: 'smp4002', customer: '고객사B', group: '대형', subgroup: '상부', product: '암', part: 'SMP-4002', note: '', route: ['가접', '로봇', '가공', '도킹'], d: 3, f: 1.2 }
  ];
  function rng(seed) { return function () { seed |= 0; seed = seed + 0x6D2B79F5 | 0; var t = Math.imul(seed ^ seed >>> 15, 1 | seed); t = t + Math.imul(t ^ t >>> 7, 61 | t) ^ t; return ((t ^ t >>> 14) >>> 0) / 4294967296; }; }
  function workdays() { var out = [], d = OPEN; while (d <= LAST) { var w = L.weekday(d); if (w >= 1 && w <= 5) out.push(d); d = L.addDays(d, 1); } return out; }

  function build(now) {
    now = now == null ? Date.now() : now;
    var R = rng(20261027);
    var db = L.emptyDb();
    db.processes = PROCS.slice();
    db.items = ITEMS.map(function (x) { return { id: x.id, customer: x.customer, group: x.group, subgroup: x.subgroup, product: x.product, part: x.part, note: x.note, route: x.route.slice() }; });
    db.opening = { date: OPEN, rows: [] };
    var stock = {};
    ITEMS.forEach(function (x) {
      stock[x.id] = {};
      x.route.forEach(function (p, i) { var q = Math.round(x.d * (0.8 + R() * 0.6)) + (i === x.route.length - 1 ? x.d : 0); stock[x.id][p] = q; db.opening.rows.push({ itemId: x.id, process: p, qty: q }); });
    });
    var clock = {}; // 공정별 그날 다음 시작 시각(분)
    function add(x, p, date, q) {
      if (!(q > 0)) return;
      var start0 = clock[p] == null ? 8 * 60 + Math.floor(R() * 20) : clock[p];
      var ct = CT[p] * (p === '출하' ? 1 : x.f) * (0.85 + R() * 0.3);
      if (x.id === 'smp4001' && p === '사상') ct *= 1.6; // ② 대형 사상이 오래 걸림
      var run = Math.round(ct * q), stops = R() < 0.35 ? 1 : 0, pause = stops ? 600 + Math.floor(R() * 600) : 0;
      var start = L.localMs(date, Math.floor(start0 / 60) + ':' + (start0 % 60));
      var end = start + (run + pause) * 1000;
      clock[p] = start0 + Math.ceil((run + pause) / 60) + 5;
      db.results.push({ id: 'r' + db.results.length, itemId: x.id, process: p, qty: q, date: date, start: start, end: end, runSec: run, stops: stops, source: p === '출하' ? '직접' : '타이머', note: '' });
    }
    // ③ 기준날짜 전 실적(재고에는 안 들어감)
    add(ITEMS[0], '로봇', '2026-10-09', 5); add(ITEMS[0], '사상', '2026-10-09', 5);
    workdays().forEach(function (date) {
      clock = {};
      ITEMS.forEach(function (x) {
        var flow = x.route.concat(['출하']);
        // 뒤 공정부터 — 그날 아침 앞 공정 완료재고 안에서만 가져다 씀(재고가 0 아래로 가지 않게)
        for (var i = flow.length - 1; i >= 0; i--) {
          var p = flow[i], want = x.d * (0.8 + R() * 0.4);
          if (p === '로봇' && date === '2026-10-20') want *= 0.35; // ① 설비 정지
          if (p === '로봇' && date === '2026-10-21') want *= 1.6;  //    잔업으로 만회
          var q = Math.round(want);
          if (i > 0) q = Math.min(q, stock[x.id][flow[i - 1]]);
          if (q <= 0) continue;
          add(x, p, date, q);
          if (p !== '출하') stock[x.id][p] += q;
          if (i > 0) stock[x.id][flow[i - 1]] -= q;
        }
      });
    });
    // 진행 중 타이머 — 누른 시각 기준
    var t1 = L.timerStart('smp1001', '로봇', now - 38 * 60000); L.timerStop(t1, now - 20 * 60000); L.timerRestart(t1, now - 12 * 60000);
    var t2 = L.timerStart('smp4001', '사상', now - 25 * 60000); L.timerStop(t2, now - 6 * 60000);
    db.timers = [t1, t2];
    db.sample = true;
    return db;
  }
  /* samples/ 폴더에 쓰는 엑셀 2개 — 관리 아이템 · 기초재고 (가져오기 양식과 같은 모양) */
  function files() {
    var db = build(0);
    return [
      { fileName: '관리아이템_샘플.xlsx', sheetName: '관리 아이템', aoa: L.itemSheet(db.items) },
      { fileName: '기초재고_샘플_' + OPEN + '.xlsx', sheetName: '기초재고', aoa: L.openingSheet(db) }
    ];
  }
  return { build: build, files: files, OPEN: OPEN, LAST: LAST };
});
