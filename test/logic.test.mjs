// 실행: node test/logic.test.mjs   (의존성 없음 — vendor 의 SheetJS 만 씀)
// 3차(2026-10-28): 기초재고 + 공정 실적 → 공정별 재고 · 타이머 순작업 시간 · C/T · 일별 ±30% · 기준날짜 자르기 · 엑셀 양식 · 옛 자료 옮기기
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import { versions } from '../scripts/bump-version.mjs';
const require = createRequire(import.meta.url);
const L = require('../js/logic.js');
const Sample = require('../js/sample-data.js');
const XLSX = require('../vendor/xlsx.full.min.js');

let passed = 0;
function test(name, fn) {
  try { fn(); passed++; console.log('  ok  ' + name); }
  catch (e) { console.error('  FAIL ' + name + '\n       ' + e.message); process.exitCode = 1; }
}
// 손으로 계산할 수 있는 작은 자료 — 아이템 A: 가공 > 용접 > 도장 > 출하
function small() {
  const db = L.emptyDb();
  db.processes = ['가공', '용접', '도장'];
  db.items = [{ id: 'a', customer: '고객사A', group: 'G1', subgroup: 'S1', product: '프레임', part: 'A-1', note: '', route: ['가공', '용접', '도장'] },
    { id: 'b', customer: '고객사B', group: 'G2', subgroup: 'S2', product: '커버', part: 'B-1', note: '', route: ['가공', '도장'] }];
  db.opening = { date: '2026-10-12', rows: [{ itemId: 'a', process: '가공', qty: 10 }, { itemId: 'a', process: '용접', qty: 5 }, { itemId: 'a', process: '도장', qty: 3 }, { itemId: 'b', process: '도장', qty: 4 }] };
  const R = (itemId, process, qty, date, runSec) => ({ id: L.newId('r'), itemId, process, qty, date, start: null, end: null, runSec: runSec == null ? null : runSec, stops: 0, source: '직접', note: '' });
  db.results = [
    R('a', '가공', 7, '2026-10-12', 700),   // 가공 +7
    R('a', '용접', 6, '2026-10-12', 1200),  // 용접 +6, 가공 −6
    R('a', '도장', 4, '2026-10-13', 400),   // 도장 +4, 용접 −4
    R('a', '출하', 5, '2026-10-13'),        // 출하 누계 +5, 도장 −5
    R('a', '가공', 99, '2026-10-09', 990),  // 기준날짜 전 → 재고에 안 들어감
    R('b', '도장', 2, '2026-10-13'),        // b: 앞 공정 가공 −2 (가공 재고 0 → −2)
    R('b', '출하', 1, '2026-10-14')
  ];
  return db;
}

test('날짜·수량 읽기', () => {
  assert.equal(L.parseDate('2026.10.12'), '2026-10-12');
  assert.equal(L.parseDate(46307), '2026-10-12');
  assert.equal(L.parseDate('2026-02-30'), null);
  assert.equal(L.parseNum('1,200'), 1200);
  assert.equal(L.parseNum(-3), null);
  assert.equal(L.parseQty('-'), 0);
  assert.equal(L.parseQty('-4'), -4);
  assert.equal(L.parseQty('abc'), null);
});

test('재고 흐름 — 실적은 그 공정 +, 바로 앞 공정 −, 첫 공정은 + 만', () => {
  const [a, b] = L.computeStock(small());
  assert.deepEqual(a.stock, { '가공': 10 + 7 - 6, '용접': 5 + 6 - 4, '도장': 3 + 4 - 5 });
  assert.equal(a.wip, 11 + 7 + 2);
  assert.deepEqual(b.stock, { '가공': -2, '도장': 4 + 2 - 1 });
  assert.deepEqual(b.negatives, ['가공']);
});

test('출하 — 마지막 공정(아이템마다 다름) 재고만 빼고 출하 누계에 더함', () => {
  const [a, b] = L.computeStock(small());
  assert.equal(a.shipped, 5); assert.equal(a.used['도장'], 5);
  assert.equal(b.shipped, 1); assert.equal(b.used['도장'], 1);
  assert.equal(a.flow[a.flow.length - 1], '출하');
});

test('기초재고 기준날짜 — 그 날(포함)부터의 실적만 재고에, 기준날짜를 옮기면 바뀜', () => {
  const db = small();
  assert.equal(L.computeStock(db)[0].stock['가공'], 11); // 10-09 의 99 는 빠짐
  db.opening.date = '2026-10-13';
  assert.equal(L.computeStock(db)[0].stock['가공'], 10);  // 10-12 실적(+7 −6)도 빠짐
  db.opening.date = '';
  assert.equal(L.computeStock(db)[0].stock['가공'], 10 + 7 + 99 - 6); // 기준날짜 없음 = 모든 실적
  const r = L.analyze(small());
  assert.ok(r.checks.some(c => /전 실적 1건/.test(c.text)));
});

test('공정 순서에 없는 공정 실적은 재고에 넣지 않고 알림', () => {
  const db = small();
  db.results.push({ id: 'x', itemId: 'b', process: '용접', qty: 3, date: '2026-10-13', runSec: null });
  const r = L.analyze(db);
  assert.equal(r.stock[1].offRoute, 1);
  assert.ok(r.checks.some(c => /공정 순서에 없는 공정 실적 1건/.test(c.text)));
  assert.ok(r.checks.some(c => c.level === '경고' && /B-1/.test(c.text))); // 가공 −2
});

test('타이머 — stop 사이는 빼고, restart 로 이어 재고, end 에 수량 → C/T', () => {
  const T0 = Date.UTC(2026, 9, 13, 0, 0, 0), m = 60000;
  let t = L.timerStart('a', '용접', T0);
  assert.ok(L.timerRunning(t));
  L.timerStop(t, T0 + 20 * m);            // 20분 작업
  assert.equal(L.timerNetMs(t, T0 + 50 * m), 20 * m); // 멈춘 동안은 늘지 않음
  L.timerRestart(t, T0 + 30 * m);         // 10분 쉼
  L.timerRestart(t, T0 + 31 * m);         // 작업 중에 또 누름 → 무시
  assert.equal(t.segs.length, 2);
  assert.equal(L.timerNetMs(t, T0 + 45 * m), 35 * m);
  assert.equal(L.timerPausedMs(t, T0 + 45 * m), 10 * m);
  const bad = L.timerFinish(t, '0', T0 + 60 * m);
  assert.equal(bad.ok, false);
  const f = L.timerFinish(t, '10', T0 + 60 * m); // 20 + 30 = 50분, 10개
  assert.ok(f.ok);
  assert.equal(f.result.runSec, 50 * 60);
  assert.equal(f.result.stops, 1);
  assert.equal(f.result.runSec / f.result.qty, 300); // C/T 5분
  assert.equal(t.segs[1].to, null); // 원래 타이머는 그대로(화면에서 지움)
});

test('타이머 — 멈춘 채로 end 하면 멈춘 시각까지만 셈, 저장·복원(새로고침) 뒤에도 같음', () => {
  const T0 = 1_000_000, m = 60000;
  const t = L.timerStart('a', '가공', T0); L.timerStop(t, T0 + 12 * m);
  assert.equal(L.timerPausedMs(t, T0 + 20 * m), 8 * m); // 지금도 멈춰 있는 시간까지
  const f = L.timerFinish(t, 3, T0 + 90 * m);
  assert.equal(f.result.runSec, 12 * 60);
  const db = small(); db.timers = [L.timerStart('a', '가공', T0)];
  const back = L.restoreDb(JSON.parse(JSON.stringify(db)));
  assert.equal(back.timers.length, 1);
  assert.equal(L.timerNetMs(back.timers[0], T0 + 7 * m), 7 * m);
});

test('C/T — 시간 합 ÷ 수량 합(수량 가중), 시간 없는 실적은 뺌', () => {
  const db = small();
  db.results.push({ id: 'y', itemId: 'a', process: '용접', qty: 4, date: '2026-10-14', runSec: 1600 });
  const ct = L.ctTable(db.results);
  const w = ct.find(c => c.itemId === 'a' && c.process === '용접');
  assert.equal(w.ct, (1200 + 1600) / (6 + 4)); // 280초
  assert.equal(w.min, 200); assert.equal(w.max, 400);
  assert.ok(!ct.some(c => c.process === '출하')); // 출하는 시간 없음
  const p = L.ctByProcess(ct, ['가공', '용접']);
  assert.equal(p[0].ct, (700 + 990) / (7 + 99));
});

test('일별 생산수량 — 실적 있는 날 평균, ±30% 경계 포함', () => {
  const R = (date, qty) => ({ itemId: 'a', process: '용접', qty, date });
  const d = L.dailyProduction([R('2026-10-12', 10), R('2026-10-13', 7), R('2026-10-14', 13), R('2026-10-15', 6.9), R('2026-10-15', 3.1), R('2026-10-16', 10)], '용접');
  assert.equal(d.avg, 10);
  assert.deepEqual(d.days.map(x => x.flag), ['', '적음', '많음', '', '']); // 7 = −30%(포함), 13 = +30%(포함)
  assert.equal(d.days[3].n, 2);
  const d2 = L.dailyProduction([R('2026-10-12', 10), R('2026-10-13', 7.1), R('2026-10-14', 12.9)], '용접');
  assert.deepEqual(d2.flagged, []);
});

test('거르기 — 고객사·그룹 등으로 아이템을 고르면 재고·C/T 도 그 아이템만', () => {
  const db = small();
  const r = L.analyze(db, { customer: '고객사B' });
  assert.deepEqual(r.items.map(i => i.part), ['B-1']);
  assert.deepEqual(r.processes, ['가공', '도장']);
  assert.equal(r.kpi.shipped, 1);
  assert.deepEqual(L.filterOptions(db.items).customer, ['고객사A', '고객사B']);
});

test('아이템 검사 — 품번 필수·중복, 공정 목록에 없는 공정·두 번 나온 공정 거부, 출하는 자동', () => {
  const procs = ['가공', '용접', '도장'];
  assert.equal(L.validateItem({ part: '', route: '가공' }, procs, []).ok, false);
  assert.match(L.validateItem({ part: 'a-1', route: '가공' }, procs, [{ id: 'z', part: 'A-1' }]).errors.part, /같은 품번/);
  assert.match(L.validateItem({ part: 'X', route: '가공 > 연마' }, procs, []).errors.route, /연마/);
  assert.match(L.validateItem({ part: 'X', route: '가공, 가공' }, procs, []).errors.route, /두 번/);
  const ok = L.validateItem({ part: ' x-9 ', route: '가공 → 용접 > 출하' }, procs, []);
  assert.ok(ok.ok); assert.equal(ok.item.part, 'X-9'); assert.deepEqual(ok.item.route, ['가공', '용접']);
  assert.deepEqual(L.flowOf(ok.item), ['가공', '용접', '출하']);
});

// 엑셀로 쓰고 다시 읽기(앱과 같은 방식)
function roundTrip(aoa) {
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), 'x');
  return XLSX.utils.sheet_to_json(XLSX.read(XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' })).Sheets.x, { header: 1, raw: true, defval: null });
}
// 10-30 받은 실제 목록과 같은 모양(값은 지어낸 것) — 한 줄에 공정 하나, 같은 품번을 공정 수만큼, 마지막 줄 「출하」
function longSheet() {
  const H = ['고객사', '그룹', '상세그룹', '품목', '품번', '비고', '공정'];
  const rows = [];
  [['X-100', '10ton'], ['X-200', '12ton']].forEach(([part, note]) => ['가공', '용접', 'MT', '검사', '출하'].forEach(pr => rows.push(['고객사C', '중형', '하부', '프레임', part, note, pr])));
  return [H].concat(rows);
}

test('관리 아이템 엑셀(한 줄에 공정 하나 · 실제 목록 모양) — 품번별로 모아 위→아래가 공정 순서, 출하 줄은 자동, 새 공정 알림', () => {
  const p = L.parseItemSheet(roundTrip(longSheet()), ['가공', '용접', '도장']);
  assert.equal(p.format, 'long');
  assert.deepEqual(p.errors, []);
  assert.equal(p.rows.length, 2);
  assert.deepEqual(p.rows.map(r => r.part), ['X-100', 'X-200']);
  assert.deepEqual(p.rows[0].route, ['가공', '용접', 'MT', '검사']);
  assert.deepEqual(L.flowOf(p.rows[1]), ['가공', '용접', 'MT', '검사', '출하']);
  assert.equal(p.rows[1].note, '12ton'); assert.equal(p.rows[0].customer, '고객사C');
  assert.deepEqual(p.newProcesses, ['MT', '검사']);
});

test('관리 아이템 엑셀(한 줄에 공정 하나) — 품번 빈 칸은 위 품번 이어 씀(셀 병합), 줄마다 다른 값 알림, 출하만 있는 품번·두 번 나온 공정 거부', () => {
  const aoa = [['관리 아이템'], ['고객사', '그룹', '상세그룹', '품목', '품번', '비고', '공정'],
    ['고객사C', '소형', '상부', '커버', 'Y-1', '', '가공'], [null, null, null, null, null, null, '도장'], ['고객사C', '소형', '상부', '커버', 'Y-1', '', '출하'],
    ['고객사C', '소형', '상부', '커버', 'Y-2', 'A', '가공'], ['고객사C', '대형', '상부', '커버', 'Y-2', 'A', '도장'],
    ['고객사C', '소형', '상부', '커버', 'Y-3', '', '출하'],
    ['고객사C', '소형', '상부', '커버', 'Y-4', '', '가공'], ['고객사C', '소형', '상부', '커버', 'Y-4', '', '가공']];
  const p = L.parseItemSheet(roundTrip(aoa), ['가공', '도장']);
  assert.equal(p.format, 'long');
  assert.deepEqual(p.rows.map(r => r.part + ':' + r.route.join('>')), ['Y-1:가공>도장', 'Y-2:가공>도장']);
  assert.equal(p.rows[1].group, '소형'); // 첫 줄 값
  assert.ok(p.errors.some(e => /Y-2/.test(e.msg) && /그룹/.test(e.msg)));
  assert.ok(p.errors.some(e => /Y-3/.test(e.msg) && /공정을 하나 이상/.test(e.msg)));
  assert.ok(p.errors.some(e => /Y-4/.test(e.msg) && /두 번/.test(e.msg)));
});

test('관리 아이템 엑셀 — 받은 양식·지금 아이템 받기는 실제 목록 모양(출하 줄 포함), 다시 넣으면 그대로, 같은 품번은 바꿈(id 유지)', () => {
  const tpl = L.itemTemplate(['가공', '용접']);
  assert.deepEqual(tpl[0], ['고객사', '그룹', '상세그룹', '품목', '품번', '비고', '공정']);
  assert.equal(tpl[1][0], '고객사A'); // 양식 예시는 가짜 고객사
  assert.deepEqual(tpl.slice(1).map(r => r[6]), ['가공', '용접', '출하']);
  const db = small();
  const sheet = L.itemSheet(db.items);
  assert.equal(sheet.length, 1 + 4 + 3); // A-1 가공·용접·도장·출하, B-1 가공·도장·출하
  const p = L.parseItemSheet(roundTrip(sheet), db.processes);
  assert.equal(p.format, 'long'); assert.deepEqual(p.errors, []); assert.deepEqual(p.newProcesses, []);
  assert.deepEqual(p.rows.map(r => r.part + ':' + r.route.join('>')), db.items.map(r => r.part + ':' + r.route.join('>')));
  p.rows[0].note = '바뀜';
  const m = L.mergeItems(db.items, p.rows);
  assert.deepEqual(m, { added: 0, updated: 2 });
  assert.equal(db.items[0].id, 'a'); assert.equal(db.items[0].note, '바뀜');
});

test('관리 아이템 엑셀(한 줄에 아이템 하나 · 3차 첫 판) — 「공정 순서」 칸도 그대로 읽음', () => {
  const db = small();
  const aoa = [['관리 아이템 목록'], L.ITEM_HEAD_WIDE, ['고객사A', 'G1', 'S1', '프레임', 'A-1', '', '가공 > 용접 > 도장'],
    ['고객사C', 'G3', '', '브래킷', 'C-1', '', '가공 > 연마'], ['', '', '', '', 'A-1', '', '가공'], ['', '', '', '', '', '', '']];
  const p = L.parseItemSheet(roundTrip(aoa), db.processes);
  assert.equal(p.format, 'wide');
  assert.equal(p.rows.length, 2);
  assert.deepEqual(p.newProcesses, ['연마']);
  assert.equal(p.errors.length, 1); assert.match(p.errors[0].msg, /같은 품번/);
  const m = L.mergeItems(db.items, p.rows);
  assert.deepEqual(m, { added: 1, updated: 1 });
});

test('10-30 답 — 출하도 공정 실적으로(직접 적기·타이머), C/T 는 인원과 무관하게 순작업 시간 ÷ 수량, 첫 공정은 자재 차감 없음', () => {
  const db = small();
  const v = L.validateManual({ date: '2026-10-14', itemId: 'a', process: '출하', qty: '2', minutes: '10' }, db);
  assert.ok(v.ok); db.results.push(v.result);
  const t0 = L.localMs('2026-10-14', '10:00'); const t = L.timerStart('a', '출하', t0); const f = L.timerFinish(t, 1, t0 + 60000);
  assert.ok(f.ok); db.results.push(f.result);
  const a = L.computeStock(db).find(o => o.item.id === 'a');
  assert.equal(a.shipped, 5 + 2 + 1);
  assert.equal(a.stock['도장'], 3 + 4 - 5 - 2 - 1);
  assert.equal(a.stock['가공'], 10 + 7 - 6); // 첫 공정: 자재를 빼지 않음
  const ct = L.ctTable(db.results).find(c => c.itemId === 'a' && c.process === '출하');
  assert.equal(ct.ct, (600 + 60) / 3);
});

test('기초재고 엑셀 — 양식 → 다시 읽기(기준날짜·「-」·없는 품번·공정 순서 밖 칸)', () => {
  const db = small();
  const aoa = L.openingSheet(db);
  assert.deepEqual(aoa[0], ['기준날짜', '2026-10-12']);
  assert.deepEqual(aoa[2], ['품번', '품목', '가공', '용접', '도장']);
  assert.equal(aoa[4][3], '-'); // B-1 은 용접을 거치지 않음
  const edit = aoa.map(r => r.slice());
  edit[0][1] = '2026.10.20'; edit[3][2] = '1,200'; edit[4][3] = 9; edit.push(['Z-9', '', 1, 1, 1]);
  const o = L.parseOpeningSheet(edit, db);
  assert.equal(o.date, '2026-10-20');
  assert.ok(o.rows.some(r => r.itemId === 'a' && r.process === '가공' && r.qty === 1200));
  assert.deepEqual(o.unknownParts, ['Z-9']);
  assert.deepEqual(o.offRoute, ['B-1 · 용접 9']);
  assert.equal(L.parseOpeningSheet([['아무것']], db).errors.length, 1);
});

test('공정 이름 바꾸기 — 아이템·기초재고·실적·타이머 함께', () => {
  const db = small(); db.timers = [L.timerStart('a', '용접', 0)];
  assert.ok(L.renameProcess(db, '용접', '로봇'));
  assert.deepEqual(db.items[0].route, ['가공', '로봇', '도장']);
  assert.ok(db.results.some(r => r.process === '로봇')); assert.equal(db.timers[0].process, '로봇');
  assert.ok(db.opening.rows.some(r => r.process === '로봇'));
  assert.equal(L.renameProcess(db, '가공', '도장'), false);
  assert.equal(L.renameProcess(db, '가공', '출하'), false);
});

test('옛 판(1·2차) 저장 자료 — 깨지지 않고 공정·모델을 옮기고 안내(notice)', () => {
  const old = { settings: { processes: [{ name: '로봇', workers: 2 }, { name: '사상' }] }, orders: [{}, {}], results: [{ item: 'X' }], deliveries: [],
    checks: [{ date: '2026-10-20', rows: [{ item: 'smp-m30', vals: {} }, { item: 'SMP-M20' }] }], stock: { rows: [{ item: 'SMP-M30' }] } };
  const db = L.restoreDb(old);
  assert.equal(db.schemaVersion, L.SCHEMA_VERSION);
  assert.deepEqual(db.processes, ['로봇', '사상']);
  assert.deepEqual(db.items.map(i => i.part), ['SMP-M30', 'SMP-M20']);
  assert.deepEqual(db.items[0].route, ['로봇', '사상']);
  assert.equal(db.notice.had, 4);
  L.analyze(db); // 계산이 깨지지 않음
  const s = L.restoreDb(Object.assign({ sample: true }, old));
  assert.equal(s.items.length, 0); assert.ok(s.notice.sample);
  for (const junk of [null, 'x', 42, [], { schemaVersion: 2, items: 'x', results: [null, 1] }]) L.analyze(L.restoreDb(junk));
});

test('저장 형식 복원 — 지운 아이템의 실적·타이머·이상한 값은 버림', () => {
  const db = small();
  db.results.push({ id: 'q', itemId: 'gone', process: '가공', qty: 1, date: '2026-10-13' }, { id: 'w', itemId: 'a', process: '가공', qty: 0, date: '2026-10-13' });
  const back = L.restoreDb(JSON.parse(JSON.stringify(db)));
  assert.equal(back.results.length, small().results.length);
  assert.deepEqual(L.computeStock(back)[0].stock, L.computeStock(small())[0].stock);
});

test('AI 프롬프트 — 품번·고객사 별칭, 실제 이름 없음, 되돌리기', () => {
  const db = small(); const r = L.analyze(db);
  const al = { items: L.assignAliases({}, r.items.map(i => i.part), L.ITEM_ALIAS), customers: L.assignAliases({}, r.items.map(i => i.customer), L.CUST_ALIAS) };
  const p = L.buildAiPrompt(r, db, { aliases: al });
  assert.ok(!/A-1|B-1|고객사A|고객사B/.test(p.prompt));
  assert.deepEqual(p.leaks, []);
  assert.equal(L.unmaskText('품목1 은 거래처2 에 감, 품목10개', al), 'A-1 은 고객사B 에 감, 품목10개');
});

test('가상 샘플 — 재고 음수 없음, 10-20 로봇 적음 · 10-21 많음, SMP-4001 사상 C/T 가 김, 진행 중 타이머 2개', () => {
  const now = Date.now();
  const db = Sample.build(now);
  const r = L.analyze(db);
  assert.equal(r.kpi.negatives, 0);
  const robot = r.daily.find(d => d.process === '로봇');
  assert.deepEqual(robot.flagged.map(x => x.date + x.flag), ['2026-10-20적음', '2026-10-21많음']);
  const sa = r.ct.filter(c => c.process === '사상');
  const big = sa.find(c => c.itemId === 'smp4001'), others = sa.filter(c => c.itemId !== 'smp4001');
  assert.ok(others.every(c => big.ct > c.ct * 1.3));
  assert.equal(db.timers.length, 2);
  assert.equal(db.timers.filter(L.timerRunning).length, 1);
  assert.ok(r.checks.some(c => /전 실적 2건/.test(c.text)));
  assert.ok(db.items.every(i => /^고객사[A-Z]$/.test(i.customer) && /^SMP-\d{4}$/.test(i.part))); // 샘플은 가짜 값만
  // 같은 씨앗 → 같은 값
  assert.equal(JSON.stringify(Sample.build(now).results), JSON.stringify(db.results));
});

test('내보내기 시트 — 공정별 재고현황·C/T·일별·실적', () => {
  const db = small(), sh = L.exportSheets(L.analyze(db), db);
  assert.deepEqual(Object.keys(sh), ['공정별 재고현황', 'C/T', '일별 생산수량', '공정 실적']);
  assert.deepEqual(sh['공정별 재고현황'][1].slice(4), ['A-1', 11, 7, 2, 5, 20]);
  assert.equal(sh['공정 실적'].length, db.results.length + 1);
});

test('판(版) — index.html ?v= · logic.js · app.js 가 같음', () => {
  const v = versions(), all = [v.logic, v.app, ...v.html];
  assert.ok(all.every(x => x && x === all[0]), JSON.stringify(v));
});

console.log('\n' + passed + '개 통과' + (process.exitCode ? ' — 실패 있음' : ''));
