// 헤드리스 브라우저로 주요 흐름 점검 (선택 — playwright-core 가 있는 PC 에서만)
//   PLAYWRIGHT_CORE=/경로/node_modules/playwright-core node scripts/check-browser.mjs
// file:// 로 index.html 을 열어
//   ① 옛 판(1·2차) 저장 자료가 있는 브라우저 → 깨지지 않고 안내가 뜸
//   ② 가상 샘플 → 6개 화면 × 1280·390px 가로 넘침 · 화면 오류
//   ③ 타이머 start → stop → restart → (새로고침) → end 수량 → 실적 1건 · 공정 재고가 바뀜
//   ④ 샘플 엑셀 2개(아이템·기초재고) 넣기
// 순서로 누르고 JS 오류를 잽니다. BREAK=1 이면 일부러 넓은 상자를 넣어 넘침 검사가 잡는지 봅니다.
import { createRequire } from 'node:module';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
const require = createRequire(import.meta.url);
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const { chromium } = require(process.env.PLAYWRIGHT_CORE || 'playwright-core');
const url = 'file://' + path.join(root, 'index.html');

const errors = [], problems = [], requests = [];
const browser = await chromium.launch(process.env.CHROME ? { executablePath: process.env.CHROME } : {}); // CHROME=브라우저 실행 파일(판이 맞는 브라우저가 없을 때)
const ctx = await browser.newContext({ acceptDownloads: true, viewport: { width: 1280, height: 900 } });
const page = await ctx.newPage();
page.on('pageerror', e => errors.push('pageerror: ' + e.message));
page.on('console', m => { if (m.type() === 'error') errors.push('console: ' + m.text()); });
page.on('request', r => { if (!r.url().startsWith('file:') && !r.url().startsWith('blob:') && !r.url().startsWith('data:')) requests.push(r.url()); });
page.on('dialog', d => d.accept());
function ok(cond, msg) { if (!cond) { problems.push(msg); console.log('  FAIL ' + msg); } else console.log('  ok  ' + msg); }
async function go(hash) { await page.evaluate(h => { location.hash = h; }, hash); await page.waitForTimeout(200); }
async function text() { return page.textContent('main'); }
async function overflow() {
  return page.evaluate(() => {
    const W = document.documentElement.clientWidth, bad = [];
    document.querySelectorAll('body *').forEach(el => {
      const r = el.getBoundingClientRect(); if (!r.width || !r.height) return;
      let p = el.parentElement, inScroll = false;
      while (p) { const o = getComputedStyle(p).overflowX; if (o === 'auto' || o === 'scroll') { inScroll = true; break; } p = p.parentElement; }
      if (!inScroll && r.right > W + 1) bad.push(el.tagName + '.' + el.className);
    });
    return bad.slice(0, 3);
  });
}
async function stockOf(part, proc) {
  return page.evaluate(([part, proc]) => {
    const t = [...document.querySelectorAll('table.list')].find(x => x.querySelector('th') && x.querySelector('th').textContent === '고객사');
    const heads = [...t.querySelectorAll('thead th')].map(x => x.textContent);
    const row = [...t.querySelectorAll('tbody tr')].find(r => r.children[1] && r.children[1].textContent === part);
    return Number(row.children[heads.indexOf(proc)].textContent.replace(/,/g, ''));
  }, [part, proc]);
}

// ① 옛 판 자료
await page.goto(url);
await page.evaluate(() => localStorage.setItem('data10-05.db', JSON.stringify({ settings: { processes: [{ name: '로봇' }, { name: '사상' }] }, orders: [{}], results: [{}], checks: [{ rows: [{ item: 'OLD-M1' }] }] })));
await page.reload(); await page.waitForTimeout(300);
let t = await text();
ok(/저장 형식이 바뀌었습니다/.test(t) && !/화면을 그리지 못했습니다/.test(t), '옛 판 자료: 깨지지 않고 안내');
ok(await page.evaluate(() => !!localStorage.getItem('data10-05.db.v1')), '옛 판 원래 글을 따로 보관');
await go('#/settings');
ok(await page.locator('main input[value="OLD-M1"]').count() === 1, '옛 모델이 관리 아이템으로 옮겨짐');
await go('#/dash'); await page.evaluate(() => localStorage.clear()); await page.reload(); await page.waitForTimeout(300);

// ② 샘플
if (process.env.BREAK) await page.addStyleTag({ content: '.page-head { min-width: 900px; }' });
await page.locator('button', { hasText: '샘플로 시작' }).first().click();
await page.waitForTimeout(500);
ok(await page.locator('#sampleBanner').isVisible(), '샘플 띠 보임');
t = await text();
ok(t.includes('공정별 · 모델별 완료재고') && t.includes('실제 걸린 시간 (C/T)') && t.includes('계산 보기'), '대시보드: 재고 표·C/T·계산 보기');
ok(await page.locator('main svg.chart').count() >= 3, '대시보드: 그래프 3개');
await page.locator('.filters select').first().selectOption('고객사B'); await page.waitForTimeout(200);
ok((await text()).includes('SMP-4001') && !(await text()).includes('SMP-1001'), '거르기: 고객사B 만');
await page.locator('.filters select').first().selectOption(''); await page.waitForTimeout(200);
for (const id of ['dash', 'entry', 'import', 'process', 'ai', 'settings']) {
  for (const w of [1280, 390]) {
    await page.setViewportSize({ width: w, height: 900 });
    await go('#/' + id);
    const bad = await overflow();
    ok(!bad.length && !/화면을 그리지 못했습니다/.test(await text()), id + ' ' + w + 'px 가로 넘침·화면 오류 없음' + (bad.length ? ' — ' + bad.join(', ') : ''));
  }
}
await go('#/process');
t = await text();
ok(/로봇 ⚑2/.test(t) && /10\/20 적음, 10\/21 많음/.test(t), '공정 분석: 로봇 10-20 적음 · 10-21 많음');
for (const old of ['재고 과부족', '생산진도율', '재고 체크', '발주서 가져']) ok(!(await page.textContent('#nav')).includes(old), '메뉴에 「' + old + '」 없음');

// ③ 타이머 (휴대폰 폭)
await page.setViewportSize({ width: 390, height: 900 });
await go('#/dash');
const before = await stockOf('SMP-2001', '사상'), beforeRobot = await stockOf('SMP-2001', '로봇');
await go('#/entry');
await page.locator('main select').first().selectOption({ label: 'SMP-2001 · 커버 (고객사A)' }); await page.waitForTimeout(200);
await page.locator('label.pill', { hasText: '3. 사상' }).click();
await page.locator('button', { hasText: '시작 (start)' }).click(); await page.waitForTimeout(300);
ok(await page.locator('.timer-card').count() === 3, '타이머 시작 → 진행 중 3개');
await page.waitForTimeout(1300);
const card = () => page.locator('.timer-card').last();
await card().locator('button', { hasText: '멈춤' }).click(); await page.waitForTimeout(200);
const c1 = await card().locator('.timer-clock').textContent();
await page.waitForTimeout(1300);
ok(await card().locator('.timer-clock').textContent() === c1, '멈춤 동안 시계가 서 있음 (' + c1 + ')');
await page.reload(); await page.waitForTimeout(400);
ok(await page.locator('.timer-card').count() === 3, '새로고침 뒤에도 타이머 그대로');
await card().locator('button', { hasText: '다시 시작' }).click(); await page.waitForTimeout(1200);
await card().locator('input[aria-label="완료 수량"]').fill('2');
await card().locator('button', { hasText: '종료' }).click(); await page.waitForTimeout(300);
ok(await page.locator('.timer-card').count() === 2, '종료 → 타이머 사라짐');
ok((await page.textContent('#toast')).includes('실적 저장'), '종료 → 실적 저장 알림');
ok((await overflow()).length === 0, '실적 입력 390px 가로 넘침 없음(타이머 있음)');
await go('#/dash');
ok(await stockOf('SMP-2001', '사상') === before + 2 && await stockOf('SMP-2001', '로봇') === beforeRobot - 2, '사상 +2 · 앞 공정(로봇) −2');
await page.setViewportSize({ width: 1280, height: 900 });

// ④ 샘플 엑셀
await go('#/settings');
await page.setInputFiles('main input[type=file][accept=".xlsx,.xls,.xlsm"]', path.join(root, 'samples', '관리아이템_샘플.xlsx')); await page.waitForTimeout(800);
ok((await page.textContent('#toast')).includes('0개 추가 · 6개 바꿈'), '아이템 엑셀 넣기 — 같은 품번 6개 바꿈');
await go('#/import');
await page.setInputFiles('main input[type=file]', path.join(root, 'samples', '기초재고_샘플_2026-10-12.xlsx')); await page.waitForTimeout(800);
ok((await page.textContent('#toast')).includes('기초재고를 가져왔습니다'), '기초재고 엑셀 넣기');
ok(!(await page.locator('#sampleBanner').isVisible()), '기초재고를 넣으면 샘플 띠 사라짐');
ok(!requests.length, '외부 주소 요청 0건' + (requests.length ? ' — ' + requests.slice(0, 3).join(', ') : ''));

await browser.close();
if (errors.length) console.log('JS 오류:\n' + errors.join('\n'));
console.log(errors.length || problems.length ? '\n문제 ' + (errors.length + problems.length) + '건' : '\n모두 통과');
process.exitCode = errors.length || problems.length ? 1 : 0;
