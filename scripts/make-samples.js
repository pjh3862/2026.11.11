// 가상 샘플 파일 만들기: node scripts/make-samples.js
// js/sample-data.js 의 자료로 가져오기 양식과 같은 모양의 .xlsx 2개를 씁니다(고객사·품번·수량은 전부 지어낸 값).
//  · 관리아이템_샘플.xlsx — 설정 > 「아이템 엑셀 넣기」 양식
//  · 기초재고_샘플_<기준날짜>.xlsx — 데이터 가져오기 > 「기초재고 엑셀 넣기」 양식
// 방금 쓴 파일을 앱과 같은 방식으로 다시 읽어 확인합니다.
const fs = require('fs');
const path = require('path');
const XLSX = require('../vendor/xlsx.full.min.js');
const L = require('../js/logic.js');
const Sample = require('../js/sample-data.js');

const out = path.join(__dirname, '..', 'samples');
fs.mkdirSync(out, { recursive: true });
for (const f of fs.readdirSync(out)) if (/\.xlsx$/.test(f)) fs.unlinkSync(path.join(out, f));

const db = Sample.build(0);
for (const f of Sample.files()) {
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(f.aoa);
  ws['!cols'] = (f.aoa[f.aoa.length - 1] || []).map(() => ({ wch: 14 }));
  XLSX.utils.book_append_sheet(wb, ws, f.sheetName);
  const file = path.join(out, f.fileName);
  fs.writeFileSync(file, XLSX.write(wb, { bookType: 'xlsx', type: 'buffer', compression: true }));
  const back = XLSX.read(fs.readFileSync(file), { type: 'buffer', cellDates: true });
  const aoa = XLSX.utils.sheet_to_json(back.Sheets[back.SheetNames[0]], { header: 1, raw: true, defval: null });
  if (/아이템/.test(f.fileName)) {
    const p = L.parseItemSheet(aoa, db.processes);
    if (p.rows.length !== db.items.length || p.errors.length) throw new Error(f.fileName + ' 다시 읽기: ' + p.rows.length + '개, 오류 ' + JSON.stringify(p.errors));
    console.log(f.fileName.padEnd(30), '아이템', p.rows.length + '개');
  } else {
    const o = L.parseOpeningSheet(aoa, db);
    if (o.date !== Sample.OPEN || o.errors.length || o.rows.filter(r => r.qty).length !== db.opening.rows.length) throw new Error(f.fileName + ' 다시 읽기 실패: ' + JSON.stringify(o.errors));
    console.log(f.fileName.padEnd(30), '기초재고', o.rows.length + '칸', '기준날짜', o.date);
  }
}
