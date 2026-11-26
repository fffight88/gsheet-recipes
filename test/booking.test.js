// Logic test for booking-auto-confirm.gs with mocked Google services (not a live Apps Script run).
// Run: node test/booking.test.js
const fs = require('fs');
const vm = require('vm');
const assert = require('assert');
const path = require('path');
const { makeSheet } = require('./sheetmock');

function makeEnv() {
  const sheets = {};
  const mails = [];
  const triggers = [];
  const ss = {
    getSheetByName: (n) => sheets[n] || null,
    insertSheet: (n) => (sheets[n] = makeSheet()),
  };
  const ctx = {
    SpreadsheetApp: { getActive: () => ss },
    MailApp: { sendEmail: (to, subject, body) => mails.push({ to, subject, body }) },
    ScriptApp: {
      getProjectTriggers: () => triggers,
      deleteTrigger: (t) => triggers.splice(triggers.indexOf(t), 1),
      newTrigger: (fn) => ({ forSpreadsheet: () => ({ onFormSubmit: () => ({ create: () => triggers.push({ getHandlerFunction: () => fn }) }) }) }),
    },
    Date,
  };
  vm.createContext(ctx);
  vm.runInContext(fs.readFileSync(path.join(__dirname, '../_includes/code/booking-auto-confirm.gs'), 'utf8'), ctx);
  return { ctx, sheets, mails, triggers };
}

const answer = (o) => ({ namedValues: Object.fromEntries(Object.entries(o).map(([k, v]) => [k, [v]])) });
const base = { 이름: '김손님', 연락처: '010-0000-0000', 이메일: 'guest@example.com', '희망 날짜': '2026-12-05', '희망 시간': '14:00', 인원: '2' };

let n = 0;
function test(name, fn) { fn(); n++; console.log('ok -', name); }

test('첫 예약은 확정, 시트 생성 + 헤더', () => {
  const { ctx, sheets, mails } = makeEnv();
  assert.strictEqual(ctx.onFormSubmit(answer(base)), '확정');
  const rows = sheets['예약목록'].rows;
  assert.strictEqual(JSON.stringify(rows[0]), JSON.stringify(['접수 시각', '이름', '연락처', '이메일', '날짜', '시간', '인원', '상태']));
  assert.strictEqual(rows[1][7], '확정');
  assert.strictEqual(mails.length, 2);
  assert.ok(mails[0].subject.includes('확정'));
  assert.strictEqual(mails[1].to, 'owner@example.com');
});

test('정원 4명: 2+2는 확정, 그다음 1명은 대기', () => {
  const { ctx } = makeEnv();
  assert.strictEqual(ctx.onFormSubmit(answer(base)), '확정');
  assert.strictEqual(ctx.onFormSubmit(answer({ ...base, 이름: '이손님' })), '확정');
  assert.strictEqual(ctx.onFormSubmit(answer({ ...base, 이름: '박손님', 인원: '1' })), '대기');
});

test('대기 인원은 정원 계산에 들어가지 않음', () => {
  const { ctx } = makeEnv();
  ctx.onFormSubmit(answer({ ...base, 인원: '4' }));                 // 확정 4
  assert.strictEqual(ctx.onFormSubmit(answer({ ...base, 인원: '3' })), '대기');
  assert.strictEqual(ctx.onFormSubmit(answer({ ...base, 인원: '1' })), '대기'); // 대기 3명은 안 셈 → 4+1>4
});

test('다른 시간은 따로 계산', () => {
  const { ctx } = makeEnv();
  ctx.onFormSubmit(answer({ ...base, 인원: '4' }));
  assert.strictEqual(ctx.onFormSubmit(answer({ ...base, '희망 시간': '16:00', 인원: '4' })), '확정');
});

test('인원 "3명" 같은 글자도 숫자로 읽고, 비어 있으면 1명', () => {
  const { ctx, sheets } = makeEnv();
  ctx.onFormSubmit(answer({ ...base, 인원: '3명' }));
  ctx.onFormSubmit(answer({ ...base, '희망 시간': '10:00', 인원: '' }));
  assert.strictEqual(sheets['예약목록'].rows[1][6], 3);
  assert.strictEqual(sheets['예약목록'].rows[2][6], 1);
});

test('이메일이 없으면 손님 메일은 건너뛰고 사장님 알림만', () => {
  const { ctx, mails } = makeEnv();
  ctx.onFormSubmit(answer({ ...base, 이메일: '' }));
  assert.strictEqual(mails.length, 1);
  assert.strictEqual(mails[0].to, 'owner@example.com');
});

test('setupTrigger는 여러 번 실행해도 트리거 1개', () => {
  const { ctx, triggers } = makeEnv();
  ctx.setupTrigger(); ctx.setupTrigger();
  assert.strictEqual(triggers.length, 1);
});

test('실제 시트처럼 날짜·시간이 자동 변환된 기존 줄도 정원 계산 (2026-11-26 실계정 버그 재현)', () => {
  const { ctx, sheets } = makeEnv();
  const sh = (sheets['예약목록'] = makeSheet());           // 고치기 전 코드로 쌓인 시트: 서식 없음
  sh.appendRow(['접수 시각', '이름', '연락처', '이메일', '날짜', '시간', '인원', '상태']);
  sh.appendRow([new Date(), '김하늘', '010', 'a@x.com', '2026. 12. 5', '14:00', 2, '확정']);
  assert.ok(sh.rows[1][4] instanceof Date && sh.rows[1][5] instanceof Date); // 변환돼 저장됨
  // 폼은 '2026. 12. 5' 모양으로 보냄(실계정 확인). 2+3 > 4 → 대기
  assert.strictEqual(ctx.onFormSubmit(answer({ ...base, '희망 날짜': '2026. 12. 5', 인원: '3' })), '대기');
});

test('날짜·시간은 글자 그대로 저장되고, 날짜 모양이 달라도 같은 예약으로 셈', () => {
  const { ctx, sheets } = makeEnv();
  ctx.onFormSubmit(answer({ ...base, '희망 날짜': '2026. 12. 5', '희망 시간': '14:00', 인원: '3' }));
  assert.strictEqual(sheets['예약목록'].rows[1][4], '2026. 12. 5');
  assert.strictEqual(sheets['예약목록'].rows[1][5], '14:00');
  assert.strictEqual(ctx.onFormSubmit(answer({ ...base, '희망 날짜': '2026-12-05', '희망 시간': '오후 2:00', 인원: '2' })), '대기');
});

console.log(`${n} passed`);
