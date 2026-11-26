// Logic test for booking-with-calendar.gs with mocked Google services (not a live Apps Script run).
// Run: node test/calendar.test.js
const fs = require('fs');
const vm = require('vm');
const assert = require('assert');
const path = require('path');
const { makeSheet } = require('./sheetmock');

function makeEnv() {
  const sheets = {};
  const mails = [];
  const events = new Map();
  let seq = 0;
  const ss = { getSheetByName: (n) => sheets[n] || null, insertSheet: (n) => (sheets[n] = makeSheet()) };
  const cal = {
    createEvent: (title, start, end, opts) => {
      const id = 'ev' + ++seq;
      events.set(id, { title, start, end, opts, deleteEvent: () => events.delete(id) });
      return { getId: () => id };
    },
    getEventById: (id) => events.get(id) || null,
  };
  const ctx = {
    SpreadsheetApp: { getActive: () => ss },
    MailApp: { sendEmail: (to, subject, body) => mails.push({ to, subject, body }) },
    CalendarApp: { getDefaultCalendar: () => cal, getCalendarById: () => cal },
    ScriptApp: { getProjectTriggers: () => [], deleteTrigger: () => {}, newTrigger: () => ({ forSpreadsheet: () => ({ onFormSubmit: () => ({ create: () => {} }) }) }) },
    Date,
  };
  vm.createContext(ctx);
  vm.runInContext(fs.readFileSync(path.join(__dirname, '../_includes/code/booking-with-calendar.gs'), 'utf8'), ctx);
  return { ctx, sheets, mails, events };
}

const answer = (o) => ({ namedValues: Object.fromEntries(Object.entries(o).map(([k, v]) => [k, [v]])) });
const base = { 이름: '김손님', 연락처: '010-0000-0000', 이메일: 'guest@example.com', '희망 날짜': '2026-12-05', '희망 시간': '14:00', 인원: '2' };

let n = 0;
function test(name, fn) { fn(); n++; console.log('ok -', name); }

test('확정 예약은 캘린더에 2시간 일정 생성, 시트에 일정ID 기록', () => {
  const { ctx, sheets, events } = makeEnv();
  assert.strictEqual(ctx.onFormSubmit(answer(base)), '확정');
  assert.strictEqual(events.size, 1);
  const ev = [...events.values()][0];
  assert.strictEqual(ev.title, '김손님 2명');
  assert.strictEqual(ev.start.getFullYear(), 2026);
  assert.strictEqual(ev.start.getMonth(), 11);
  assert.strictEqual(ev.start.getDate(), 5);
  assert.strictEqual(ev.start.getHours(), 14);
  assert.strictEqual((ev.end - ev.start) / 60000, 120);
  assert.strictEqual(sheets['예약목록'].rows[1][8], 'ev1');
});

test('대기 예약은 일정을 만들지 않음', () => {
  const { ctx, events } = makeEnv();
  ctx.onFormSubmit(answer({ ...base, 인원: '4' }));
  assert.strictEqual(ctx.onFormSubmit(answer({ ...base, 인원: '1' })), '대기');
  assert.strictEqual(events.size, 1);
});

test('날짜·시간 여러 표기 읽기', () => {
  const { ctx } = makeEnv();
  const cases = [
    ['2026. 12. 5', '오후 2:00', 14, 0],
    ['2026/12/05', '14시', 14, 0],
    ['2026-12-05', '오전 9:30', 9, 30],
    ['2026-12-05', '오후 12:00', 12, 0],
    ['2026-12-05', '오전 12:00', 0, 0],
  ];
  for (const [d, t, h, m] of cases) {
    const dt = ctx.parseDateTime_(d, t);
    assert.ok(dt, `${d} ${t}`);
    assert.strictEqual(dt.getDate(), 5);
    assert.strictEqual(dt.getHours(), h, `${t}`);
    assert.strictEqual(dt.getMinutes(), m, `${t}`);
  }
  assert.strictEqual(ctx.parseDateTime_('다음주', '14:00'), null);
});

test('날짜를 못 읽으면 일정 없이 예약만 기록', () => {
  const { ctx, sheets, events } = makeEnv();
  ctx.onFormSubmit(answer({ ...base, '희망 날짜': '다음주 토요일' }));
  assert.strictEqual(events.size, 0);
  assert.strictEqual(sheets['예약목록'].rows[1][8], '');
});

test('syncCalendar: 취소로 바꾼 예약의 일정 삭제, 확정인데 일정 없으면 생성', () => {
  const { ctx, sheets, events } = makeEnv();
  ctx.onFormSubmit(answer(base));                                   // ev1
  ctx.onFormSubmit(answer({ ...base, '희망 시간': '16:00' }));       // ev2
  const rows = sheets['예약목록'].rows;
  rows[1][7] = '취소';                                              // 사장님이 손으로 취소
  rows.push([new Date(), '박손님', '010', '', '2026-12-06', '10:00', 1, '확정', '']); // 손으로 추가
  const r = ctx.syncCalendar();
  assert.strictEqual(r.created, 1);
  assert.strictEqual(r.removed, 1);
  assert.ok(!events.has('ev1'));
  assert.ok(events.has('ev2'));
  assert.strictEqual(rows[1][8], '');
  assert.strictEqual(rows[3][8], 'ev3');
});

test('syncCalendar를 두 번 돌려도 일정이 늘지 않음', () => {
  const { ctx, events } = makeEnv();
  ctx.onFormSubmit(answer(base));
  ctx.syncCalendar(); ctx.syncCalendar();
  assert.strictEqual(events.size, 1);
});

test('취소된 인원은 정원 계산에서 빠짐', () => {
  const { ctx, sheets } = makeEnv();
  ctx.onFormSubmit(answer({ ...base, 인원: '4' }));
  sheets['예약목록'].rows[1][7] = '취소';
  assert.strictEqual(ctx.onFormSubmit(answer({ ...base, 인원: '4' })), '확정');
});

test('실제 시트처럼 날짜·시간이 변환돼도 syncCalendar가 제 시각에 일정 생성 (2026-11-26 실계정 버그 재현)', () => {
  const { ctx, sheets, events } = makeEnv();
  const sh = (sheets['예약목록'] = makeSheet());           // 고치기 전 코드로 쌓인 시트: 서식 없음
  sh.appendRow(['접수 시각', '이름', '연락처', '이메일', '날짜', '시간', '인원', '상태', '일정ID']);
  sh.appendRow([new Date(), '김하늘', '010', 'a@x.com', '2026. 12. 5', '14:00', 2, '확정', '']);
  assert.ok(sh.rows[1][4] instanceof Date);
  const r = ctx.syncCalendar();
  assert.strictEqual(r.created, 1);
  const ev = [...events.values()][0];
  assert.strictEqual(ev.start.getFullYear(), 2026); assert.strictEqual(ev.start.getMonth(), 11);
  assert.strictEqual(ev.start.getDate(), 5); assert.strictEqual(ev.start.getHours(), 14); assert.strictEqual(ev.start.getMinutes(), 0);
});

test('변환된 기존 줄이 있어도 정원 계산 (2+3 > 4 → 대기)', () => {
  const { ctx, sheets } = makeEnv();
  const sh = (sheets['예약목록'] = makeSheet());
  sh.appendRow(['접수 시각', '이름', '연락처', '이메일', '날짜', '시간', '인원', '상태', '일정ID']);
  sh.appendRow([new Date(), '김하늘', '010', 'a@x.com', '2026. 12. 5', '14:00', 2, '확정', 'evX']);
  assert.strictEqual(ctx.onFormSubmit(answer({ ...base, '희망 날짜': '2026. 12. 5', 인원: '3' })), '대기');
});

console.log(`${n} passed`);
