// The generator must produce code that still runs and uses the entered values (mocked Google services).
// Run: node test/generator.test.js
const fs = require('fs'), vm = require('vm'), assert = require('assert'), path = require('path');
const { makeSheet } = require('./sheetmock');
const { buildBookingCode } = require('../assets/js/booking-generator.js');
const read = (f) => fs.readFileSync(path.join(__dirname, '../_includes/code', f), 'utf8');

function run(code) {
  const sheets = {}, mails = [], events = [];
  const ss = { getSheetByName: (n) => sheets[n] || null, insertSheet: (n) => (sheets[n] = makeSheet()) };
  const cal = { createEvent: (t, s, e, o) => { events.push({ t, s, e }); return { getId: () => 'ev' + events.length }; }, getEventById: () => null };
  const ctx = { SpreadsheetApp: { getActive: () => ss }, MailApp: { sendEmail: (to, subject, body) => mails.push({ to, subject, body }) },
    CalendarApp: { getDefaultCalendar: () => cal, getCalendarById: () => cal },
    ScriptApp: { getProjectTriggers: () => [], deleteTrigger() {}, newTrigger: () => ({ forSpreadsheet: () => ({ onFormSubmit: () => ({ create() {} }) }) }) },
    Session: { getScriptTimeZone: () => 'Asia/Seoul' }, Utilities: { formatDate: (d) => String(d) }, Logger: { log() {} }, console, Date };
  vm.createContext(ctx); vm.runInContext(code, ctx);
  return { ctx, sheets, mails, events };
}
const cfg = { shopName: "냥이's 미용실", ownerEmail: 'boss@example.com', sheetName: '예약 정리', capacity: 2, slotMinutes: 60,
  fields: { name: '성함', phone: '전화', email: '메일 주소', date: '날짜', time: '시간', people: '인원수' } };
const answers = (name, people) => ({ namedValues: { '성함': [name], '전화': ['-'], '메일 주소': [name + '@example.com'], '날짜': ['2026. 12. 5'], '시간': ['14:00'], '인원수': [String(people)] } });

let n = 0;
for (const file of ['booking-auto-confirm.gs', 'booking-with-calendar.gs']) {
  const tpl = read(file), code = buildBookingCode(tpl, cfg);
  assert.ok(code.includes("SHOP_NAME: '냥이\\'s 미용실'"), 'shop name escaped'); n++;
  assert.ok(/CAPACITY_PER_SLOT: 2\b/.test(code)); n++;
  assert.ok(code.includes("name: '성함'") && code.includes("people: '인원수'")); n++;
  assert.strictEqual(code.split('\n').length, tpl.split('\n').length, 'line count unchanged'); n++;
  const env = run(code);
  env.ctx.onFormSubmit(answers('가', 2));
  env.ctx.onFormSubmit(answers('나', 1));
  assert.ok(env.sheets['예약 정리'], 'custom sheet name used'); n++;
  const toGuest = env.mails.filter((m) => m.to.endsWith('@example.com') && m.to !== 'boss@example.com');
  assert.ok(toGuest[0].subject.includes("냥이's 미용실") && toGuest[0].subject.includes('확정'), toGuest[0].subject); n++;
  assert.ok(toGuest[1].subject.includes('대기'), 'capacity 2: second booking waits'); n++;
  assert.ok(env.mails.some((m) => m.to === 'boss@example.com'), 'owner notified'); n++;
}
assert.throws(() => buildBookingCode(read('booking-auto-confirm.gs'), Object.assign({}, cfg, { capacity: 0 }))); n++;
const evil = buildBookingCode(read('booking-auto-confirm.gs'), Object.assign({}, cfg, { shopName: "x'; MailApp.sendEmail('a','b','c'); '" }));
const e2 = run(evil); assert.strictEqual(e2.mails.length, 0, 'injected text stays a string'); n++;
console.log('generator tests passed:', n);
