/**
 * 구글폼 예약 → 확정/대기 메일 + 확정 예약을 구글 캘린더에 자동 등록
 * 레시피: https://dahae.superblog.kr/
 *
 * 1편(예약 확인 메일)의 코드를 이 코드로 통째로 바꾸면 됩니다.
 * 폼 질문 제목은 FIELDS와 똑같아야 합니다.
 */
const CONFIG = {
  SHOP_NAME: '○○공방',
  OWNER_EMAIL: 'owner@example.com',
  SHEET_NAME: '예약목록',
  CAPACITY_PER_SLOT: 4,
  CALENDAR_ID: '',          // 비워 두면 내 기본 캘린더. 따로 만든 캘린더면 설정의 '캘린더 ID'를 넣으세요
  SLOT_MINUTES: 120,        // 수업 한 타임 길이(분)
};

const FIELDS = {
  name: '이름',
  phone: '연락처',
  email: '이메일',
  date: '희망 날짜',
  time: '희망 시간',
  people: '인원',
};

const HEADER = ['접수 시각', '이름', '연락처', '이메일', '날짜', '시간', '인원', '상태', '일정ID'];
const COL = { date: 5, time: 6, people: 7, status: 8, eventId: 9 }; // 1부터 센 열 번호

function onFormSubmit(e) {
  const v = readAnswers_(e.namedValues);
  const sheet = getSheet_();
  const booked = countPeople_(sheet, v.date, v.time);
  const status = booked + v.people <= CONFIG.CAPACITY_PER_SLOT ? '확정' : '대기';

  // 날짜·시간 앞의 ' 는 시트가 날짜 값으로 바꾸지 않고 글자 그대로 두게 합니다 (시트에는 안 보임)
  sheet.appendRow([new Date(), v.name, v.phone, v.email, "'" + v.date, "'" + v.time, v.people, status, '']);
  const row = sheet.getLastRow();

  if (status === '확정') {
    const id = createEvent_(v);
    if (id) sheet.getRange(row, COL.eventId).setValue(id);
  }

  if (v.email) MailApp.sendEmail(v.email, subjectFor_(status), bodyFor_(status, v));
  MailApp.sendEmail(
    CONFIG.OWNER_EMAIL,
    `[새 예약·${status}] ${v.date} ${v.time} ${v.name} ${v.people}명`,
    `${v.name}님 (${v.phone})\n${v.date} ${v.time}, ${v.people}명\n상태: ${status}`
  );
  return status;
}

/**
 * 시트를 손으로 고친 뒤 실행하세요 (또는 하루 한 번 시간 트리거로).
 * - 상태가 '확정'인데 일정이 없으면 만들고
 * - 상태가 '취소'인데 일정이 남아 있으면 지웁니다.
 */
function syncCalendar() {
  const sheet = getSheet_();
  const rows = sheet.getDataRange().getDisplayValues(); // 시트에 보이는 글자 그대로 읽기
  let created = 0;
  let removed = 0;
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const status = r[COL.status - 1];
    const id = r[COL.eventId - 1];
    if (status === '확정' && !id) {
      const newId = createEvent_({ name: r[1], phone: r[2], date: r[COL.date - 1], time: r[COL.time - 1], people: Number(r[COL.people - 1]) });
      if (newId) {
        sheet.getRange(i + 1, COL.eventId).setValue(newId);
        created++;
      }
    } else if (status === '취소' && id) {
      const ev = getCalendar_().getEventById(id);
      if (ev) ev.deleteEvent();
      sheet.getRange(i + 1, COL.eventId).setValue('');
      removed++;
    }
  }
  return { created, removed };
}

function setupTrigger() {
  const ss = SpreadsheetApp.getActive();
  ScriptApp.getProjectTriggers()
    .filter((t) => t.getHandlerFunction() === 'onFormSubmit')
    .forEach((t) => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('onFormSubmit').forSpreadsheet(ss).onFormSubmit().create();
}

function createEvent_(v) {
  const start = parseDateTime_(v.date, v.time);
  if (!start) return '';
  const end = new Date(start.getTime() + CONFIG.SLOT_MINUTES * 60 * 1000);
  const ev = getCalendar_().createEvent(`${v.name} ${v.people}명`, start, end, {
    description: `연락처: ${v.phone}\n예약 시트에서 자동 등록`,
  });
  return ev.getId();
}

function getCalendar_() {
  return CONFIG.CALENDAR_ID ? CalendarApp.getCalendarById(CONFIG.CALENDAR_ID) : CalendarApp.getDefaultCalendar();
}

/** 날짜·시간 글자를 Date로 바꿉니다. 읽는 모양은 dateKey_, timeKey_ 설명 참고. */
function parseDateTime_(date, time) {
  const d = dateKey_(date).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  const t = timeKey_(time).match(/^(\d{2}):(\d{2})$/);
  if (!d || !t) return null;
  return new Date(Number(d[1]), Number(d[2]) - 1, Number(d[3]), Number(t[1]), Number(t[2]));
}

function readAnswers_(nv) {
  const get = (key) => ((nv[FIELDS[key]] || [''])[0] || '').toString().trim();
  const people = parseInt(get('people').replace(/[^0-9]/g, ''), 10);
  return {
    name: get('name'),
    phone: get('phone'),
    email: get('email'),
    date: get('date'),
    time: get('time'),
    people: isNaN(people) || people < 1 ? 1 : people,
  };
}

function getSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    sheet.appendRow(HEADER);
  }
  return sheet;
}

/** 같은 날짜·시간에 이미 '확정'된 인원 합계 (시트에 보이는 값으로 비교) */
function countPeople_(sheet, date, time) {
  const rows = sheet.getDataRange().getDisplayValues().slice(1);
  return rows
    .filter((r) => dateKey_(r[COL.date - 1]) === dateKey_(date) && timeKey_(r[COL.time - 1]) === timeKey_(time) && r[COL.status - 1] === '확정')
    .reduce((sum, r) => sum + Number(r[COL.people - 1] || 0), 0);
}

/**
 * 구글 시트는 '2026. 12. 5', '14:00' 같은 글자를 날짜·시간 값으로 바꿔 저장합니다.
 * 그래도 같은 예약끼리 비교되도록 'YYYY-MM-DD', 'HH:MM' 모양으로 맞춥니다.
 * 날짜는 '2026-12-05', '2026. 12. 5', '12/5/2026', 시간은 '14:00', '오후 2:00', '14시'를 읽습니다.
 */
function dateKey_(s) {
  const str = String(s);
  let m = str.match(/(\d{4})\D+(\d{1,2})\D+(\d{1,2})/);
  if (m) return `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}`;
  m = str.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return `${m[3]}-${m[1].padStart(2, '0')}-${m[2].padStart(2, '0')}`;
  return str.trim();
}

function timeKey_(s) {
  const str = String(s);
  const t = str.match(/(\d{1,2})(?:\s*[:시]\s*(\d{1,2}))?/);
  if (!t) return str.trim();
  let h = Number(t[1]);
  if (/오후|PM/i.test(str) && h < 12) h += 12;
  if (/오전|AM/i.test(str) && h === 12) h = 0;
  return `${String(h).padStart(2, '0')}:${String(t[2] || 0).padStart(2, '0')}`;
}

function subjectFor_(status) {
  return status === '확정'
    ? `[${CONFIG.SHOP_NAME}] 예약이 확정됐어요`
    : `[${CONFIG.SHOP_NAME}] 예약 대기로 접수됐어요`;
}

function bodyFor_(status, v) {
  const head = `${v.name}님, ${CONFIG.SHOP_NAME}입니다.\n\n`;
  const detail = `날짜: ${v.date}\n시간: ${v.time}\n인원: ${v.people}명\n\n`;
  if (status === '확정') {
    return head + '예약이 확정됐어요.\n\n' + detail + '일정이 바뀌면 이 메일로 답장 주세요.';
  }
  return head + '말씀하신 시간은 자리가 다 찼어요. 대기로 접수해 두었고, 자리가 나면 먼저 연락드릴게요.\n\n' + detail;
}
