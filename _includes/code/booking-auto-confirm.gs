/**
 * 구글폼 예약 → 시트 정리 → 손님 확인 메일 + 사장님 알림
 * 레시피: https://dahae.superblog.kr/
 *
 * 폼 질문 제목은 아래 FIELDS와 똑같아야 합니다. 다르면 FIELDS 쪽 글자를 폼에 맞게 고치세요.
 */
const CONFIG = {
  SHOP_NAME: '○○공방',                 // 메일에 들어갈 가게 이름
  OWNER_EMAIL: 'owner@example.com',     // 새 예약 알림을 받을 사장님 메일
  SHEET_NAME: '예약목록',               // 정리용 시트 이름 (없으면 자동으로 만듭니다)
  CAPACITY_PER_SLOT: 4,                 // 같은 날짜·시간에 받을 수 있는 최대 인원
};

const FIELDS = {
  name: '이름',
  phone: '연락처',
  email: '이메일',
  date: '희망 날짜',
  time: '희망 시간',
  people: '인원',
};

const HEADER = ['접수 시각', '이름', '연락처', '이메일', '날짜', '시간', '인원', '상태'];

/** 폼이 제출될 때마다 실행됩니다 (setupTrigger로 한 번 연결). */
function onFormSubmit(e) {
  const v = readAnswers_(e.namedValues);
  const sheet = getSheet_();
  const booked = countPeople_(sheet, v.date, v.time);
  const status = booked + v.people <= CONFIG.CAPACITY_PER_SLOT ? '확정' : '대기';

  // 날짜·시간 앞의 ' 는 시트가 날짜 값으로 바꾸지 않고 글자 그대로 두게 합니다 (시트에는 안 보임)
  sheet.appendRow([new Date(), v.name, v.phone, v.email, "'" + v.date, "'" + v.time, v.people, status]);

  if (v.email) {
    MailApp.sendEmail(v.email, subjectFor_(status), bodyFor_(status, v));
  }
  MailApp.sendEmail(
    CONFIG.OWNER_EMAIL,
    `[새 예약·${status}] ${v.date} ${v.time} ${v.name} ${v.people}명`,
    `${v.name}님 (${v.phone})\n${v.date} ${v.time}, ${v.people}명\n상태: ${status}\n` +
      `같은 시간 확정 인원: ${status === '확정' ? booked + v.people : booked}/${CONFIG.CAPACITY_PER_SLOT}`
  );
  return status;
}

/** 처음 한 번만 실행: 폼 제출 트리거를 연결합니다. */
function setupTrigger() {
  const ss = SpreadsheetApp.getActive();
  ScriptApp.getProjectTriggers()
    .filter((t) => t.getHandlerFunction() === 'onFormSubmit')
    .forEach((t) => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('onFormSubmit').forSpreadsheet(ss).onFormSubmit().create();
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
    .filter((r) => dateKey_(r[4]) === dateKey_(date) && timeKey_(r[5]) === timeKey_(time) && r[7] === '확정')
    .reduce((sum, r) => sum + Number(r[6] || 0), 0);
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
