// 예약 자동화 코드 만들기 — 입력값을 레시피 코드의 CONFIG·FIELDS에 넣어 돌려준다. 브라우저 안에서만 동작한다.
(function (root) {
  function q(s) { // 작은따옴표 문자열 안에 안전하게 넣기
    return String(s).replace(/\\/g, '\\\\').replace(/'/g, "\\'").replace(/[\r\n]+/g, ' ').trim();
  }
  function setStr(code, block, key, value) {
    var re = new RegExp('(const ' + block + ' = \\{[\\s\\S]*?\\b' + key + ":\\s*)'[^'\\n]*'");
    if (!re.test(code)) throw new Error(block + '.' + key + ' 자리를 찾지 못했어요');
    return code.replace(re, function (m, head) { return head + "'" + q(value) + "'"; });
  }
  function setNum(code, block, key, value) {
    var n = parseInt(value, 10);
    if (!(n > 0)) throw new Error(key + ' 값은 1 이상의 숫자여야 해요');
    var re = new RegExp('(const ' + block + ' = \\{[\\s\\S]*?\\b' + key + ':\\s*)\\d+');
    if (!re.test(code)) throw new Error(block + '.' + key + ' 자리를 찾지 못했어요');
    return code.replace(re, function (m, head) { return head + n; });
  }
  function buildBookingCode(template, cfg) {
    var c = template;
    c = setStr(c, 'CONFIG', 'SHOP_NAME', cfg.shopName);
    c = setStr(c, 'CONFIG', 'OWNER_EMAIL', cfg.ownerEmail);
    c = setStr(c, 'CONFIG', 'SHEET_NAME', cfg.sheetName || '예약목록');
    c = setNum(c, 'CONFIG', 'CAPACITY_PER_SLOT', cfg.capacity);
    if (/SLOT_MINUTES/.test(c) && cfg.slotMinutes) c = setNum(c, 'CONFIG', 'SLOT_MINUTES', cfg.slotMinutes);
    var f = cfg.fields || {};
    ['name', 'phone', 'email', 'date', 'time', 'people'].forEach(function (k) {
      if (f[k]) c = setStr(c, 'FIELDS', k, f[k]);
    });
    return c;
  }
  root.buildBookingCode = buildBookingCode;
  if (typeof module !== 'undefined' && module.exports) module.exports = { buildBookingCode: buildBookingCode };
})(typeof window !== 'undefined' ? window : globalThis);
