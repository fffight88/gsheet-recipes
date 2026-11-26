// Sheet mock that behaves like real Google Sheets for this recipe (observed in a real sheet, 2026-11-26):
// strings like '2026-12-05' / '2026. 12. 5' / '14:00' written into E (date) / F (time) are stored as
// Date values and displayed as '2026. 12. 5' / '14:00' (ko_KR). A leading apostrophe keeps the text
// as-is (the apostrophe itself is not stored/shown). setNumberFormat('@') is NOT modelled: in the real
// sheet appendRow overrode it on the next row.
function makeSheet() {
  const rows = [];
  const store = (v, c) => {
    if (typeof v !== 'string') return v;
    if (v.startsWith("'")) return v.slice(1);
    let m;
    if (c === 4 && (m = v.match(/^(\d{4})\D+(\d{1,2})\D+(\d{1,2})\.?$/))) return new Date(+m[1], +m[2] - 1, +m[3]);
    if (c === 5 && (m = v.match(/^(\d{1,2}):(\d{2})$/))) return new Date(1899, 11, 30, +m[1], +m[2]);
    return v;
  };
  const show = (v, c) => {
    if (v instanceof Date && c === 4) return `${v.getFullYear()}. ${v.getMonth() + 1}. ${v.getDate()}`;
    if (v instanceof Date && c === 5) return `${v.getHours()}:${String(v.getMinutes()).padStart(2, '0')}`;
    return v === undefined || v === null ? '' : String(v);
  };
  return {
    rows,
    appendRow: (r) => rows.push(r.map(store)),
    getLastRow: () => rows.length,
    getRange: (r, c) => ({ setValue: (v) => { rows[r - 1][c - 1] = store(v, c - 1); } }),
    getDataRange: () => ({
      getValues: () => rows.map((r) => r.slice()),
      getDisplayValues: () => rows.map((r) => r.map(show)),
    }),
  };
}
module.exports = { makeSheet };
