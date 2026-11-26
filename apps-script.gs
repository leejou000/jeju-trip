// 제주 여행 투표 저장소. 구글 시트의 확장 프로그램 > Apps Script에 통째로 붙여넣는다.
// 시트 한 줄 = 한 사람: A열 멤버 키, B열 투표(JSON), C열 마지막 저장 시각
const KEYS = ["prof", "daun", "jh", "sh", "camila", "sy", "jw", "yw", "jiu"];
const POLLS = ["d1h", "d2a", "d2d", "d2g", "d3a", "d3d", "d3g"];
const OPTS = ["a", "b"];

function sheet_() { return SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]; }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function doGet() {
  const out = {};
  sheet_().getDataRange().getValues().forEach(r => {
    if (KEYS.indexOf(r[0]) >= 0) { try { out[r[0]] = JSON.parse(r[1]); } catch (e) {} }
  });
  return json_({ ok: true, picks: out });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const b = JSON.parse(e.postData.contents);
    if (KEYS.indexOf(b.who) < 0) return json_({ ok: false });
    const p = {};
    POLLS.forEach(k => { if (b.picks && OPTS.indexOf(b.picks[k]) >= 0) p[k] = b.picks[k]; });
    const sh = sheet_();
    const row = sh.getDataRange().getValues().findIndex(r => r[0] === b.who) + 1;
    if (row) sh.getRange(row, 2, 1, 2).setValues([[JSON.stringify(p), new Date()]]);
    else sh.appendRow([b.who, JSON.stringify(p), new Date()]);
    return json_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}
