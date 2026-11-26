// 제주 여행 투표와 게시판 저장소. 구글 시트의 확장 프로그램 > Apps Script에 통째로 붙여넣는다.
// 첫 번째 시트: 한 줄 = 한 사람. A열 멤버 키, B열 투표(JSON), C열 마지막 저장 시각
// 「게시판」 시트: 한 줄 = 글 하나. A 시각, B 멤버 키, C 날짜(d1, d2, d3, 빈칸=전체), D 내용, E 답(시트에서 직접 적으면 페이지에 보임)
const KEYS = ["prof", "daun", "jh", "sh", "camila", "sy", "jw", "yw", "jiu", "jc"];
// 일정 ID와 선택지 ID는 모양만 검사한다. 페이지에 투표가 새로 생겨도 다시 배포할 필요가 없게
const POLL_RE = /^d[1-3][a-z0-9]{1,4}$/;
const OPT_RE = /^[a-h]$/;
const DAYS = ["", "d1", "d2", "d3"];

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sheet_() { return ss_().getSheets()[0]; }
function board_() { return ss_().getSheetByName("게시판") || ss_().insertSheet("게시판", ss_().getNumSheets()); }
// 「채팅」 시트: 한 줄 = 메시지 하나. A 시각, B 멤버 키, C 내용. 페이지에는 마지막 200개만 보낸다
function chat_() { return ss_().getSheetByName("채팅") || ss_().insertSheet("채팅", ss_().getNumSheets()); }
const safe_ = t => /^[=+\-@]/.test(t) ? "'" + t : t;
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function doGet() {
  const out = {};
  sheet_().getDataRange().getValues().forEach(r => {
    if (KEYS.indexOf(r[0]) >= 0) { try { out[r[0]] = JSON.parse(r[1]); } catch (e) {} }
  });
  const posts = board_().getDataRange().getValues()
    .filter(r => KEYS.indexOf(r[1]) >= 0 && r[3])
    .map(r => ({ at: r[0] instanceof Date ? r[0].getTime() : 0, who: r[1], day: String(r[2] || ""), text: String(r[3]), reply: String(r[4] || "") }));
  // 채팅은 시각 키를 t로 (게시판 글의 at과 섞이지 않게)
  const chat = chat_().getDataRange().getValues()
    .filter(r => KEYS.indexOf(r[1]) >= 0 && r[2])
    .slice(-200)
    .map(r => ({ t: r[0] instanceof Date ? r[0].getTime() : 0, who: r[1], text: String(r[2]) }));
  return json_({ ok: true, picks: out, posts: posts, chat: chat });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const b = JSON.parse(e.postData.contents);
    if (KEYS.indexOf(b.who) < 0) return json_({ ok: false });
    if (b.type === "post") {
      const text = String(b.text || "").trim().slice(0, 500);
      if (!text) return json_({ ok: false });
      // =, +, -, @로 시작하면 시트가 수식으로 읽지 않게 앞에 ' 를 붙인다
      board_().appendRow([new Date(), b.who, DAYS.indexOf(b.day) >= 0 ? b.day : "", /^[=+\-@]/.test(text) ? "'" + text : text, ""]);
      return json_({ ok: true });
    }
    if (b.type === "chat") {
      const text = String(b.text || "").trim().slice(0, 300);
      if (!text) return json_({ ok: false });
      chat_().appendRow([new Date(), b.who, safe_(text)]);
      return json_({ ok: true });
    }
    const p = {};
    Object.keys(b.picks || {}).slice(0, 40).forEach(k => { if (POLL_RE.test(k) && OPT_RE.test(String(b.picks[k]))) p[k] = b.picks[k]; });
    const sh = sheet_();
    const row = sh.getDataRange().getValues().findIndex(r => r[0] === b.who) + 1;
    if (row) sh.getRange(row, 2, 1, 2).setValues([[JSON.stringify(p), new Date()]]);
    else sh.appendRow([b.who, JSON.stringify(p), new Date()]);
    return json_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}
