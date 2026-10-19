/**
 * 청첩장 서버 (Google Apps Script) — 방명록 + 참석 여부(RSVP)
 *
 * 이 코드는 "하객명단" 스프레드시트에 붙인다. 같은 파일 안에
 *   - 방명록   탭: 하객 축하 글
 *   - 참석응답 탭: 참석 여부 응답
 * 두 탭이 추가되고, 기존 "하객명단","집계" 탭은 건드리지 않는다.
 *
 * 준비 순서
 *  1. 하객명단 시트 메뉴 [확장 프로그램] -> [Apps Script] 에 이 코드를 통째로 붙여넣는다.
 *  2. 아래 SALT 값을 아무 긴 문자열로 바꾼다 (한 번 정하면 바꾸지 않는다).
 *  3. 위쪽 함수 선택 상자에서 setup 을 고르고 ▶ 실행 -> 권한 허용. 탭 2개와 머리글이 자동으로 만들어진다.
 *  4. [배포] -> [새 배포] -> 유형 "웹 앱" / 실행 사용자: 나 / 액세스 권한: 모든 사용자 -> 배포
 *     -> 나온 .../exec 주소를 index.html 의 SHEET_URL 에 넣는다.
 *  * 코드를 고친 뒤에는 [배포] -> [배포 관리] -> 새 버전으로 다시 배포해야 반영된다.
 */

var SALT = "안녕하세요경민입니다잘부탁드립니다";
var MASTER_PASSWORD = "qwer1234";   // 신랑,신부용 — 어떤 글이든 지울 수 있다. 이 값은 서버(이 파일)에만 두고 index.html에는 절대 넣지 않는다.
var GB_SHEET = "방명록";
var RSVP_SHEET = "참석응답";

var MAX_NAME = 20, MAX_MSG = 500, MAX_MEMO = 200;
var RATE_LIMIT_SEC = 20;    // 같은 이름으로 연속 작성 제한(초)
var GLOBAL_PER_MIN = 30;    // 전체 분당 작성 한도 (스팸 방지)

var GB_HEADER = ["id", "작성시각", "이름", "내용", "비밀번호해시"];
var RSVP_HEADER = ["응답시각", "이름", "구분", "참석", "인원(본인포함)", "식사", "연락처", "메모"];

/* ---------- 최초 1회: 탭과 머리글 만들기 ---------- */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  [[GB_SHEET, GB_HEADER], [RSVP_SHEET, RSVP_HEADER]].forEach(function (d) {
    var sh = ss.getSheetByName(d[0]) || ss.insertSheet(d[0]);
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, d[1].length).setValues([d[1]]).setFontWeight("bold");
      sh.setFrozenRows(1);
    }
  });
  Logger.log("완료: '" + GB_SHEET + "', '" + RSVP_SHEET + "' 탭 준비됨");
}

/* ---------- 요청 처리 ---------- */
function doGet(e) {
  var p = (e && e.parameter) || {};
  var cb = String(p.callback || "");
  var result;
  try {
    if (p.action === "list") result = listItems();
    else if (p.action === "add") result = addItem(p);
    else if (p.action === "delete") result = deleteItem(p);
    else if (p.action === "rsvp") result = addRsvp(p);
    else result = { ok: false, error: "지원하지 않는 요청입니다." };
  } catch (err) {
    result = { ok: false, error: "처리 중 오류가 발생했습니다." };
  }
  return respond(result, cb);
}

/* JSONP 응답 — callback 이름은 영문,숫자,_ 만 허용 */
function respond(obj, cb) {
  var json = JSON.stringify(obj);
  if (/^[A-Za-z0-9_]{1,64}$/.test(cb)) {
    return ContentService.createTextOutput(cb + "(" + json + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function sheet_(name, header) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(header); sh.setFrozenRows(1); }
  return sh;
}

function hash_(pw) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, SALT + ":" + pw, Utilities.Charset.UTF_8);
  return bytes.map(function (b) { return ("0" + (b & 0xff).toString(16)).slice(-2); }).join("");
}

/* 스팸 제한: 같은 이름 연속 작성, 전체 분당 한도. 통과하면 true */
function rateOk_(prefix, name) {
  var cache = CacheService.getScriptCache();
  var nameKey = prefix + "n_" + Utilities.base64EncodeWebSafe(name);
  if (cache.get(nameKey)) return false;
  var minKey = prefix + "m_" + Math.floor(Date.now() / 60000);
  var count = Number(cache.get(minKey) || 0);
  if (count >= GLOBAL_PER_MIN) return false;
  cache.put(nameKey, "1", RATE_LIMIT_SEC);
  cache.put(minKey, String(count + 1), 120);
  return true;
}

/* ================= 방명록 ================= */
function listItems() {
  var sh = sheet_(GB_SHEET, GB_HEADER);
  var last = sh.getLastRow();
  if (last < 2) return { ok: true, items: [] };
  var rows = sh.getRange(2, 1, last - 1, 4).getValues();   // 비밀번호 해시(5열)는 읽지 않는다
  var tz = Session.getScriptTimeZone() || "Asia/Seoul";
  var items = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (!r[0]) continue;
    var d = r[1] ? Utilities.formatDate(new Date(r[1]), tz, "MM.dd") : "";
    items.push({ id: String(r[0]), date: d, name: String(r[2]), message: String(r[3]) });
  }
  items.reverse();
  return { ok: true, items: items };
}

function addItem(p) {
  var name = String(p.name || "").trim(), msg = String(p.message || "").trim(), pw = String(p.password || "");
  if (!name || name.length > MAX_NAME) return { ok: false, error: "이름을 확인해 주세요." };
  if (!msg || msg.length > MAX_MSG) return { ok: false, error: "내용을 확인해 주세요." };
  if (pw.length < 4 || pw.length > 20) return { ok: false, error: "비밀번호를 4자 이상 입력해 주세요." };
  if (!rateOk_("gb", name)) return { ok: false, error: "잠시 후 다시 작성해 주세요." };

  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try { sheet_(GB_SHEET, GB_HEADER).appendRow([Utilities.getUuid().slice(0, 8), new Date(), name, msg, hash_(pw)]); }
  finally { lock.releaseLock(); }
  return { ok: true };
}

function deleteItem(p) {
  var id = String(p.id || ""), pw = String(p.password || "");
  if (!id || !pw) return { ok: false, error: "비밀번호를 확인해 주세요." };
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    var sh = sheet_(GB_SHEET, GB_HEADER), last = sh.getLastRow();
    if (last < 2) return { ok: false, error: "글을 찾을 수 없습니다." };
    var ids = sh.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === id) {
        var row = i + 2;
        var isMaster = (pw === MASTER_PASSWORD);
        if (!isMaster && String(sh.getRange(row, 5).getValue()) !== hash_(pw)) return { ok: false, error: "비밀번호가 맞지 않습니다." };
        sh.deleteRow(row);
        return { ok: true };
      }
    }
    return { ok: false, error: "글을 찾을 수 없습니다." };
  } finally { lock.releaseLock(); }
}

/* ================= 참석 여부(RSVP) ================= */
function addRsvp(p) {
  var name = String(p.name || "").trim();
  var side = String(p.side || "");
  var attend = String(p.attend || "");
  var count = parseInt(p.count, 10);
  var meal = String(p.meal || "");
  var phone = String(p.phone || "").replace(/[^0-9]/g, "");
  var memo = String(p.memo || "").trim();

  if (!name || name.length > MAX_NAME) return { ok: false, error: "성함을 입력해 주세요." };
  if (side !== "신랑측" && side !== "신부측") return { ok: false, error: "신랑측,신부측을 선택해 주세요." };
  if (attend !== "참석" && attend !== "불참") return { ok: false, error: "참석 여부를 선택해 주세요." };
  if (attend === "불참") { count = 0; meal = "-"; }
  else {
    if (!(count >= 1 && count <= 10)) return { ok: false, error: "참석 인원을 선택해 주세요." };
    if (meal !== "예" && meal !== "아니오") return { ok: false, error: "식사 여부를 선택해 주세요." };
  }
  if (phone && !/^0\d{8,10}$/.test(phone)) return { ok: false, error: "연락처를 확인해 주세요." };
  if (memo.length > MAX_MEMO) return { ok: false, error: "메모가 너무 깁니다." };
  if (!rateOk_("rs", name)) return { ok: false, error: "잠시 후 다시 시도해 주세요." };

  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try { sheet_(RSVP_SHEET, RSVP_HEADER).appendRow([new Date(), name, side, attend, count, meal, phone ? "'" + phone : "", memo]); }
  finally { lock.releaseLock(); }
  return { ok: true };
}
