/**
 * 대학 정보 오류 신고 + 대학 추가 요청 접수기 (Google Apps Script)
 *
 * - 사이트에서 보낸 신고는 "신고" 탭에, 대학 추가 요청은 "추가 요청" 탭에 한 줄씩 쌓고,
 *   관리자에게 메일로 알려줘요.
 * - 설치 방법은 '신고-접수-설정-가이드.md' 를 보세요.
 */

const ADMIN_EMAIL = "admin@example.com";   // ★ 신고 알림 메일을 받을 주소로 바꾸세요
const SHEET_NAME  = "신고";
const STATUS_LIST = ["신규", "확인 중", "반영 완료", "반려"];
const HEADERS = ["접수시각", "상태", "학교", "학과", "오류 위치", "내용", "근거 링크", "연락처", "데이터 확인일", "페이지", "처리 메모"];

const ADD_SHEET_NAME = "추가 요청";
const ADD_STATUS_LIST = ["신규", "검토 중", "추가 완료", "보류"];
const ADD_HEADERS = ["접수시각", "상태", "대학명", "학과명", "의견", "처리 메모"];

const LIMIT = { school: 100, dept: 200, section: 60, message: 1000, source: 300, contact: 120, checked: 20, page: 300, opinion: 500 };
const MAX_PER_HOUR = 100;                  // 시간당 최대 접수 건수 (도배 방지)

/** 처음 한 번만 직접 실행하세요: 두 시트("신고", "추가 요청")의 머리글과 '상태' 드롭다운을 만들어요. */
function setup() {
  prepareSheet_(getSheet_(SHEET_NAME), HEADERS, STATUS_LIST);
  prepareSheet_(getSheet_(ADD_SHEET_NAME), ADD_HEADERS, ADD_STATUS_LIST);
}
function prepareSheet_(sh, headers, statusList) {
  if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#14213d").setFontColor("#ffffff");
  }
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(statusList, true).build();
  sh.getRange(2, 2, 2000, 1).setDataValidation(rule);
}

/**
 * (이미 예전 버전으로 "추가 요청" 탭을 만들어 두셨다면 한 번만 실행하세요)
 * 더 이상 수집하지 않는 '페이지' 열을 "추가 요청" 탭에서 지워요.
 */
function migrateAddSheet() {
  const sh = getSheet_(ADD_SHEET_NAME);
  if (sh.getLastRow() === 0) return;
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const i = headers.indexOf("페이지");
  if (i >= 0) sh.deleteColumn(i + 1);
}

/** 사이트에서 오는 신고(POST)를 받아요. */
function doPost(e) {
  try {
    const raw = (e && e.postData && e.postData.contents) || "";
    const p = JSON.parse(raw);
    if (p.hp) return json_({ ok: true });                          // 스팸(숨김 칸에 값이 있으면 무시)
    if (!throttleOk_()) return json_({ ok: false, error: "busy" });
    if (p.type === "add") return handleAdd_(p);                    // 대학 추가 요청

    const message = clean_(p.message, LIMIT.message, true);
    if (message.length < 5) return json_({ ok: false, error: "short" });

    const row = [
      new Date(), "신규",
      clean_(p.school, LIMIT.school), clean_(p.dept, LIMIT.dept), clean_(p.section, LIMIT.section),
      message, clean_(p.source, LIMIT.source), clean_(p.contact, LIMIT.contact),
      clean_(p.dataCheckedAt, LIMIT.checked), clean_(p.page, LIMIT.page), ""
    ].map((v, i) => i === 0 || i === 1 ? v : safe_(v));            // 수식 삽입 방지

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try { getSheet_(SHEET_NAME).appendRow(row); } finally { lock.releaseLock(); }

    try { notify_(row); } catch (mailErr) { /* 메일 한도를 넘어도 시트에는 이미 기록돼 있어요 */ }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: "bad request" });
  }
}

/** 대학 추가 요청: 대학명·학과명·의견 */
function handleAdd_(p) {
  const univ = clean_(p.univ, LIMIT.school), dept = clean_(p.dept, LIMIT.dept);
  if (univ.length < 2 || dept.length < 2) return json_({ ok: false, error: "short" });
  const row = [new Date(), "신규", univ, dept, clean_(p.opinion, LIMIT.opinion, true), ""]
    .map((v, i) => i < 2 ? v : safe_(v));
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try { getSheet_(ADD_SHEET_NAME).appendRow(row); } finally { lock.releaseLock(); }
  try {
    MailApp.sendEmail(ADMIN_EMAIL, "[대학 추가 요청] " + univ + " · " + dept,
      ["새 대학 추가 요청이 들어왔어요.", "", "대학명: " + univ, "학과명: " + dept, "", "의견:", row[4] || "-", "",
       "시트에서 상태를 바꿔 처리 현황을 관리하세요:", SpreadsheetApp.getActiveSpreadsheet().getUrl()].join("\n"));
  } catch (mailErr) { /* 메일 한도를 넘어도 시트에는 이미 기록돼 있어요 */ }
  return json_({ ok: true });
}

/** 주소를 브라우저로 열어 보면 동작 여부를 확인할 수 있어요. */
function doGet() { return ContentService.createTextOutput("신고 접수기가 동작 중이에요."); }

function notify_(row) {
  const sheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  const body = [
    "새 정보 오류 신고가 들어왔어요.", "",
    "학교: " + row[2], "학과: " + (row[3] || "-"), "오류 위치: " + row[4], "",
    "내용:", row[5], "",
    "근거 링크: " + (row[6] || "-"), "연락처: " + (row[7] || "-"),
    "데이터 확인일: " + (row[8] || "-"), "",
    "시트에서 상태를 바꿔 처리 현황을 관리하세요:", sheetUrl
  ].join("\n");
  MailApp.sendEmail(ADMIN_EMAIL, "[정보 오류 신고] " + row[2] + (row[3] ? " · " + row[3] : ""), body);
}

function getSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}
function throttleOk_() {
  const cache = CacheService.getScriptCache();
  const key = "cnt_" + Math.floor(Date.now() / 3600000);
  const n = +(cache.get(key) || 0);
  if (n >= MAX_PER_HOUR) return false;
  cache.put(key, String(n + 1), 4000);
  return true;
}
function clean_(v, max, keepNewline) {
  let t = String(v == null ? "" : v);
  t = keepNewline ? t.replace(/[\u0000-\u0009\u000b-\u001f]/g, " ") : t.replace(/[\u0000-\u001f]/g, " ");
  return t.trim().slice(0, max);
}
function safe_(v) { return /^[=+\-@]/.test(v) ? "'" + v : v; }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
