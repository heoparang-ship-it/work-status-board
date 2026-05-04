const SHEET_NAME = "근태기록";
const HEADERS = [
  "수신시각",
  "이벤트",
  "사용자",
  "사용자ID",
  "근태일",
  "출근시각",
  "퇴근시각",
  "외근횟수",
  "세션수",
  "현재상태",
  "메시지",
  "원본"
];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const payload = parsePayload_(e);
    const records = Array.isArray(payload.records)
      ? payload.records
      : [payload.record || payload];
    const rows = records
      .filter(Boolean)
      .map((record) => toSheetRow_(payload, record));

    if (rows.length) {
      const sheet = getAttendanceSheet_();
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, HEADERS.length).setValues(rows);
    }

    return json_({ ok: true, appended: rows.length });
  } catch (error) {
    return json_({ ok: false, error: String(error) });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  const callback = e && e.parameter && e.parameter.callback;
  const result = { ok: true, sheet: SHEET_NAME, time: new Date().toISOString() };

  if (callback) {
    return ContentService
      .createTextOutput(`${callback}(${JSON.stringify(result)})`)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return json_(result);
}

function parsePayload_(e) {
  const raw = e && e.postData && e.postData.contents;
  if (!raw) return {};
  return JSON.parse(raw);
}

function getAttendanceSheet_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new Error("Google Sheet에서 확장 프로그램 > Apps Script로 스크립트를 열어 주세요.");
  }

  const sheet = spreadsheet.getSheetByName(SHEET_NAME) || spreadsheet.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function toSheetRow_(payload, record) {
  const receivedAt = new Date();
  return [
    receivedAt,
    record.eventType || payload.type || "",
    record.username || "",
    record.userId || "",
    record.date || "",
    record.clockInAt || "",
    record.clockOutAt || "",
    Number(record.outsideCount || 0),
    Number(record.sessionCount || 0),
    record.status || "",
    record.message || "",
    payload.source || ""
  ];
}

function json_(value) {
  return ContentService
    .createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}
