const SHEET_NAME = "근태기록";
const BOARD_STATE_KEY = "officeBoardState";
const HEADERS = [
  "날짜",
  "사용자",
  "출근 시각",
  "퇴근 시각",
  "외근 시각"
];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const payload = parsePayload_(e);
    if (payload.type === "board_state") {
      saveBoardState_(payload.boardState);
      return json_({ ok: true, saved: true });
    }

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
  const action = e && e.parameter && e.parameter.action;
  const result = action === "state"
    ? { ok: true, boardState: getBoardState_(), time: new Date().toISOString() }
    : { ok: true, sheet: SHEET_NAME, time: new Date().toISOString() };

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

function saveBoardState_(value) {
  if (!value || typeof value !== "object") return;
  PropertiesService.getScriptProperties().setProperty(BOARD_STATE_KEY, JSON.stringify(value));
}

function getBoardState_() {
  const raw = PropertiesService.getScriptProperties().getProperty(BOARD_STATE_KEY);
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch (error) {
    return null;
  }
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
  } else {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    const extraColumnCount = sheet.getLastColumn() - HEADERS.length;
    if (extraColumnCount > 0) {
      sheet.deleteColumns(HEADERS.length + 1, extraColumnCount);
    }
  }
  return sheet;
}

function toSheetRow_(payload, record) {
  return [
    formatDate_(record.date),
    record.username || "",
    formatTime_(record.clockInAt),
    formatTime_(record.clockOutAt),
    formatTimeList_(record.outsideAt)
  ];
}

function formatDate_(value) {
  if (!value) return "";
  const date = new Date(`${value}T00:00:00+09:00`);
  if (Number.isNaN(date.getTime())) return value;
  return Utilities.formatDate(date, "Asia/Seoul", "yyyy-MM-dd");
}

function formatTime_(value) {
  if (!value) return "";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return value;
  return Utilities.formatDate(date, "Asia/Seoul", "HH:mm");
}

function formatTimeList_(value) {
  if (!value) return "";
  return String(value)
    .split(",")
    .map((item) => formatTime_(item.trim()))
    .filter(Boolean)
    .join(", ");
}

function json_(value) {
  return ContentService
    .createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}
