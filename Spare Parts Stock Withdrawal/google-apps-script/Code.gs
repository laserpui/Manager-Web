const PARTFLOW = Object.freeze({
  SHEET_NAME: "Transactions",
  SPREADSHEET_ID_PROPERTY: "SPREADSHEET_ID",
  LOCK_TIMEOUT_MS: 30000,
  MAX_PARTS: 50,
  UNITS: ["ชิ้น", "ชุด", "กล่อง"],
  STATUS_BORROWED: "borrowed",
  STATUS_RETURNED: "returned",
  HEADERS: [
    "ID",
    "Request ID",
    "Borrow Date",
    "Due Date",
    "Actual Return Date",
    "Job Number",
    "Requester Name",
    "Part Names JSON",
    "Quantity",
    "Unit",
    "Source",
    "Source Serial Number",
    "Note",
    "Status",
    "Created At",
    "Updated At",
    "Last Request ID",
    "Customer Name"
  ]
});

function doGet(event) {
  const requestId = Utilities.getUuid();
  try {
    const action = String((event && event.parameter && event.parameter.action) || "list").toLowerCase();
    if (action === "health") {
      const sheet = getTransactionsSheet_();
      return jsonResponse_({
        ok: true,
        requestId: requestId,
        data: {
          service: "PartFlow API",
          sheetName: sheet.getName(),
          spreadsheetId: sheet.getParent().getId(),
          timeZone: Session.getScriptTimeZone()
        }
      });
    }
    if (action === "list") {
      return jsonResponse_({
        ok: true,
        requestId: requestId,
        data: { records: listRecords_() }
      });
    }
    throw apiError_("UNKNOWN_ACTION", "ไม่รองรับคำขอ " + action);
  } catch (error) {
    return errorResponse_(error, requestId);
  }
}

function doPost(event) {
  const fallbackRequestId = Utilities.getUuid();
  let responseRequestId = fallbackRequestId;
  try {
    const body = parseRequestBody_(event);
    const action = String(body.action || "").toLowerCase();
    const requestId = validateRequestId_(body.requestId || fallbackRequestId);
    responseRequestId = requestId;

    if (action === "create") {
      return jsonResponse_({
        ok: true,
        requestId: requestId,
        data: withWriteLock_(function () {
          return createRecord_(body.record || {}, requestId);
        })
      });
    }

    if (action === "return") {
      return jsonResponse_({
        ok: true,
        requestId: requestId,
        data: withWriteLock_(function () {
          return returnRecord_(body.id, body.actualReturnDate, requestId);
        })
      });
    }

    throw apiError_("UNKNOWN_ACTION", "ไม่รองรับคำขอ " + (action || "(ว่าง)"));
  } catch (error) {
    return errorResponse_(error, responseRequestId);
  }
}

function setupPartFlow() {
  const sheet = getTransactionsSheet_();
  formatTransactionsSheet_(sheet);
  SpreadsheetApp.flush();
  return {
    spreadsheetId: sheet.getParent().getId(),
    sheetName: sheet.getName(),
    headers: PARTFLOW.HEADERS.slice()
  };
}

function setSpreadsheetId(spreadsheetId) {
  const value = String(spreadsheetId || "").trim();
  if (!value) throw new Error("กรุณาระบุ Spreadsheet ID");
  const spreadsheet = SpreadsheetApp.openById(value);
  PropertiesService.getScriptProperties().setProperty(PARTFLOW.SPREADSHEET_ID_PROPERTY, spreadsheet.getId());
  return setupPartFlow();
}

function clearSpreadsheetId() {
  PropertiesService.getScriptProperties().deleteProperty(PARTFLOW.SPREADSHEET_ID_PROPERTY);
}

function listRecords_() {
  const sheet = getTransactionsSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, PARTFLOW.HEADERS.length).getValues();
  return values
    .filter(function (row) { return String(row[0] || "").trim(); })
    .map(rowToRecord_)
    .sort(function (left, right) {
      return String(right.createdAt || right.borrowDate).localeCompare(String(left.createdAt || left.borrowDate));
    });
}

function createRecord_(payload, requestId) {
  const sheet = getTransactionsSheet_();
  const duplicate = findRecordByRequestId_(sheet, requestId);
  if (duplicate) {
    return { record: duplicate.record, idempotent: true };
  }

  const normalized = validateCreatePayload_(payload);
  const now = new Date().toISOString();
  const record = {
    id: Utilities.getUuid(),
    requestId: requestId,
    borrowDate: normalized.borrowDate,
    dueDate: normalized.dueDate,
    actualReturnDate: "",
    jobNumber: normalized.jobNumber,
    requesterName: normalized.requesterName,
    customerName: normalized.customerName,
    partNames: normalized.partNames,
    quantity: normalized.quantity,
    unit: normalized.unit,
    destination: normalized.destination,
    sourceSerialNumber: normalized.sourceSerialNumber,
    note: normalized.note,
    status: PARTFLOW.STATUS_BORROWED,
    createdAt: now,
    updatedAt: now,
    lastRequestId: requestId
  };

  const row = recordToRow_(record);
  sheet.getRange(sheet.getLastRow() + 1, 1, 1, row.length).setValues([row]);
  SpreadsheetApp.flush();
  return { record: record, idempotent: false };
}

function returnRecord_(id, actualReturnDate, requestId) {
  const recordId = requiredText_(id, "รหัสรายการ", 100);
  const returnedDate = validateIsoDate_(actualReturnDate, "วันที่คืนจริง");
  const sheet = getTransactionsSheet_();
  const found = findRecordById_(sheet, recordId);
  if (!found) throw apiError_("NOT_FOUND", "ไม่พบรายการที่ต้องการรับคืน");

  const record = found.record;
  if (record.lastRequestId === requestId) {
    return { record: record, idempotent: true };
  }
  if (record.status === PARTFLOW.STATUS_RETURNED) {
    throw apiError_("ALREADY_RETURNED", "รายการนี้ถูกบันทึกการคืนแล้ว");
  }
  if (returnedDate < record.borrowDate) {
    throw apiError_("INVALID_RETURN_DATE", "วันที่คืนจริงต้องเป็นวันเดียวกับหรือหลังวันที่ยืม");
  }

  record.actualReturnDate = returnedDate;
  record.status = PARTFLOW.STATUS_RETURNED;
  record.updatedAt = new Date().toISOString();
  record.lastRequestId = requestId;
  sheet.getRange(found.rowNumber, 1, 1, PARTFLOW.HEADERS.length).setValues([recordToRow_(record)]);
  SpreadsheetApp.flush();
  return { record: record, idempotent: false };
}

function validateCreatePayload_(payload) {
  const borrowDate = validateIsoDate_(payload.borrowDate, "วันที่ยืม");
  const dueDate = validateIsoDate_(payload.dueDate, "วันที่กำหนดคืน");
  if (dueDate < borrowDate) {
    throw apiError_("INVALID_DUE_DATE", "วันที่กำหนดคืนต้องเป็นวันเดียวกับหรือหลังวันที่ยืม");
  }

  const quantity = Number(payload.quantity);
  if (!Number.isInteger(quantity) || quantity < 1 || quantity > PARTFLOW.MAX_PARTS) {
    throw apiError_("INVALID_QUANTITY", "จำนวนต้องเป็นเลขจำนวนเต็มตั้งแต่ 1 ถึง " + PARTFLOW.MAX_PARTS);
  }

  if (!Array.isArray(payload.partNames)) {
    throw apiError_("INVALID_PARTS", "ข้อมูลชื่ออะไหล่ไม่ถูกต้อง");
  }
  const partNames = payload.partNames.map(function (name, index) {
    return requiredText_(name, "ชื่ออะไหล่ชิ้นที่ " + (index + 1), 250);
  });
  if (partNames.length !== quantity) {
    throw apiError_("PART_COUNT_MISMATCH", "จำนวนชื่ออะไหล่ต้องตรงกับจำนวนที่ระบุ");
  }

  const unit = requiredText_(payload.unit, "หน่วย", 30);
  if (PARTFLOW.UNITS.indexOf(unit) === -1) {
    throw apiError_("INVALID_UNIT", "หน่วยต้องเป็น ชิ้น ชุด หรือกล่อง");
  }

  return {
    borrowDate: borrowDate,
    dueDate: dueDate,
    jobNumber: requiredText_(payload.jobNumber, "เลขงาน", 120),
    requesterName: requiredText_(payload.requesterName, "ชื่อผู้เบิก", 200),
    customerName: optionalText_(payload.customerName, 250),
    partNames: partNames,
    quantity: quantity,
    unit: unit,
    destination: requiredText_(payload.destination, "นำมาจาก", 250),
    sourceSerialNumber: optionalText_(payload.sourceSerialNumber, 250),
    note: optionalText_(payload.note, 2000)
  };
}

function getTransactionsSheet_() {
  const spreadsheetId = PropertiesService.getScriptProperties().getProperty(PARTFLOW.SPREADSHEET_ID_PROPERTY);
  let spreadsheet = null;

  if (spreadsheetId) {
    spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  } else {
    spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  }

  if (!spreadsheet) {
    throw apiError_(
      "SPREADSHEET_NOT_CONFIGURED",
      "ไม่พบ Google Sheet ให้ผูกสคริปต์กับชีต หรือเรียก setSpreadsheetId('SPREADSHEET_ID') ก่อน Deploy"
    );
  }

  let sheet = spreadsheet.getSheetByName(PARTFLOW.SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(PARTFLOW.SHEET_NAME);
  ensureHeaders_(sheet);
  return sheet;
}

function ensureHeaders_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, PARTFLOW.HEADERS.length).setValues([PARTFLOW.HEADERS]);
    formatTransactionsSheet_(sheet);
    return;
  }

  const existing = sheet.getRange(1, 1, 1, PARTFLOW.HEADERS.length).getDisplayValues()[0];
  const lastHeaderIndex = PARTFLOW.HEADERS.length - 1;
  const isPreviousSchema = PARTFLOW.HEADERS.slice(0, lastHeaderIndex).every(function (header, index) {
    return existing[index] === header;
  }) && !existing[lastHeaderIndex];
  if (isPreviousSchema) {
    sheet.getRange(1, PARTFLOW.HEADERS.length).setValue(PARTFLOW.HEADERS[lastHeaderIndex]);
    formatTransactionsSheet_(sheet);
    return;
  }
  const mismatch = PARTFLOW.HEADERS.some(function (header, index) {
    return existing[index] !== header;
  });
  if (mismatch) {
    throw apiError_(
      "SCHEMA_MISMATCH",
      "หัวตารางในชีต " + PARTFLOW.SHEET_NAME + " ไม่ตรงกับโครงสร้าง PartFlow กรุณาตรวจสอบก่อนใช้งาน"
    );
  }
}

function formatTransactionsSheet_(sheet) {
  const headerRange = sheet.getRange(1, 1, 1, PARTFLOW.HEADERS.length);
  headerRange
    .setBackground("#183556")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 240);
  sheet.setColumnWidth(2, 240);
  sheet.setColumnWidths(3, 3, 130);
  sheet.setColumnWidths(6, 2, 170);
  sheet.setColumnWidth(8, 320);
  sheet.setColumnWidths(9, 2, 90);
  sheet.setColumnWidths(11, 3, 190);
  sheet.setColumnWidth(14, 110);
  sheet.setColumnWidths(15, 3, 210);
  sheet.setColumnWidth(18, 220);
  if (sheet.getMaxRows() > 1) {
    sheet.getRange(2, 3, sheet.getMaxRows() - 1, 3).setNumberFormat("@");
    sheet.getRange(2, 15, sheet.getMaxRows() - 1, 3).setNumberFormat("@");
  }
}

function findRecordById_(sheet, id) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const values = sheet.getRange(2, 1, lastRow - 1, PARTFLOW.HEADERS.length).getValues();
  for (let index = 0; index < values.length; index += 1) {
    if (String(values[index][0]) === id) {
      return { rowNumber: index + 2, record: rowToRecord_(values[index]) };
    }
  }
  return null;
}

function findRecordByRequestId_(sheet, requestId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const values = sheet.getRange(2, 1, lastRow - 1, PARTFLOW.HEADERS.length).getValues();
  for (let index = 0; index < values.length; index += 1) {
    if (String(values[index][1]) === requestId || String(values[index][16]) === requestId) {
      return { rowNumber: index + 2, record: rowToRecord_(values[index]) };
    }
  }
  return null;
}

function recordToRow_(record) {
  return [
    safeCellText_(record.id),
    safeCellText_(record.requestId),
    record.borrowDate,
    record.dueDate,
    record.actualReturnDate || "",
    safeCellText_(record.jobNumber),
    safeCellText_(record.requesterName),
    JSON.stringify(record.partNames),
    record.quantity,
    record.unit,
    safeCellText_(record.destination),
    safeCellText_(record.sourceSerialNumber),
    safeCellText_(record.note),
    record.status,
    record.createdAt,
    record.updatedAt,
    safeCellText_(record.lastRequestId),
    safeCellText_(record.customerName)
  ];
}

function rowToRecord_(row) {
  let partNames = [];
  try {
    const parsed = JSON.parse(String(row[7] || "[]"));
    if (Array.isArray(parsed)) partNames = parsed.map(String);
  } catch (error) {
    partNames = String(row[7] || "").split("|").map(function (name) { return name.trim(); }).filter(Boolean);
  }

  return {
    id: cleanCellText_(row[0]),
    requestId: cleanCellText_(row[1]),
    borrowDate: asIsoDateString_(row[2]),
    dueDate: asIsoDateString_(row[3]),
    actualReturnDate: asIsoDateString_(row[4]),
    jobNumber: cleanCellText_(row[5]),
    requesterName: cleanCellText_(row[6]),
    partNames: partNames,
    quantity: Number(row[8]) || partNames.length,
    unit: String(row[9] || ""),
    destination: cleanCellText_(row[10]),
    sourceSerialNumber: cleanCellText_(row[11]),
    note: cleanCellText_(row[12]),
    status: String(row[13] || PARTFLOW.STATUS_BORROWED),
    createdAt: String(row[14] || ""),
    updatedAt: String(row[15] || ""),
    lastRequestId: cleanCellText_(row[16]),
    customerName: cleanCellText_(row[17])
  };
}

function asIsoDateString_(value) {
  if (!value) return "";
  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return String(value).slice(0, 10);
}

function parseRequestBody_(event) {
  if (!event || !event.postData || !event.postData.contents) {
    throw apiError_("EMPTY_BODY", "ไม่พบข้อมูลที่ส่งมา");
  }
  try {
    return JSON.parse(event.postData.contents);
  } catch (error) {
    throw apiError_("INVALID_JSON", "รูปแบบข้อมูลที่ส่งมาไม่ถูกต้อง");
  }
}

function withWriteLock_(callback) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(PARTFLOW.LOCK_TIMEOUT_MS)) {
    throw apiError_("BUSY", "ระบบกำลังบันทึกข้อมูลอื่น กรุณาลองใหม่อีกครั้ง");
  }
  try {
    return callback();
  } finally {
    lock.releaseLock();
  }
}

function validateRequestId_(value) {
  const requestId = String(value || "").trim();
  if (!requestId || requestId.length > 120) {
    throw apiError_("INVALID_REQUEST_ID", "Request ID ไม่ถูกต้อง");
  }
  return requestId;
}

function validateIsoDate_(value, label) {
  const text = String(value || "").trim();
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(text);
  if (!match) throw apiError_("INVALID_DATE", label + " ไม่ถูกต้อง");

  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const date = new Date(Date.UTC(year, month - 1, day));
  if (
    date.getUTCFullYear() !== year ||
    date.getUTCMonth() !== month - 1 ||
    date.getUTCDate() !== day
  ) {
    throw apiError_("INVALID_DATE", label + " ไม่ถูกต้อง");
  }
  return text;
}

function requiredText_(value, label, maxLength) {
  const text = String(value == null ? "" : value).trim();
  if (!text) throw apiError_("REQUIRED_FIELD", "กรุณาระบุ" + label);
  if (text.length > maxLength) throw apiError_("TEXT_TOO_LONG", label + "ยาวเกิน " + maxLength + " ตัวอักษร");
  return text;
}

function optionalText_(value, maxLength) {
  const text = String(value == null ? "" : value).trim();
  if (text.length > maxLength) throw apiError_("TEXT_TOO_LONG", "ข้อมูลยาวเกิน " + maxLength + " ตัวอักษร");
  return text;
}

function safeCellText_(value) {
  const text = String(value == null ? "" : value);
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function cleanCellText_(value) {
  const text = String(value == null ? "" : value);
  return /^'[=+\-@]/.test(text) ? text.slice(1) : text;
}

function apiError_(code, message) {
  const error = new Error(message);
  error.code = code;
  return error;
}

function errorResponse_(error, requestId) {
  console.error(error && error.stack ? error.stack : error);
  return jsonResponse_({
    ok: false,
    requestId: requestId,
    error: {
      code: String((error && error.code) || "SERVER_ERROR"),
      message: String((error && error.message) || "เกิดข้อผิดพลาดที่เซิร์ฟเวอร์")
    }
  });
}

function jsonResponse_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
