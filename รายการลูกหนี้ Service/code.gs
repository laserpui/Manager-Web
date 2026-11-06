// ==========================================
// Google Apps Script - API สำหรับหลังบ้าน
// ==========================================
const SHEET_GID = 2130399355; // รหัส GID ของแท็บ Sheet (2130399355)

function getTargetSheet() {
  // ดึง Spreadsheet ที่สคริปต์นี้ถูกผูกไว้โดยอัตโนมัติ
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === SHEET_GID) return sheets[i];
  }
  return ss.getSheets()[0]; // หากไม่พบแท็บที่ระบุ ให้ใช้แท็บแรก
}

function findHeaderColumn_(headers, keywords) {
  for (var i = 0; i < headers.length; i++) {
    var header = (headers[i] || '').toString().trim().toLowerCase();
    for (var j = 0; j < keywords.length; j++) {
      if (header.indexOf(keywords[j]) !== -1) return i + 1;
    }
  }
  return 0;
}

function getSheetColumns_(sheet, ensureWritableColumns) {
  var lastColumn = Math.max(sheet.getLastColumn(), 1);
  var headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0];
  var timestamp = findHeaderColumn_(headers, ['ประทับเวลา', 'timestamp']);
  var base = timestamp ? 2 : 1;
  var columns = {
    timestamp: timestamp,
    date: findHeaderColumn_(headers, ['วัน เดือน ปี', 'date']) || base,
    invoice: findHeaderColumn_(headers, ['invoice']) || base + 1,
    resp1: findHeaderColumn_(headers, ['ผู้รับผิดชอบที่ 1', 'responsible 1']) || base + 2,
    resp2: findHeaderColumn_(headers, ['ผู้รับผิดชอบที่ 2', 'responsible 2']) || base + 3,
    jobNo: findHeaderColumn_(headers, ['เลขงาน', 'job']) || base + 4,
    customer: findHeaderColumn_(headers, ['ลูกค้า', 'customer']) || base + 5,
    machine: findHeaderColumn_(headers, ['ชื่อเครื่อง', 'machine']) || base + 6,
    serial: findHeaderColumn_(headers, ['ซีเรียล', 'serial']) || base + 7,
    product: findHeaderColumn_(headers, ['รายการสินค้า', 'product']) || base + 8,
    status: findHeaderColumn_(headers, ['สถานะ', 'status']),
    note: findHeaderColumn_(headers, ['หมายเหตุ', 'note', 'remark']),
    amount: findHeaderColumn_(headers, ['จำนวนเงิน', 'ยอดเงิน', 'amount'])
  };

  if (ensureWritableColumns) {
    if (!columns.status) {
      columns.status = ++lastColumn;
      sheet.getRange(1, columns.status).setValue('สถานะ');
    }
    if (!columns.note) {
      columns.note = ++lastColumn;
      sheet.getRange(1, columns.note).setValue('หมายเหตุ');
    }
    if (!columns.amount) {
      columns.amount = ++lastColumn;
      sheet.getRange(1, columns.amount).setValue('จำนวนเงิน');
    }
  }
  return columns;
}

function doGet(e) {
  var action = e.parameter.action;
  var callback = e.parameter.callback;
  var result = { status: 'error', message: 'Unknown action' };

  try {
    if (action === 'fetchCsv') {
      var csvUrl = e.parameter.url;
      var response = UrlFetchApp.fetch(csvUrl);
      result = { success: true, status: 'success', content: response.getContentText() };
    }
    else {
      var sheet = getTargetSheet();

      if (action === 'add') {
        var amountText = (e.parameter.amount || '').toString().replace(/,/g, '').trim();
        var amount = Number(amountText);
        if (amountText === '' || !isFinite(amount) || amount < 0) {
          throw new Error('จำนวนเงินไม่ถูกต้อง');
        }

        var lock = LockService.getDocumentLock();
        lock.waitLock(10000);
        try {
          var addColumns = getSheetColumns_(sheet, true);
          var addRow = sheet.getLastRow() + 1;
          var addWidth = sheet.getLastColumn();
          var addValues = [];
          for (var addIndex = 0; addIndex < addWidth; addIndex++) addValues.push('');
          if (addColumns.timestamp) addValues[addColumns.timestamp - 1] = new Date();
          addValues[addColumns.date - 1] = e.parameter.date || '';
          addValues[addColumns.invoice - 1] = e.parameter.invoice || '';
          addValues[addColumns.resp1 - 1] = e.parameter.resp1 || '';
          addValues[addColumns.resp2 - 1] = e.parameter.resp2 || '';
          addValues[addColumns.jobNo - 1] = e.parameter.jobNo || '';
          addValues[addColumns.customer - 1] = e.parameter.customer || '';
          addValues[addColumns.machine - 1] = e.parameter.machine || '';
          addValues[addColumns.serial - 1] = e.parameter.serial || '';
          addValues[addColumns.product - 1] = e.parameter.product || '';
          addValues[addColumns.status - 1] = e.parameter.status || 'ค้างชำระ';
          addValues[addColumns.note - 1] = e.parameter.note || '';
          addValues[addColumns.amount - 1] = amount;
          sheet.getRange(addRow, 1, 1, addWidth).setValues([addValues]);
          sheet.getRange(addRow, addColumns.amount).setNumberFormat('#,##0.00');
        } finally {
          lock.releaseLock();
        }
        result = { status: 'ok', message: 'บันทึกข้อมูลสำเร็จ' };
      }
      else if (action === 'edit') {
        var editRow = parseInt(e.parameter.row, 10);
        var editAmountText = (e.parameter.amount || '').toString().replace(/,/g, '').trim();
        var editAmount = Number(editAmountText);
        if (!isFinite(editRow) || editRow < 2) throw new Error('แถวข้อมูลไม่ถูกต้อง');
        if (editAmountText === '' || !isFinite(editAmount) || editAmount < 0) throw new Error('จำนวนเงินไม่ถูกต้อง');

        var editLock = LockService.getDocumentLock();
        editLock.waitLock(10000);
        try {
          if (editRow > sheet.getLastRow()) throw new Error('ไม่พบข้อมูลที่ต้องการแก้ไข');
          var editColumns = getSheetColumns_(sheet, true);
          var currentInvoice = sheet.getRange(editRow, editColumns.invoice).getDisplayValue().toString().trim();
          var currentJobNo = sheet.getRange(editRow, editColumns.jobNo).getDisplayValue().toString().trim();
          var currentSerial = sheet.getRange(editRow, editColumns.serial).getDisplayValue().toString().trim();
          if (currentInvoice !== (e.parameter.originalInvoice || '').toString().trim() ||
              currentJobNo !== (e.parameter.originalJobNo || '').toString().trim() ||
              currentSerial !== (e.parameter.originalSerial || '').toString().trim()) {
            throw new Error('ข้อมูลแถวนี้มีการเปลี่ยนแปลง กรุณารีเฟรช Dashboard แล้วลองใหม่');
          }

          sheet.getRange(editRow, editColumns.date).setValue(e.parameter.date || '');
          sheet.getRange(editRow, editColumns.invoice).setValue(e.parameter.invoice || '');
          sheet.getRange(editRow, editColumns.resp1).setValue(e.parameter.resp1 || '');
          sheet.getRange(editRow, editColumns.resp2).setValue(e.parameter.resp2 || '');
          sheet.getRange(editRow, editColumns.jobNo).setValue(e.parameter.jobNo || '');
          sheet.getRange(editRow, editColumns.customer).setValue(e.parameter.customer || '');
          sheet.getRange(editRow, editColumns.machine).setValue(e.parameter.machine || '');
          sheet.getRange(editRow, editColumns.serial).setValue(e.parameter.serial || '');
          sheet.getRange(editRow, editColumns.product).setValue(e.parameter.product || '');
          sheet.getRange(editRow, editColumns.amount).setValue(editAmount).setNumberFormat('#,##0.00');
        } finally {
          editLock.releaseLock();
        }
        result = { status: 'ok', message: 'แก้ไขแถว ' + editRow + ' สำเร็จ' };
      }
      else if (action === 'update') {
        var row = parseInt(e.parameter.row);
        var status = e.parameter.status || '';
        var note = e.parameter.note || '';
        sheet.getRange(row, 11).setValue(status);  // คอลัมน์ K
        sheet.getRange(row, 12).setValue(note);     // คอลัมน์ L
        result = { status: 'ok', message: 'อัปเดตแถว ' + row + ' สำเร็จ' };
      }
      else if (action === 'delete') {
        var row = parseInt(e.parameter.row);
        if (row < 2) {
          result = { status: 'error', message: 'ไม่สามารถลบแถวหัวตารางได้' };
        } else {
          sheet.deleteRow(row);
          result = { status: 'ok', message: 'ลบแถว ' + row + ' สำเร็จ' };
        }
      }
    }
  } catch (err) {
    result = { status: 'error', message: err.toString() };
  }

  var output = JSON.stringify(result);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + output + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(output)
    .setMimeType(ContentService.MimeType.JSON);
}
