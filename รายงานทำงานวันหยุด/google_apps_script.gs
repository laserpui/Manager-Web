/**
 * Google Apps Script for Holiday Work Report System
 * Paste this script into your Google Sheet's Extensions > Apps Script editor.
 * Deploy as a Web App: 
 * - Execute as: "Me"
 * - Who has access: "Anyone"
 */

// Configure Sheet name (uses the first sheet if empty)
const SHEET_NAME = "การตอบแบบฟอร์ม 1"; 

// Data Split Configuration
// "employee" -> Split into sheets named by employee name
// "month" -> Split into sheets named by month (e.g. "มิถุนายน 2569")
// "none" -> Save in the main sheet only
const SPLIT_BY = "employee";

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (SHEET_NAME) {
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(["วันทำงานวันหยุด", "รายชื่อพนักงาน", "รายละเอียดงาน", "เวลาบันทึก"]);
      // Format headers
      sheet.getRange(1, 1, 1, 4).setFontWeight("bold").setBackground("#f3f4f6");
    }
    return sheet;
  }
  
  // Default to first sheet
  const sheets = ss.getSheets();
  if (sheets.length === 0) {
    const sheet = ss.insertSheet("บันทึกรายงาน");
    sheet.appendRow(["วันทำงานวันหยุด", "รายชื่อพนักงาน", "รายละเอียดงาน", "เวลาบันทึก"]);
    sheet.getRange(1, 1, 1, 4).setFontWeight("bold").setBackground("#f3f4f6");
    return sheet;
  }
  
  // If first sheet exists but is empty, setup headers
  const sheet = sheets[0];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["วันทำงานวันหยุด", "รายชื่อพนักงาน", "รายละเอียดงาน", "เวลาบันทึก"]);
    sheet.getRange(1, 1, 1, 4).setFontWeight("bold").setBackground("#f3f4f6");
  }
  return sheet;
}

/**
 * Handle GET requests: Fetch all records from the Google Sheet
 */
function doGet(e) {
  try {
    // Check if proxy CSV request
    if (e && e.parameter && e.parameter.action === 'fetchCsv') {
      const csvUrl = e.parameter.url;
      const response = UrlFetchApp.fetch(csvUrl);
      return jsonResponse({ success: true, status: "success", content: response.getContentText() }, e);
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Get unique employees
    const uniqueEmployees = new Set();
    
    // Fallback/Existing hardcoded list from index.html
    const defaultList = ["กษิเดช", "กันต์ศักดิ์", "จีรศักดิ์", "ชาติชาย", "ณัฐชัย", "ณัฐพงษ์", "ธนบูรณ์", "ประวิทย์", "พงษ์ศักดิ์", "ภาสพงษ์", "สมชาย", "อภิลักษณ์", "เอกเลิศ"];
    defaultList.forEach(name => uniqueEmployees.add(name));
    
    // Read employees from sheet tabs if SPLIT_BY is "employee"
    if (SPLIT_BY === "employee") {
      const sheets = ss.getSheets();
      sheets.forEach(s => {
        const name = s.getName().trim();
        if (name && name !== SHEET_NAME && name !== "บันทึกรายงาน" && !name.startsWith("รายเดือน_") && !name.includes("การตอบแบบฟอร์ม")) {
          uniqueEmployees.add(name);
        }
      });
    }

    const sheet = getSheet();
    const lastRow = sheet.getLastRow();
    
    // Read employees from the main sheet data column to capture any others
    if (lastRow > 1) {
      try {
        const lastCol = sheet.getLastColumn();
        const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
        const employeeIdx = headers.findIndex(h => h.includes("รายชื่อพนักงาน") || h.toLowerCase().includes("employee") || h.toLowerCase().includes("name"));
        if (employeeIdx !== -1) {
          const employeeValues = sheet.getRange(2, employeeIdx + 1, lastRow - 1, 1).getValues();
          employeeValues.forEach(row => {
            const name = String(row[0]).trim();
            if (name) {
              uniqueEmployees.add(name);
            }
          });
        }
      } catch (err) {
        console.error("Error reading employee names from data: " + err.toString());
      }
    }
    
    const employees = Array.from(uniqueEmployees);
    employees.sort();
    
    // If only header exists or empty
    if (lastRow <= 1) {
      return jsonResponse({ status: "success", data: [], employees: employees }, e);
    }
    
    // Get headers to determine column indices dynamically
    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
    
    // Find matching columns
    const dateIdx = headers.findIndex(h => h.includes("วันทำงานวันหยุด") || h.includes("วัน เดือน ปี") || h.toLowerCase().includes("date"));
    const employeeIdx = headers.findIndex(h => h.includes("รายชื่อพนักงาน") || h.toLowerCase().includes("employee") || h.toLowerCase().includes("name"));
    const detailsIdx = headers.findIndex(h => h.includes("รายละเอียดงาน") || h.toLowerCase().includes("detail") || h.toLowerCase().includes("job"));
    const timestampIdx = headers.findIndex(h => h.includes("เวลาบันทึก") || h.toLowerCase().includes("timestamp") || h.includes("ประทับเวลา"));
    
    const dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
    const values = dataRange.getValues();
    
    const records = values.map((row, idx) => {
      // Date field parsing
      let dateVal = dateIdx !== -1 ? row[dateIdx] : "";
      let formattedDate = "";
      if (dateVal instanceof Date) {
        formattedDate = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "yyyy-MM-dd");
      } else {
        formattedDate = String(dateVal);
      }
      
      // Timestamp field parsing
      let timestampVal = timestampIdx !== -1 ? row[timestampIdx] : "";
      let formattedTimestamp = "";
      if (timestampVal instanceof Date) {
        formattedTimestamp = Utilities.formatDate(timestampVal, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
      } else {
        formattedTimestamp = String(timestampVal);
      }
      
      return {
        date: formattedDate,
        employee: employeeIdx !== -1 ? String(row[employeeIdx]) : "",
        details: detailsIdx !== -1 ? String(row[detailsIdx]) : "",
        timestamp: formattedTimestamp,
        rowNumber: idx + 2 // Store physical row index (starts at row 2)
      };
    });
    
    // Sort records descending by date
    records.sort((a, b) => {
      let timeA = 0;
      let timeB = 0;
      
      if (a.date) {
        const parsedA = new Date(a.date.split(' ')[0] || a.date);
        timeA = isNaN(parsedA.getTime()) ? 0 : parsedA.getTime();
      }
      if (b.date) {
        const parsedB = new Date(b.date.split(' ')[0] || b.date);
        timeB = isNaN(parsedB.getTime()) ? 0 : parsedB.getTime();
      }
      
      // Sort by date, and if dates are equal, sort by timestamp
      if (timeB !== timeA) {
        return timeB - timeA;
      }
      
      let tsA = a.timestamp ? new Date(a.timestamp).getTime() : 0;
      let tsB = b.timestamp ? new Date(b.timestamp).getTime() : 0;
      return (isNaN(tsB) ? 0 : tsB) - (isNaN(tsA) ? 0 : tsA);
    });
    
    return jsonResponse({ status: "success", data: records, employees: employees }, e);
  } catch (error) {
    return jsonResponse({ status: "error", message: error.toString() }, e);
  }
}

/**
 * Handle POST requests: Save a new record or delete a record
 */
function doPost(e) {
  try {
    // Parse request content
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({ status: "error", message: "ไม่พบข้อมูลส่งมา (No post data)" });
    }
    
    const payload = JSON.parse(e.postData.contents);
    
    // ----------------------------------------------------
    // CASE 1: Delete Record
    // ----------------------------------------------------
    if (payload.action === 'delete') {
      const date = payload.date;
      const employee = payload.employee;
      const details = payload.details;
      const rowNumber = payload.rowNumber;
      
      if (!date || !employee || !details) {
        return jsonResponse({ status: "error", message: "ระบุเงื่อนไขในการลบข้อมูลไม่ครบถ้วน" });
      }
      
      const sheet = getSheet();
      const lastRow = sheet.getLastRow();
      
      if (lastRow <= 1) {
        return jsonResponse({ status: "error", message: "ไม่พบรายการข้อมูลในฐานข้อมูล" });
      }
      
      const lastCol = sheet.getLastColumn();
      const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
      
      const dateIdx = headers.findIndex(h => h.includes("วันทำงานวันหยุด") || h.includes("วัน เดือน ปี") || h.toLowerCase().includes("date"));
      const employeeIdx = headers.findIndex(h => h.includes("รายชื่อพนักงาน") || h.toLowerCase().includes("employee") || h.toLowerCase().includes("name"));
      const detailsIdx = headers.findIndex(h => h.includes("รายละเอียดงาน") || h.toLowerCase().includes("detail") || h.toLowerCase().includes("job"));
      
      let deletedFromMain = false;
      
      // 1. Direct row deletion (Optimized path using sent rowNumber)
      if (rowNumber && rowNumber > 1 && rowNumber <= lastRow) {
        const rowValues = sheet.getRange(rowNumber, 1, 1, lastCol).getValues()[0];
        let rowDate = rowValues[dateIdx] instanceof Date ? Utilities.formatDate(rowValues[dateIdx], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(rowValues[dateIdx]);
        let rowEmployee = String(rowValues[employeeIdx]);
        let rowDetails = String(rowValues[detailsIdx]);
        
        if (rowDate === date && rowEmployee === employee && rowDetails === details) {
          sheet.deleteRow(rowNumber);
          deletedFromMain = true;
        }
      }
      
      // 2. Search fallback (in case row shifted due to concurrent deletes/inserts)
      if (!deletedFromMain) {
        const dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
        const values = dataRange.getValues();
        for (let i = 0; i < values.length; i++) {
          const rowValues = values[i];
          let rowDate = rowValues[dateIdx] instanceof Date ? Utilities.formatDate(rowValues[dateIdx], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(rowValues[dateIdx]);
          let rowEmployee = String(rowValues[employeeIdx]);
          let rowDetails = String(rowValues[detailsIdx]);
          
          if (rowDate === date && rowEmployee === employee && rowDetails === details) {
            sheet.deleteRow(i + 2); // i is 0-based index of values which starts at row 2
            deletedFromMain = true;
            break;
          }
        }
      }
      
      if (!deletedFromMain) {
        return jsonResponse({ status: "error", message: "ไม่พบข้อมูลที่ระบุในฐานข้อมูลหลัก" });
      }
      
      // 3. Delete from separate sheets if enabled
      if (SPLIT_BY === "employee" && employee) {
        deleteFromSeparateSheet(employee, date, employee, details);
      } else if (SPLIT_BY === "month" && date) {
        const monthName = getMonthSheetName(date);
        deleteFromSeparateSheet(monthName, date, employee, details);
      }
      
      return jsonResponse({ status: "success", message: "ลบข้อมูลสำเร็จเรียบร้อยแล้ว!" });
    }
    
    // ----------------------------------------------------
    // CASE 2: Save New Record
    // ----------------------------------------------------
    const date = payload.date;
    const employee = payload.employee;
    const details = payload.details;
    
    if (!date || !employee || !details) {
      return jsonResponse({ status: "error", message: "กรอกข้อมูลไม่ครบถ้วน (Missing required fields)" });
    }
    
    const sheet = getSheet();
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
    
    // Find matching columns
    const dateIdx = headers.findIndex(h => h.includes("วันทำงานวันหยุด") || h.includes("วัน เดือน ปี") || h.toLowerCase().includes("date"));
    const employeeIdx = headers.findIndex(h => h.includes("รายชื่อพนักงาน") || h.toLowerCase().includes("employee") || h.toLowerCase().includes("name"));
    const detailsIdx = headers.findIndex(h => h.includes("รายละเอียดงาน") || h.toLowerCase().includes("detail") || h.toLowerCase().includes("job"));
    const timestampIdx = headers.findIndex(h => h.includes("เวลาบันทึก") || h.toLowerCase().includes("timestamp") || h.includes("ประทับเวลา"));
    
    // 2.1 Duplicate check: Check if exact same record exists
    if (lastRow > 1) {
      const dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
      const values = dataRange.getValues();
      for (let i = 0; i < values.length; i++) {
        const row = values[i];
        let rowDate = row[dateIdx] instanceof Date ? Utilities.formatDate(row[dateIdx], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(row[dateIdx]);
        let rowEmployee = String(row[employeeIdx]);
        let rowDetails = String(row[detailsIdx]);
        
        if (rowDate === date && rowEmployee === employee && rowDetails === details) {
          return jsonResponse({ status: "error", message: "พบรายงานนี้ในระบบอยู่แล้ว (ป้องกันข้อมูลซ้ำซ้อน)" });
        }
      }
    }
    
    const timestamp = new Date();
    
    // Append to main sheet
    if (dateIdx !== -1 || employeeIdx !== -1 || detailsIdx !== -1) {
      const newRow = new Array(headers.length).fill("");
      if (dateIdx !== -1) newRow[dateIdx] = date;
      if (employeeIdx !== -1) newRow[employeeIdx] = employee;
      if (detailsIdx !== -1) newRow[detailsIdx] = details;
      if (timestampIdx !== -1) newRow[timestampIdx] = timestamp;
      sheet.appendRow(newRow);
    } else {
      // Default fallback
      sheet.appendRow([date, employee, details, timestamp]);
    }
    
    // 2.2 Split sheets writing
    if (SPLIT_BY === "employee" && employee) {
      appendToSeparateSheet(employee, date, employee, details, timestamp);
    } else if (SPLIT_BY === "month" && date) {
      const monthName = getMonthSheetName(date);
      appendToSeparateSheet(monthName, date, employee, details, timestamp);
    }
    
    return jsonResponse({ status: "success", message: "บันทึกข้อมูลเรียบร้อยแล้ว!" });
  } catch (error) {
    return jsonResponse({ status: "error", message: error.toString() });
  }
}

/**
 * Helper: Append record to a separate sheet, creating it with headers if missing
 */
function appendToSeparateSheet(sheetName, date, employee, details, timestamp) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  const headers = ["วันทำงานวันหยุด", "รายชื่อพนักงาน", "รายละเอียดงาน", "เวลาบันทึก"];
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(headers);
    // Style subsheet headers with soft blue theme
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#e0f2fe");
  }
  
  const lastCol = sheet.getLastColumn();
  const currentHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  
  const dateIdx = currentHeaders.findIndex(h => h.includes("วันทำงานวันหยุด") || h.includes("วัน เดือน ปี") || h.toLowerCase().includes("date"));
  const employeeIdx = currentHeaders.findIndex(h => h.includes("รายชื่อพนักงาน") || h.toLowerCase().includes("employee") || h.toLowerCase().includes("name"));
  const detailsIdx = currentHeaders.findIndex(h => h.includes("รายละเอียดงาน") || h.toLowerCase().includes("detail") || h.toLowerCase().includes("job"));
  const timestampIdx = currentHeaders.findIndex(h => h.includes("เวลาบันทึก") || h.toLowerCase().includes("timestamp") || h.includes("ประทับเวลา"));
  
  const newRow = new Array(currentHeaders.length).fill("");
  if (dateIdx !== -1) newRow[dateIdx] = date;
  if (employeeIdx !== -1) newRow[employeeIdx] = employee;
  if (detailsIdx !== -1) newRow[detailsIdx] = details;
  if (timestampIdx !== -1) newRow[timestampIdx] = timestamp;
  
  sheet.appendRow(newRow);
}

/**
 * Helper: Delete record from a separate sheet
 */
function deleteFromSeparateSheet(sheetName, date, employee, details) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return;
  
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;
  
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  
  const dateIdx = headers.findIndex(h => h.includes("วันทำงานวันหยุด") || h.includes("วัน เดือน ปี") || h.toLowerCase().includes("date"));
  const employeeIdx = headers.findIndex(h => h.includes("รายชื่อพนักงาน") || h.toLowerCase().includes("employee") || h.toLowerCase().includes("name"));
  const detailsIdx = headers.findIndex(h => h.includes("รายละเอียดงาน") || h.toLowerCase().includes("detail") || h.toLowerCase().includes("job"));
  
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  for (let i = 0; i < data.length; i++) {
    const rowValues = data[i];
    let rowDate = rowValues[dateIdx] instanceof Date ? Utilities.formatDate(rowValues[dateIdx], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(rowValues[dateIdx]);
    let rowEmployee = employeeIdx !== -1 ? String(rowValues[employeeIdx]) : "";
    let rowDetails = detailsIdx !== -1 ? String(rowValues[detailsIdx]) : "";
    
    const matchEmployee = employeeIdx === -1 || rowEmployee === employee;
    if (rowDate === date && matchEmployee && rowDetails === details) {
      sheet.deleteRow(i + 2); // delete first matching row in separate sheet
      break;
    }
  }
}

/**
 * Helper: Generate month sheet name in Thai format (e.g. "มิถุนายน 2569")
 */
function getMonthSheetName(dateString) {
  const parts = dateString.split("-");
  if (parts.length === 3) {
    const year = parts[0];
    const month = parts[1];
    const monthsThai = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"];
    const monthIndex = parseInt(month, 10) - 1;
    if (monthIndex >= 0 && monthIndex < 12) {
      const yearThai = parseInt(year, 10) + 543;
      return monthsThai[monthIndex] + " " + yearThai;
    }
  }
  return "รายเดือน_" + dateString.substring(0, 7);
}

/**
 * Helper to build JSON Response with CORS headers enabled
 */
function jsonResponse(data, e) {
  const JSONString = JSON.stringify(data);
  if (e && e.parameter && e.parameter.callback) {
    return ContentService.createTextOutput(e.parameter.callback + '(' + JSONString + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(JSONString)
    .setMimeType(ContentService.MimeType.JSON);
}
