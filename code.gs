// =========================================================================
// System: FineStamp - ระบบบันทึกเวลาปฏิบัติงาน
// Version: 1.3.0 (AI Core)
// Developer: I3AYY & AI Assistant
// Description: Backend Google Apps Script (Server-side logic)
// =========================================================================

// --- Configuration ---
const SHEET_USERS = "Users";
const SHEET_RECORDS = "TimeRecords";
const SHEET_DUTY = "DutyRoster";
const SHEET_ADV_SCHED = "AdvanceSchedule"; // Sheet สำหรับตารางล่วงหน้า Admin

// การตั้งค่า Telegram Notify
const TELEGRAM_TOKEN = "XXX"; 
const TELEGRAM_CHAT_ID = "XXX";

// --- Include HTML Files Function (แยกไฟล์) ---
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function doGet() {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('FineStamp - ระบบบันทึกเวลาปฏิบัติงาน v1.3.0 AI Core')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1.0, user-scalable=no');
}

// --- Helper Functions ---
function _formatDateForHtml(val, timezone) {
  if (!val) return "";
  try {
    if (val instanceof Date) return Utilities.formatDate(val, timezone, "yyyy-MM-dd");
    const d = new Date(val);
    if (!isNaN(d.getTime())) return Utilities.formatDate(d, timezone, "yyyy-MM-dd");
  } catch(e) {}
  return String(val);
}

function _formatTimeForHtml(val, timezone) {
  if (!val) return "";
  try {
    if (val instanceof Date) return Utilities.formatDate(val, timezone, "HH:mm");
    let str = String(val).trim();
    const parts = str.split(':');
    if (parts.length >= 2) {
      const h = parts[0].length === 1 ? '0' + parts[0] : parts[0];
      const m = parts[1].length === 1 ? '0' + parts[1] : parts[1];
      return `${h}:${m.substring(0, 2)}`;
    }
  } catch(e) {}
  return String(val);
}

function _getThaiDateString(dateObj) {
  const thaiMonths = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"];
  const day = dateObj.getDate();
  const month = thaiMonths[dateObj.getMonth()];
  const year = dateObj.getFullYear() + 543;
  return `${day} ${month} ${year}`;
}

// --- User & Auth ---
function loginUser(userId, password) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_USERS);
  if (!sheet) return JSON.stringify({ isOk: false, message: "System Error: Sheet not found." });

  const data = sheet.getDataRange().getDisplayValues();
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[1] == userId && row[5] == password) {
      return JSON.stringify({
        isOk: true,
        user: {
          id: row[0],
          user_id: row[1],
          first_name: row[2],
          last_name: row[3],
          work_groups: row[4],
          role: row[8] ? row[8].trim() : "User",
          profile_image: row[9] || ""
        }
      });
    }
  }
  return JSON.stringify({ isOk: false, message: "ID หรือรหัสผ่านไม่ถูกต้อง" });
}

function registerUser(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_USERS);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_USERS);
      sheet.appendRow(["id", "user_id", "first_name", "last_name", "work_groups", "password", "created_date", "created_time", "role", "profile_image", "email"]);
    }
    const data = JSON.parse(payload);
    const users = sheet.getDataRange().getDisplayValues();
    if (users.some(row => row[1] == data.user_id)) return JSON.stringify({ isOk: false, message: "User ID นี้มีอยู่ในระบบแล้ว" });

    const now = new Date();
    sheet.appendRow([
      data.id, data.user_id, data.first_name, data.last_name, data.work_groups, data.password,
      Utilities.formatDate(now, "GMT+7", "yyyy-MM-dd"),
      Utilities.formatDate(now, "GMT+7", "HH:mm:ss"),
      "User", "", ""
    ]);
    return JSON.stringify({ isOk: true });
  } catch (e) {
    return JSON.stringify({ isOk: false, message: e.toString() });
  } finally {
    lock.releaseLock();
  }
}

// --- Profile Image Handling ---
function saveProfileImage(userId, base64Data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_USERS);
    const data = sheet.getDataRange().getDisplayValues();
    
    let rowIndex = -1;
    let oldFileUrl = "";
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] == userId) {
        rowIndex = i + 1; 
        oldFileUrl = data[i][9]; 
        break;
      }
    }

    if (rowIndex === -1) return JSON.stringify({ isOk: false, message: "User not found" });

    if (oldFileUrl && oldFileUrl.includes("drive.google.com")) {
      try {
        const idMatch = oldFileUrl.match(/id=([^&]+)/);
        if (idMatch && idMatch[1]) {
          DriveApp.getFileById(idMatch[1]).setTrashed(true); 
        }
      } catch (err) {
        console.log("Delete old file error: " + err);
      }
    }

    const folderName = "FineStamp_Profiles";
    const folders = DriveApp.getFoldersByName(folderName);
    let folder;
    if (folders.hasNext()) folder = folders.next();
    else folder = DriveApp.createFolder(folderName);

    const contentType = base64Data.substring(5, base64Data.indexOf(';'));
    const bytes = Utilities.base64Decode(base64Data.substr(base64Data.indexOf('base64,')+7));
    const blob = Utilities.newBlob(bytes, contentType, `profile_${userId}_${Date.now()}.jpg`);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    const fileUrl = `https://drive.google.com/thumbnail?id=${file.getId()}&sz=s400`;
    sheet.getRange(rowIndex, 10).setValue(fileUrl); 

    return JSON.stringify({ isOk: true, url: fileUrl });

  } catch(e) {
    return JSON.stringify({ isOk: false, message: e.toString() });
  }
}

// --- Records Management ---
function getUserRecords(userId, month, year) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_RECORDS);
  if (!sheet) return JSON.stringify([]);
  
  const data = sheet.getDataRange().getValues();
  const tz = Session.getScriptTimeZone();
  const records = [];
  
  const targetMonth = parseInt(month, 10);
  const targetYear = parseInt(year, 10);

  for (let i = data.length - 1; i >= 1; i--) {
    const row = data[i];
    if (String(row[1]) === String(userId)) {
      let dObj = null;
      if (row[5] instanceof Date) { dObj = row[5]; }
      else if (typeof row[5] === 'string' && row[5].trim() !== "") { dObj = new Date(row[5]); }
      
      if (dObj && !isNaN(dObj.getTime())) {
         if ((dObj.getMonth() + 1) === targetMonth && dObj.getFullYear() === targetYear) {
             records.push({
                id: row[0],
                clock_in_date: _formatDateForHtml(row[5], tz),
                clock_in_time: _formatTimeForHtml(row[6], tz),
                clock_out_date: _formatDateForHtml(row[7], tz),
                clock_out_time: _formatTimeForHtml(row[8], tz),
                work_type: row[9]
             });
         }
      }
    }
  }
  return JSON.stringify(records);
}

function saveTimeRecord(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_RECORDS);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_RECORDS);
      sheet.appendRow(["id", "user_id", "first_name", "last_name", "work_groups", "clock_in_date", "clock_in_time", "clock_out_date", "clock_out_time", "work_type", "log_date", "log_time"]);
      sheet.setFrozenRows(1);
    }

    const rec = JSON.parse(payload);
    const allData = sheet.getDataRange().getValues();
    const tz = Session.getScriptTimeZone();
    let rowIndex = -1;
    let existingRow = null;

    for (let i = 1; i < allData.length; i++) {
      if (String(allData[i][0]) === String(rec.id)) {
        rowIndex = i + 1;
        existingRow = allData[i];
        break;
      }
    }

    let finalClockInDate = rec.clock_in_date;
    let finalClockInTime = rec.clock_in_time;
    let finalWorkType = rec.work_type;

    if (rowIndex !== -1 && existingRow && !rec.is_manual_edit) {
        if (!finalClockInDate) finalClockInDate = _formatDateForHtml(existingRow[5], tz);
        if (!finalClockInTime) finalClockInTime = _formatTimeForHtml(existingRow[6], tz);
        if (!finalWorkType) finalWorkType = existingRow[9];
    }

    const now = new Date();
    const rowData = [
      rec.id, rec.user_id, rec.first_name, rec.last_name, rec.work_groups,
      finalClockInDate, finalClockInTime,
      rec.clock_out_date || "", rec.clock_out_time || "",
      finalWorkType || "",
      Utilities.formatDate(now, "GMT+7", "yyyy-MM-dd"),
      Utilities.formatDate(now, "GMT+7", "HH:mm:ss")
    ];
    const stringRowData = rowData.map(d => String(d));

    if (rowIndex !== -1) {
      sheet.getRange(rowIndex, 1, 1, stringRowData.length).setValues([stringRowData]);
    } else {
      sheet.appendRow(stringRowData);
    }

    // --- TELEGRAM NOTIFY LOGIC (ลงเวลา) ---
    let icon = "🔵"; 
    let title = "ลงเวลาเข้างาน";
    let nameDisplay = `${rec.first_name} ${rec.last_name}`;
    let workTypeDisplay = finalWorkType || "-";
    let timeInDisplay = finalClockInTime ? `${finalClockInTime} น.` : "-";
    let timeOutDisplay = "";

    if (rec.is_manual_edit) {
        icon = "✏️";
        title = "แก้ไขเวลาปฏิบัติงาน";
        timeOutDisplay = rec.clock_out_time ? `${rec.clock_out_time} น.` : "-";
    }
    else if (rec.work_type === "เจาะเลือดเช้า") {
        icon = "🟢"; 
        title = "ลงเวลาเข้า-ออกงาน";
        timeOutDisplay = rec.clock_out_time ? `${rec.clock_out_time} น.` : "08:00 น.";
    } 
    else if (rec.clock_out_time && rec.clock_out_time.trim() !== "") {
        icon = "🔴"; 
        title = "ลงเวลาออกงาน";
        timeOutDisplay = `${rec.clock_out_time} น.`;
    } 
    else {
        icon = "🔵"; 
        title = "ลงเวลาเข้างาน";
        timeOutDisplay = "<i>..............</i>"; 
    }

    const msg = `<b>${icon} ${title}</b>\n` +
                `➖➖➖➖➖➖➖➖➖➖\n` +
                `👤 <b>ชื่อ-สกุล:</b>  ${nameDisplay}\n` +
                `💼 <b>งาน:</b>            ${workTypeDisplay}\n` +
                `🕐 <b>เวลามา:</b>      ${timeInDisplay}\n` +
                `🏁 <b>เวลากลับ:</b>   ${timeOutDisplay}`;
    
    sendTelegramMsg(msg);
    // ------------------------------------

    return JSON.stringify({ isOk: true, record: rec });
  } catch (e) {
    return JSON.stringify({ isOk: false, message: e.toString() });
  } finally {
    lock.releaseLock();
  }
}

function deleteTimeRecord(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const data = JSON.parse(payload);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_RECORDS);
    
    if (!sheet) return JSON.stringify({ isOk: false, message: "Sheet not found." });

    const allData = sheet.getDataRange().getValues();
    let rowIndex = -1;

    for (let i = 1; i < allData.length; i++) {
      if (String(allData[i][0]) === String(data.record_id) && String(allData[i][1]) === String(data.user_id)) {
        rowIndex = i + 1; 
        break;
      }
    }

    if (rowIndex !== -1) {
      sheet.deleteRow(rowIndex);
      return JSON.stringify({ isOk: true });
    } else {
      return JSON.stringify({ isOk: false, message: "ไม่พบข้อมูล หรือคุณไม่มีสิทธิ์ลบรายการนี้" });
    }

  } catch (e) {
    return JSON.stringify({ isOk: false, message: e.toString() });
  } finally {
    lock.releaseLock();
  }
}

// --- DUTY ROSTER FUNCTIONS ---
function getMonthDuty(userId, month, year) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_DUTY);
  if (!sheet) return JSON.stringify([]);

  const data = sheet.getDataRange().getValues();
  const results = [];
  
  const targetM = parseInt(month, 10);
  const targetY = parseInt(year, 10);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (String(row[1]) !== String(userId)) continue; 

    let dObj = null;
    if (row[2] instanceof Date) dObj = row[2];
    else if (typeof row[2] === 'string') dObj = new Date(row[2]);

    if (dObj && !isNaN(dObj.getTime())) {
      if ((dObj.getMonth() + 1) === targetM && dObj.getFullYear() === targetY) {
        results.push({
          date: Utilities.formatDate(dObj, "GMT+7", "yyyy-MM-dd"),
          shifts: row[3] ? row[3].split(',') : []
        });
      }
    }
  }
  return JSON.stringify(results);
}

function saveDutyRecord(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_DUTY);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_DUTY);
      sheet.appendRow(["id", "user_id", "shift_date", "shift_type", "created_at"]);
    }

    const data = JSON.parse(payload); 
    const allData = sheet.getDataRange().getValues();
    const targetDateStr = data.date;
    
    let rowIndex = -1;
    for (let i = 1; i < allData.length; i++) {
      let dStr = "";
      if (allData[i][2] instanceof Date) dStr = Utilities.formatDate(allData[i][2], "GMT+7", "yyyy-MM-dd");
      else dStr = String(allData[i][2]);

      if (String(allData[i][1]) === String(data.user_id) && dStr === targetDateStr) {
        rowIndex = i + 1;
        break;
      }
    }

    const shiftStr = data.shifts.join(',');
    const now = new Date();

    if (rowIndex !== -1) {
      if (data.shifts.length === 0) {
        sheet.deleteRow(rowIndex);
      } else {
        sheet.getRange(rowIndex, 4).setValue(shiftStr); 
        sheet.getRange(rowIndex, 5).setValue(now); 
      }
    } else {
      if (data.shifts.length > 0) {
        sheet.appendRow([
          'D_' + Date.now(),
          data.user_id,
          targetDateStr,
          shiftStr,
          now
        ]);
      }
    }
    return JSON.stringify({ isOk: true });
  } catch (e) {
    return JSON.stringify({ isOk: false, message: e.toString() });
  } finally {
    lock.releaseLock();
  }
}

// --- REPORT FUNCTIONS ---
function getUsersByGroup(groupName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_USERS);
  if (!sheet) return JSON.stringify([]);
  
  const data = sheet.getDataRange().getDisplayValues();
  const users = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const groups = row[4];
    if (groups && groups.includes(groupName)) {
      users.push({
        user_id: row[1],
        name: `${row[2]} ${row[3]}`
      });
    }
  }
  users.sort((a, b) => a.name.localeCompare(b.name, 'th'));
  return JSON.stringify(users);
}

function getAllUsers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_USERS);
  if (!sheet) return JSON.stringify([]);
  
  const data = sheet.getDataRange().getDisplayValues();
  const users = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    users.push({
      user_id: row[1],
      name: `${row[2]} ${row[3]}`,
      groups: row[4]
    });
  }
  users.sort((a, b) => a.name.localeCompare(b.name, 'th'));
  return JSON.stringify(users);
}

function getMonthlyReport(monthStr, yearStr, targetGroup, targetUserId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_RECORDS);
    if (!sheet) return JSON.stringify({ isOk: false, message: "ไม่พบ Sheet TimeRecords" });

    const data = sheet.getDataRange().getValues();
    const tz = Session.getScriptTimeZone();
    const targetMonth = parseInt(monthStr, 10);
    const targetYear = parseInt(yearStr, 10);

    const result = [];
    const customOrderAP = ["รุ่งตะวัน", "ปวรวรรชน์", "พิสิฏฐ์", "ธนภรณ์", "บุษบา", "รัชนี"];
    const thaiMonthsShort = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."];

    const GROUP_AP = "พยาธิวิทยากายวิภาค";
    const GROUP_CP = "พยาธิวิทยาคลินิกและเทคนิคการแพทย์";

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      let dObj = null;
      if (row[5] instanceof Date) { dObj = row[5]; }
      else if (typeof row[5] === 'string' && row[5].trim() !== "") { dObj = new Date(row[5]); }
      
      if (!dObj || isNaN(dObj.getTime())) continue;
      
      const m = dObj.getMonth() + 1;
      const y = dObj.getFullYear();
      if (m !== targetMonth || y !== targetYear) continue;

      const recordUserId = String(row[1]);
      if (targetUserId && targetUserId !== "ALL") {
        if (recordUserId !== targetUserId) continue;
      }

      const workType = row[9] ? String(row[9]).trim() : "";
      const recordGroup = row[4] ? String(row[4]).trim() : "";
      
      let includeRecord = false;

      if (targetGroup === GROUP_CP) {
        if (workType === "เจาะเลือดเช้า" || workType === "เวรแล็บ") {
          includeRecord = true;
        } else if (workType === "" && recordGroup.includes(GROUP_CP)) {
          includeRecord = true;
        }
      } else if (targetGroup === GROUP_AP) {
        if (workType === "เวร HPV") {
          includeRecord = true;
        } else if (workType === "" && recordGroup.includes(GROUP_AP)) {
          includeRecord = true;
        }
      }

      if (includeRecord) {
        const day = dObj.getDate();
        const monthIndex = dObj.getMonth();
        const yearBE = dObj.getFullYear() + 543;
        const dateThaiStr = `${day} ${thaiMonthsShort[monthIndex]} ${yearBE}`;
        
        result.push({
          dateObj: dObj,
          dateDisplay: dateThaiStr,
          fullName: `${row[2]} ${row[3]}`,
          firstName: String(row[2]).trim(),
          timeIn: _formatTimeForHtml(row[6], tz),
          timeOut: _formatTimeForHtml(row[8], tz),
          workType: workType
        });
      }
    }

    result.sort((a, b) => {
      if (a.dateObj.getTime() !== b.dateObj.getTime()) {
        return a.dateObj.getTime() - b.dateObj.getTime();
      }
      if (targetGroup === GROUP_AP) {
        let idxA = customOrderAP.indexOf(a.firstName);
        let idxB = customOrderAP.indexOf(b.firstName);
        if (idxA === -1) idxA = 999;
        if (idxB === -1) idxB = 999;
        return idxA - idxB;
      } else {
        if (a.timeIn < b.timeIn) return -1;
        if (a.timeIn > b.timeIn) return 1;
        return 0;
      }
    });

    const finalOutput = result.map(r => ({
      date: r.dateDisplay,
      name: r.fullName,
      in: r.timeIn,
      out: r.timeOut
    }));

    return JSON.stringify({ isOk: true, data: finalOutput });

  } catch (e) {
    return JSON.stringify({ isOk: false, message: e.toString() });
  }
}

// --- ADMIN ADVANCE SCHEDULE ---
function getAdvanceSchedule(month, year) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_ADV_SCHED);
  if (!sheet) return JSON.stringify([]);

  const data = sheet.getDataRange().getValues();
  const results = [];
  const targetM = parseInt(month, 10);
  const targetY = parseInt(year, 10);
  const tz = Session.getScriptTimeZone();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    let dObj = null;
    if (row[3] instanceof Date) dObj = row[3];
    else if (typeof row[3] === 'string') dObj = new Date(row[3]);

    if (dObj && !isNaN(dObj.getTime())) {
      if ((dObj.getMonth() + 1) === targetM && dObj.getFullYear() === targetY) {
        results.push({
          id: row[0],
          user_id: row[1],
          name: row[2],
          date: Utilities.formatDate(dObj, tz, "yyyy-MM-dd"),
          shift_type: row[4],
          hours: parseFloat(row[5]) || 0
        });
      }
    }
  }
  return JSON.stringify(results);
}

function saveAdvanceSchedule(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_ADV_SCHED);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_ADV_SCHED);
      sheet.appendRow(["id", "user_id", "name", "date", "shift_type", "hours", "updated_at"]);
      sheet.setFrozenRows(1);
    }

    const data = JSON.parse(payload);
    const allData = sheet.getDataRange().getValues();
    const now = new Date();
    let rowIndex = -1;

    if (data.id) {
      for (let i = 1; i < allData.length; i++) {
        if (String(allData[i][0]) === String(data.id)) {
          rowIndex = i + 1;
          break;
        }
      }
    }

    if (rowIndex !== -1) {
      sheet.getRange(rowIndex, 2).setValue(data.user_id);
      sheet.getRange(rowIndex, 3).setValue(data.name);
      sheet.getRange(rowIndex, 4).setValue(data.date);
      sheet.getRange(rowIndex, 5).setValue(data.shift_type);
      sheet.getRange(rowIndex, 6).setValue(data.hours);
      sheet.getRange(rowIndex, 7).setValue(now);
    } else {
      const newId = 'ADV_' + Date.now() + Math.floor(Math.random() * 1000);
      sheet.appendRow([
        newId, data.user_id, data.name, data.date, data.shift_type, data.hours, now
      ]);
    }

    return JSON.stringify({ isOk: true });
  } catch (e) {
    return JSON.stringify({ isOk: false, message: e.toString() });
  } finally {
    lock.releaseLock();
  }
}

function deleteAdvanceSchedule(id) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_ADV_SCHED);
    if (!sheet) return JSON.stringify({ isOk: false, message: "Sheet not found" });

    const allData = sheet.getDataRange().getValues();
    let rowIndex = -1;

    for (let i = 1; i < allData.length; i++) {
      if (String(allData[i][0]) === String(id)) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex !== -1) {
      sheet.deleteRow(rowIndex);
      return JSON.stringify({ isOk: true });
    } else {
      return JSON.stringify({ isOk: false, message: "Data not found" });
    }
  } catch (e) {
    return JSON.stringify({ isOk: false, message: e.toString() });
  } finally {
    lock.releaseLock();
  }
}


// --- User Profile Management ---
function changePassword(userId, newPassword) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_USERS);
    if (!sheet) return JSON.stringify({ isOk: false, message: "Sheet User ไม่พบ" });

    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]) === String(userId)) {
        rowIndex = i + 1; 
        break;
      }
    }

    if (rowIndex !== -1) {
      sheet.getRange(rowIndex, 6).setValue(newPassword);
      return JSON.stringify({ isOk: true });
    } else {
      return JSON.stringify({ isOk: false, message: "ไม่พบผู้ใช้งานนี้" });
    }

  } catch (e) {
    return JSON.stringify({ isOk: false, message: e.toString() });
  } finally {
    lock.releaseLock();
  }
}

// --- TELEGRAM FUNCTION ---
function sendTelegramMsg(message) {
  if (!TELEGRAM_TOKEN || TELEGRAM_TOKEN === "XXXX" || !TELEGRAM_CHAT_ID || TELEGRAM_CHAT_ID === "XXXX") {
      console.log("Telegram not configured.");
      return;
  }
  const url = "https://api.telegram.org/bot" + TELEGRAM_TOKEN + "/sendMessage";
  const options = {
    "method": "post",
    "payload": {
      "chat_id": TELEGRAM_CHAT_ID,
      "text": message,
      "parse_mode": "HTML"
    },
    "muteHttpExceptions": true
  };
  try { UrlFetchApp.fetch(url, options); } catch (e) { console.log(e); }
}


// =========================================================================
// DAILY EMAIL NOTIFICATION SYSTEM (TRIGGER MANAGEMENT)
// =========================================================================

function checkEmailTriggerStatus() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'sendDailyShiftAlerts') {
      return true;
    }
  }
  return false;
}

function removeDailyTrigger() {
  try {
      const triggers = ScriptApp.getProjectTriggers();
      let removed = false;
      for (let i = 0; i < triggers.length; i++) {
        if (triggers[i].getHandlerFunction() === 'sendDailyShiftAlerts') {
          ScriptApp.deleteTrigger(triggers[i]);
          removed = true;
        }
      }
      return removed;
  } catch(e) {
      throw new Error(e.message);
  }
}

function setupDailyTrigger() {
  try {
      removeDailyTrigger(); // ลบของเก่าออกก่อน
      
      ScriptApp.newTrigger('sendDailyShiftAlerts')
        .timeBased()
        .everyDays(1)
        .atHour(5) // ระบบจะรันในช่วงเวลา 05:00 - 06:00
        .create();
        
      console.log("ตั้งค่า Trigger สำหรับแจ้งเตือนเวรรายวันเรียบร้อยแล้ว (ทำงานทุกวัน 05:00 - 06:00 น.)");
      return true;
  } catch(e) {
      throw new Error(e.message);
  }
}

function sendDailyShiftAlerts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userSheet = ss.getSheetByName(SHEET_USERS);
  const advSheet = ss.getSheetByName(SHEET_ADV_SCHED);

  if (!userSheet || !advSheet) {
    console.log("Error: ไม่พบ Sheet Users หรือ AdvanceSchedule");
    return;
  }

  const today = new Date();
  const tz = Session.getScriptTimeZone();
  const todayStr = Utilities.formatDate(today, tz, "yyyy-MM-dd");
  const todayThaiStr = _getThaiDateString(today);

  const userData = userSheet.getDataRange().getDisplayValues();
  const userMap = {}; 
  for (let i = 1; i < userData.length; i++) {
    let uId = String(userData[i][1]).trim();
    let fName = String(userData[i][2]).trim();
    let lName = String(userData[i][3]).trim();
    let email = userData[i][10] ? String(userData[i][10]).trim() : ""; 

    if (uId) {
      userMap[uId] = {
        name: `${fName} ${lName}`,
        email: email
      };
    }
  }

  const advData = advSheet.getDataRange().getValues();
  const todayShiftsMap = {}; 

  for (let i = 1; i < advData.length; i++) {
    let dateVal = advData[i][3];
    let rowDateStr = "";

    if (dateVal instanceof Date) {
      rowDateStr = Utilities.formatDate(dateVal, tz, "yyyy-MM-dd");
    } else {
      rowDateStr = String(dateVal).trim();
    }

    if (rowDateStr === todayStr) {
      let uId = String(advData[i][1]).trim();
      let shiftType = String(advData[i][4]).trim();
      let hours = parseFloat(advData[i][5]) || 0;

      if (!todayShiftsMap[uId]) todayShiftsMap[uId] = [];
      todayShiftsMap[uId].push({ type: shiftType, hours: hours });
    }
  }

  for (let uId in todayShiftsMap) {
    if (userMap[uId] && userMap[uId].email && userMap[uId].email.includes("@")) {
      let userName = userMap[uId].name;
      let userEmail = userMap[uId].email;
      let shifts = todayShiftsMap[uId];

      let htmlBody = generateShiftEmailHtml(userName, todayThaiStr, shifts);

      try {
        MailApp.sendEmail({
          to: userEmail,
          subject: `🔔 แจ้งเตือนตารางปฏิบัติงานประจำวัน - FineStamp`,
          htmlBody: htmlBody,
          name: "FineStamp"
        });
        console.log(`Sent email to ${userEmail} for user ${userName}`);
      } catch (err) {
        console.error(`Failed to send email to ${userEmail}: ${err.message}`);
      }
    }
  }
}

function generateShiftEmailHtml(userName, dateThaiStr, shifts) {
  let shiftsHtml = "";
  let totalHours = 0;

  shifts.forEach((shift, index) => {
    totalHours += shift.hours;
    shiftsHtml += `
      <tr>
        <td style="padding: 12px 15px; border-bottom: 1px solid #e2e8f0; color: #334155; font-size: 15px;">
          <span style="display: inline-block; width: 8px; height: 8px; background-color: #0ea5e9; border-radius: 50%; margin-right: 8px;"></span>
          ${shift.type}
        </td>
        <td style="padding: 12px 15px; border-bottom: 1px solid #e2e8f0; color: #64748b; font-size: 15px; text-align: right; font-weight: 600;">
          ${shift.hours} ชม.
        </td>
      </tr>
    `;
  });

  return `
    <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f8fafc; padding: 30px 15px; margin: 0;">
      <div style="max-width: 600px; margin: 0 auto; background-color: #ffffff; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 15px rgba(0,0,0,0.05);">
        
        <div style="background-color: #0ea5e9; padding: 25px 20px; text-align: center;">
          <h1 style="color: #ffffff; margin: 0; font-size: 24px; font-weight: 600; letter-spacing: 0.5px;">FineStamp</h1>
          <p style="color: #e0f2fe; margin: 5px 0 0 0; font-size: 14px;">ระบบแจ้งเตือนตารางปฏิบัติงานประจำวัน</p>
        </div>

        <div style="padding: 30px 25px;">
          <h2 style="color: #0f172a; font-size: 18px; margin-top: 0; margin-bottom: 20px;">สวัสดีครับ คุณ ${userName},</h2>
          <p style="color: #475569; font-size: 15px; line-height: 1.6; margin-bottom: 25px;">
            ระบบ <strong>FineStamp</strong> ขอแจ้งสรุปภาระงานและตารางปฏิบัติงานล่วงหน้าของคุณ ประจำวันที่ <strong style="color: #0ea5e9;">${dateThaiStr}</strong> ดังรายการต่อไปนี้ครับ
          </p>

          <table style="width: 100%; border-collapse: collapse; margin-bottom: 25px; border: 1px solid #e2e8f0; border-radius: 8px; overflow: hidden;">
            <thead>
              <tr style="background-color: #f1f5f9;">
                <th style="padding: 12px 15px; text-align: left; color: #475569; font-size: 14px; font-weight: 600; border-bottom: 2px solid #e2e8f0;">ประเภทงาน / เวร</th>
                <th style="padding: 12px 15px; text-align: right; color: #475569; font-size: 14px; font-weight: 600; border-bottom: 2px solid #e2e8f0;">จำนวนเวลา</th>
              </tr>
            </thead>
            <tbody>
              ${shiftsHtml}
            </tbody>
            <tfoot>
              <tr>
                <td style="padding: 12px 15px; text-align: right; color: #0f172a; font-size: 14px; font-weight: 700;">รวมภาระงานวันนี้:</td>
                <td style="padding: 12px 15px; text-align: right; color: #0ea5e9; font-size: 16px; font-weight: 700;">${totalHours} ชม.</td>
              </tr>
            </tfoot>
          </table>

          <div style="text-align: center; margin-top: 30px;">
            <p style="color: #64748b; font-size: 14px; margin-bottom: 10px;">อย่าลืมเข้าไปกด <strong>ลงเวลาปฏิบัติงาน</strong> ในระบบด้วยนะครับ</p>
          </div>
        </div>

        <div style="background-color: #f8fafc; padding: 20px 25px; border-top: 1px solid #e2e8f0; text-align: center;">
          <p style="font-size: 12px; color: #94a3b8; margin: 0 0 10px 0; line-height: 1.5;">
            อีเมลฉบับนี้ถูกส่งโดยอัตโนมัติจากระบบ FineStamp กรุณาอย่าตอบกลับ
          </p>
          <p style="font-size: 11px; color: #cbd5e1; margin: 0; line-height: 1.5;">
            <strong>ประกาศความเป็นส่วนตัว (PDPA):</strong> ข้อมูลตารางปฏิบัติงานนี้ถือเป็นข้อมูลความลับส่วนบุคคล ทางระบบมีการจัดเก็บและประมวลผลเพื่อประโยชน์ในการบริหารจัดการภายในองค์กรเท่านั้น โปรดอย่าส่งต่ออีเมลนี้ให้บุคคลอื่นที่ไม่เกี่ยวข้อง
          </p>
        </div>

      </div>
    </div>
  `;
}

// =========================================================================
// TELEGRAM NOTIFICATION SYSTEM (TOMORROW SHIFTS) - NEW in v1.3.0
// =========================================================================

function checkTelegramTriggerStatus() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'sendTomorrowShiftTelegramAlerts') {
      return true;
    }
  }
  return false;
}

function removeTelegramTrigger() {
  try {
      const triggers = ScriptApp.getProjectTriggers();
      let removed = false;
      for (let i = 0; i < triggers.length; i++) {
        if (triggers[i].getHandlerFunction() === 'sendTomorrowShiftTelegramAlerts') {
          ScriptApp.deleteTrigger(triggers[i]);
          removed = true;
        }
      }
      return removed;
  } catch(e) {
      throw new Error(e.message);
  }
}

function setupTelegramTrigger() {
  try {
      removeTelegramTrigger(); // ลบของเก่าออกก่อน
      
      ScriptApp.newTrigger('sendTomorrowShiftTelegramAlerts')
        .timeBased()
        .everyDays(1)
        .atHour(16) // ระบบจะรันในช่วงเวลา 16:00 - 17:00 น.
        .create();
        
      console.log("ตั้งค่า Trigger สำหรับแจ้งเตือนเวรพรุ่งนี้ (Telegram) เรียบร้อยแล้ว (ทำงานทุกวัน 16:00 - 17:00 น.)");
      return true;
  } catch(e) {
      throw new Error(e.message);
  }
}

function sendTomorrowShiftTelegramAlerts() {
  if (!TELEGRAM_TOKEN || TELEGRAM_TOKEN === "XXXX" || !TELEGRAM_CHAT_ID || TELEGRAM_CHAT_ID === "XXXX") {
      console.log("Telegram not configured.");
      return;
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const advSheet = ss.getSheetByName(SHEET_ADV_SCHED);
  if (!advSheet) return;

  // หาวันที่พรุ่งนี้
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  const tz = Session.getScriptTimeZone();
  const tomorrowStr = Utilities.formatDate(tomorrow, tz, "yyyy-MM-dd");
  const tomorrowThaiStr = _getThaiDateString(tomorrow);

  const advData = advSheet.getDataRange().getValues();
  
  // จัดกลุ่มรายชื่อตามประเภทเวร เพื่อให้แสดงผลอ่านง่าย
  const shiftsMap = {}; 
  let hasShifts = false;

  for (let i = 1; i < advData.length; i++) {
    let dateVal = advData[i][3];
    let rowDateStr = "";
    
    if (dateVal instanceof Date) {
      rowDateStr = Utilities.formatDate(dateVal, tz, "yyyy-MM-dd");
    } else {
      rowDateStr = String(dateVal).trim();
    }

    if (rowDateStr === tomorrowStr) {
      let name = String(advData[i][2]).trim();
      let shiftType = String(advData[i][4]).trim();
      let hours = parseFloat(advData[i][5]) || 0;

      if (!shiftsMap[shiftType]) {
          shiftsMap[shiftType] = [];
      }
      shiftsMap[shiftType].push({ name: name, hours: hours });
      hasShifts = true;
    }
  }

  // ถ้าไม่มีเวรพรุ่งนี้เลย ก็ไม่ต้องส่ง
  if (!hasShifts) return; 

  // --- สร้างข้อความ Telegram (HTML Parse Mode) แตกต่างจากแจ้งเตือนลงเวลา ---
  let msg = `📅 <b>แจ้งเตือนเวรวันพรุ่งนี้</b>\n`;
  msg += `ประจำวันที่: <b>${tomorrowThaiStr}</b>\n`;
  msg += `➖➖➖➖➖➖➖➖➖➖\n\n`;

  // ลำดับการแสดงผลประเภทเวร (เพื่อความสวยงามและเป็นระเบียบ)
  const shiftOrder = ["เจาะเลือดเช้า", "เวรแล็บ", "เวร HPV"];
  const allShiftTypes = Object.keys(shiftsMap);
  
  allShiftTypes.sort((a, b) => {
      let idxA = shiftOrder.indexOf(a);
      let idxB = shiftOrder.indexOf(b);
      idxA = idxA === -1 ? 999 : idxA; // ถ้าเป็นเวรประเภทอื่นให้อยู่ท้าย
      idxB = idxB === -1 ? 999 : idxB;
      return idxA - idxB;
  });

  // วนลูปสร้างข้อความตามประเภทเวร
  for (let type of allShiftTypes) {
      // ใส่ Emoji ตามประเภทเวร
      let icon = "🏥";
      if (type.includes("เจาะเลือด")) icon = "💉";
      else if (type.includes("แล็บ") || type.includes("Lab")) icon = "🔬";
      else if (type.includes("HPV")) icon = "🧬";

      msg += `${icon} <b>${type}</b>\n`;
      for (let shift of shiftsMap[type]) {
          msg += `  ▫️ ${shift.name} (${shift.hours} ชม.)\n`;
      }
      msg += `\n`;
  }

  msg += `➖➖➖➖➖➖➖➖➖➖\n`;
  msg += `<i>แจ้งเตือนอัตโนมัติจาก FineStamp</i>`;

  sendTelegramMsg(msg);
}
