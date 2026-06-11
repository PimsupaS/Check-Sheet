// ================================================================
// ISUZU FRG — Apps Script v3
// รองรับ GET (payload param), POST (JSON body), POST (form payload)
// ================================================================

const SHEET_NAME = "submissions";

function doGet(e) {
  try {
    // ถ้ามี payload parameter = Worker ส่งข้อมูลมา
    if (e.parameter && e.parameter.payload) {
      return saveData(JSON.parse(e.parameter.payload));
    }
    // ถ้าไม่มี = Dashboard ขอดูข้อมูล
    return getData();
  } catch (err) {
    return buildResponse({ result: "error", message: err.toString() });
  }
}

function doPost(e) {
  try {
    let data;
    if (e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else if (e.parameter && e.parameter.payload) {
      data = JSON.parse(e.parameter.payload);
    } else {
      return buildResponse({ result: "error", message: "No data" });
    }
    return saveData(data);
  } catch (err) {
    return buildResponse({ result: "error", message: err.toString() });
  }
}

function saveData(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow([
      "ID","วันที่","กะ","Stage","Group","แผนก",
      "Worker","เวลา Worker","Foreman","เวลา Foreman",
      "ผ่าน","ไม่ผ่าน","รายการไม่ผ่าน","หมายเหตุ",
      "สถานะ","JSON เต็ม"
    ]);
    sheet.getRange(1,1,1,16)
      .setFontWeight("bold")
      .setBackground("#CC0000")
      .setFontColor("white");
    sheet.setFrozenRows(1);
  }

  const checks    = data.checks || [];
  const passCount = checks.filter(c => c.result === "pass").length;
  const failCount = checks.filter(c => c.result === "fail").length;
  const failNames = checks
    .filter(c => c.result === "fail")
    .map(c => c.name.replace(/\n/g," "))
    .join(", ");

  sheet.appendRow([
    data.id || "",
    data.date || "",
    data.shift || "",
    data.stage || "",
    data.group || "",
    data.dept_name || "",
    data.worker || "",
    data.timestamp_worker || "",
    data.foreman || "",
    data.timestamp_foreman || "",
    passCount,
    failCount,
    failNames,
    data.remark || "",
    data.status || "pending",
    JSON.stringify(data)
  ]);

  return buildResponse({ result: "success" });
}

function getData() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return buildResponse({ result: "success", rows: [] });

  const values  = sheet.getDataRange().getValues();
  const headers = values[0];
  const rows    = values.slice(1).map(row => {
    const obj = {};
    headers.forEach((h,i) => { obj[h] = row[i]; });
    return obj;
  });
  return buildResponse({ result: "success", rows });
}

function buildResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
