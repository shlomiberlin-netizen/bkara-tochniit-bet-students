// ===== ריכוז בקרות שטח — Apps Script =====
// תומך בשני סוגי טפסים: תכנית ב' (גיליון1) ומכרז פריפריה (גיליון "פריפריה")

// ---- שמות גיליונות ----
var SHEET_NAME_B = "גיליון1";      // תכנית ב' — מעגל השנה
var SHEET_NAME_P = "פריפריה";      // מכרז פריפריה

// ---- מזהי שדות — תכנית ב' ----
var IDS_B = [
  "s1i1","s1i2","s1i3",
  "s2i1","s2i2","s2i3","s2i4",
  "s3i1","s3i2","s3i3","s3i4",
  "s4i1","s4i2","s4i3","s4i4",
  "s5i1","s5i2","s5i3",
  "s6i1","s6i2","s6i3",
  "s7i1","s7i2","s7i3",
  "s8i1","s8i2","s8i3","s8i4"
];

// ---- מזהי שדות — פריפריה ----
var IDS_P = [
  "p1i1","p1i2","p1i3",
  "p2i1",
  "p3i1","p3i2","p3i3","p3i4","p3i5",
  "p4i1",
  "p5i1","p5i2","p5i3",
  "p6i1","p6i2",
  "p7i1","p7i2","p7i3"
];

// ---- קבלת גיליונות ----
function getSheetB() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_B)
         || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
}

function getSheetP() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME_P);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME_P);
    var headers = ["תאריך", "שם בקר", "עיר/יישוב", "שם סניף"];
    IDS_P.forEach(function(id) { headers.push(id); });
    IDS_P.forEach(function(id) { headers.push("הערות_" + id); });
    headers.push("ממצאים עיקריים");
    headers.push("פריטים שאינם עומדים");
    headers.push("המלצות לתיקון");
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sheet;
}

// ---- בניית שורות ----
function buildRowB(data) {
  var row = [
    data.visitDate || new Date(),
    data.inspector || "",
    data.branch || "",
    data.institution || ""
  ];
  IDS_B.forEach(function(id) { row.push(data[id] ? (data[id].status || "") : ""); });
  IDS_B.forEach(function(id) { row.push(data[id] ? (data[id].note || "") : ""); });
  return row;
}

function buildRowP(data) {
  var row = [
    data.visitDate || new Date(),
    data.inspector || "",
    data.city || "",
    data.branch || ""
  ];
  IDS_P.forEach(function(id) { row.push(data[id] ? (data[id].status || "") : ""); });
  IDS_P.forEach(function(id) { row.push(data[id] ? (data[id].note || "") : ""); });
  row.push(data.findings || "");
  row.push(data.nonCompliant || "");
  row.push(data.recommendations || "");
  return row;
}

// ---- חיפוש שורה לפי מפתח ----
// לתכנית ב': עמודה B=שם בקר (index 1), עמודה C=שם סניף (index 2)
function findRowByKeyB(sheet, key) {
  var parts = key.split("||");
  var inspector = parts[1] ? parts[1].toString().trim() : "";
  var branch = parts[2] ? parts[2].toString().trim() : "";
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var rowInsp = data[i][1] ? data[i][1].toString().trim() : "";
    var rowBranch = data[i][2] ? data[i][2].toString().trim() : "";
    if (rowInsp === inspector && rowBranch === branch) {
      return i + 1; // מספר שורה (1-based)
    }
  }
  return -1;
}

// לפריפריה: עמודה B=שם בקר (index 1), עמודה D=שם סניף (index 3)
function findRowByKeyP(sheet, key) {
  var parts = key.split("||");
  var inspector = parts[1] ? parts[1].toString().trim() : "";
  var branch = parts[2] ? parts[2].toString().trim() : "";
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var rowInsp = data[i][1] ? data[i][1].toString().trim() : "";
    var rowBranch = data[i][3] ? data[i][3].toString().trim() : "";
    if (rowInsp === inspector && rowBranch === branch) {
      return i + 1;
    }
  }
  return -1;
}

// ---- נקודת כניסה POST ----
function doPost(e) {
  var data = JSON.parse(e.postData.contents);

  if (data.formType === "periphery") {
    return handlePeriphery(data);
  } else {
    return handleTochniitBet(data);
  }
}

// ---- טיפול בטופס תכנית ב' ----
function handleTochniitBet(data) {
  var sheet = getSheetB();

  if (data.action === "delete") {
    var rowNum = parseInt(data.rowIndex) + 2;
    sheet.deleteRow(rowNum);
    return respond("deleted");
  }

  if (data.action === "deleteByKey") {
    var rowNum = findRowByKeyB(sheet, data.origKey);
    if (rowNum > 0) sheet.deleteRow(rowNum);
    return respond("deleted");
  }

  if (data.action === "update") {
    var rowNum = parseInt(data.rowIndex) + 2;
    var row = buildRowB(data);
    sheet.getRange(rowNum, 1, 1, row.length).setValues([row]);
    return respond("updated");
  }

  if (data.action === "updateByKey") {
    var rowNum = findRowByKeyB(sheet, data.origKey);
    if (rowNum > 0) {
      var row = buildRowB(data);
      sheet.getRange(rowNum, 1, 1, row.length).setValues([row]);
    }
    return respond("updated");
  }

  sheet.appendRow(buildRowB(data));
  return respond("success");
}

// ---- טיפול בטופס פריפריה ----
function handlePeriphery(data) {
  var sheet = getSheetP();

  if (data.action === "delete") {
    var rowNum = parseInt(data.rowIndex) + 2;
    sheet.deleteRow(rowNum);
    return respond("deleted");
  }

  if (data.action === "deleteByKey") {
    var rowNum = findRowByKeyP(sheet, data.origKey);
    if (rowNum > 0) sheet.deleteRow(rowNum);
    return respond("deleted");
  }

  if (data.action === "update") {
    var rowNum = parseInt(data.rowIndex) + 2;
    var row = buildRowP(data);
    sheet.getRange(rowNum, 1, 1, row.length).setValues([row]);
    return respond("updated");
  }

  if (data.action === "updateByKey") {
    var rowNum = findRowByKeyP(sheet, data.origKey);
    if (rowNum > 0) {
      var row = buildRowP(data);
      sheet.getRange(rowNum, 1, 1, row.length).setValues([row]);
    }
    return respond("updated");
  }

  sheet.appendRow(buildRowP(data));
  return respond("success");
}

// ---- נקודת כניסה GET ----
// ?sheet=periphery → גיליון פריפריה
// ללא פרמטר → גיליון תכנית ב'
function doGet(e) {
  if (e && e.parameter && e.parameter.sheet === "periphery") {
    return respondWithRows(getSheetP());
  } else {
    return respondWithRows(getSheetB());
  }
}

function respondWithRows(sheet) {
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return respond2([]);
  var headers = data[0];
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var row = {};
    for (var j = 0; j < headers.length; j++) {
      row[headers[j]] = data[i][j];
    }
    rows.push(row);
  }
  return respond2(rows);
}

// ---- פונקציות עזר ----
function respond(result) {
  return ContentService
    .createTextOutput(JSON.stringify({ result: result }))
    .setMimeType(ContentService.MimeType.JSON);
}

function respond2(rows) {
  return ContentService
    .createTextOutput(JSON.stringify({ rows: rows }))
    .setMimeType(ContentService.MimeType.JSON);
}
