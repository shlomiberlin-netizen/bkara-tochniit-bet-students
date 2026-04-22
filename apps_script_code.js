// ===== תכנית ב׳ — בקרות שטח =====
// גרסה 4: תמיכה מלאה בהוספה / עדכון לפי אינדקס / עדכון לפי מפתח / מחיקה

var SHEET_NAME = "גיליון1";

var IDS = ["s1i1","s1i2","s1i3","s2i1","s2i2","s2i3","s2i4","s3i1","s3i2","s3i3","s3i4","s4i1","s4i2","s4i3","s4i4","s5i1","s5i2","s5i3","s6i1","s6i2","s6i3","s7i1","s7i2","s7i3","s8i1","s8i2","s8i3","s8i4"];

function getSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME)
         || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
}

function buildRow(data) {
  var row = [data.visitDate || new Date(), data.inspector || "", data.branch || "", data.institution || ""];
  IDS.forEach(function(id) { row.push(data[id] ? (data[id].status || "") : ""); });
  IDS.forEach(function(id) { row.push(data[id] ? (data[id].note || "") : ""); });
  return row;
}

function findRowByKey(sheet, key) {
  // key = "תאריך||שם בקר||שם סניף"
  var parts = key.split("||");
  var dateKey = parts[0] ? parts[0].toString().trim() : "";
  var inspector = parts[1] ? parts[1].toString().trim() : "";
  var branch = parts[2] ? parts[2].toString().trim() : "";
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var rowDate = data[i][0] ? data[i][0].toString().trim() : "";
    var rowInsp = data[i][1] ? data[i][1].toString().trim() : "";
    var rowBranch = data[i][2] ? data[i][2].toString().trim() : "";
    if (rowInsp === inspector && rowBranch === branch) {
      // match by inspector + branch (date formats can vary)
      return i + 1; // 1-based row number
    }
  }
  return -1;
}

function doPost(e) {
  var sheet = getSheet();
  var data = JSON.parse(e.postData.contents);

  // מחיקה לפי אינדקס
  if (data.action === "delete") {
    var rowNum = parseInt(data.rowIndex) + 2;
    sheet.deleteRow(rowNum);
    return respond("deleted");
  }

  // מחיקה לפי מפתח (שם בקר + סניף)
  if (data.action === "deleteByKey") {
    var rowNum = findRowByKey(sheet, data.origKey);
    if (rowNum > 0) sheet.deleteRow(rowNum);
    return respond("deleted");
  }

  // עדכון לפי אינדקס
  if (data.action === "update") {
    var rowNum = parseInt(data.rowIndex) + 2;
    var row = buildRow(data);
    sheet.getRange(rowNum, 1, 1, row.length).setValues([row]);
    return respond("updated");
  }

  // עדכון לפי מפתח
  if (data.action === "updateByKey") {
    var rowNum = findRowByKey(sheet, data.origKey);
    if (rowNum > 0) {
      var row = buildRow(data);
      sheet.getRange(rowNum, 1, 1, row.length).setValues([row]);
    }
    return respond("updated");
  }

  // הוספת שורה חדשה
  sheet.appendRow(buildRow(data));
  return respond("success");
}

function doGet(e) {
  var sheet = getSheet();
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return respond2([]);

  var headers = data[0];
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var row = {};
    for (var j = 0; j < headers.length; j++) { row[headers[j]] = data[i][j]; }
    rows.push(row);
  }
  return respond2(rows);
}

function respond(result) {
  return ContentService.createTextOutput(JSON.stringify({ result: result })).setMimeType(ContentService.MimeType.JSON);
}

function respond2(rows) {
  return ContentService.createTextOutput(JSON.stringify({ rows: rows })).setMimeType(ContentService.MimeType.JSON);
}


var SHEET_NAME = "גיליון1"; // שנה אם שם הגיליון שלך שונה

function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME)
              || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = JSON.parse(e.postData.contents);

  // מחיקת שורה
  if (data.action === "delete") {
    var rowNum = parseInt(data.rowIndex) + 2; // +2 כי שורה 1 = כותרות, ו-index מתחיל מ-0
    sheet.deleteRow(rowNum);
    return ContentService
      .createTextOutput(JSON.stringify({ result: "deleted" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // עדכון שורה קיימת
  if (data.action === "update") {
    var rowNum = parseInt(data.rowIndex) + 2;
    var ids = [
      "s1i1","s1i2","s1i3",
      "s2i1","s2i2","s2i3","s2i4",
      "s3i1","s3i2","s3i3","s3i4",
      "s4i1","s4i2","s4i3","s4i4",
      "s5i1","s5i2","s5i3",
      "s6i1","s6i2","s6i3",
      "s7i1","s7i2","s7i3",
      "s8i1","s8i2","s8i3","s8i4"
    ];
    var row = [
      data.visitDate || "",
      data.inspector || "",
      data.branch || "",
      data.institution || ""
    ];
    ids.forEach(function(id) {
      row.push(data[id] ? (data[id].status || "") : "");
    });
    ids.forEach(function(id) {
      row.push(data[id] ? (data[id].note || "") : "");
    });
    var range = sheet.getRange(rowNum, 1, 1, row.length);
    range.setValues([row]);
    return ContentService
      .createTextOutput(JSON.stringify({ result: "updated" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // הוספת שורה חדשה (ברירת מחדל)
  var ids = [
    "s1i1","s1i2","s1i3",
    "s2i1","s2i2","s2i3","s2i4",
    "s3i1","s3i2","s3i3","s3i4",
    "s4i1","s4i2","s4i3","s4i4",
    "s5i1","s5i2","s5i3",
    "s6i1","s6i2","s6i3",
    "s7i1","s7i2","s7i3",
    "s8i1","s8i2","s8i3","s8i4"
  ];
  var row = [
    data.visitDate || new Date(),
    data.inspector || "",
    data.branch || "",
    data.institution || ""
  ];
  ids.forEach(function(id) {
    row.push(data[id] ? (data[id].status || "") : "");
  });
  ids.forEach(function(id) {
    row.push(data[id] ? (data[id].note || "") : "");
  });
  sheet.appendRow(row);

  return ContentService
    .createTextOutput(JSON.stringify({ result: "success" }))
    .setMimeType(ContentService.MimeType.JSON);
}

// --- החזרת נתונים לדשבורד ---
function doGet(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME)
              || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();

  if (data.length <= 1) {
    return ContentService
      .createTextOutput(JSON.stringify({ rows: [] }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var headers = data[0];
  var rows = [];

  for (var i = 1; i < data.length; i++) {
    var row = {};
    for (var j = 0; j < headers.length; j++) {
      row[headers[j]] = data[i][j];
    }
    rows.push(row);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ rows: rows }))
    .setMimeType(ContentService.MimeType.JSON);
}
