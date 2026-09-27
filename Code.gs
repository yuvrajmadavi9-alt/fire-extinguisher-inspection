function testFireSafetySheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();

  for (var i = 0; i < sheets.length; i++) {
    Logger.log(sheets[i].getName());
  }
}
function testNewInspection() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName("Fire Extinguisher Monthly Inspection");

  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();

  Logger.log("Last Row: " + lastRow);
  Logger.log("Last Column: " + lastColumn);
}
function testHeaders() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName("Fire Extinguisher Monthly Inspection");

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  for (var i = 0; i < headers.length; i++) {
    Logger.log((i + 1) + " = " + headers[i]);
  }
}
function testLatestInspection() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName("Fire Extinguisher Monthly Inspection");

  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();

  var headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  var data = sheet.getRange(lastRow, 1, 1, lastColumn).getValues()[0];

  Logger.log("----- LATEST INSPECTION -----");
  Logger.log("Row Number: " + lastRow);

  for (var i = 0; i < lastColumn; i++) {
    Logger.log((i + 1) + " | " + headers[i] + " = " + data[i]);
  }
}
function checkLatestInspection() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName("Fire Extinguisher Monthly Inspection");

  var lastRow = sheet.getLastRow();

  var inspection = sheet.getRange(lastRow, 7, 1, 6).getValues()[0];

  var parameters = [
    "Safety Pin",
    "Discharge Hose",
    "Discharge Nozzle / Horn",
    "Pressure / Weight",
    "Cap Assembly",
    "Inspection Sticker"
  ];

  var defects = [];

  for (var i = 0; i < inspection.length; i++) {
    var value = String(inspection[i]).trim().toUpperCase();

    if (value !== "OK") {
      defects.push(parameters[i] + " = " + inspection[i]);
    }
  }

  Logger.log("----- INSPECTION RESULT -----");

  if (defects.length === 0) {
    Logger.log("🟢 ALL OK");
  } else {
    Logger.log("🔴 DEFECT FOUND");

    for (var j = 0; j < defects.length; j++) {
      Logger.log(defects[j]);
    }
  }
}
function getLatestDefectDetails() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName("Fire Extinguisher Monthly Inspection");

  var lastRow = sheet.getLastRow();

  var data = sheet.getRange(lastRow, 1, 1, 17).getValues()[0];

  var inspectionDate = data[1];
  var location = data[2];
  var assetNo = data[3];
  var extinguisherType = data[4];
  var capacity = data[5];

  var parameters = [
    "Safety Pin",
    "Discharge Hose",
    "Discharge Nozzle / Horn",
    "Pressure / Weight",
    "Cap Assembly",
    "Inspection Sticker"
  ];

  var defects = [];

  for (var i = 0; i < 6; i++) {
    var value = String(data[i + 6]).trim().toUpperCase();

    if (value !== "OK") {
      defects.push(parameters[i] + " = " + data[i + 6]);
    }
  }

  var remarks = data[15];
  var checkedBy = data[16];

  Logger.log("----- DEFECT DETAILS -----");
  Logger.log("Inspection Date: " + inspectionDate);
  Logger.log("Asset No.: " + assetNo);
  Logger.log("Location: " + location);
  Logger.log("Extinguisher Type: " + extinguisherType);
  Logger.log("Capacity: " + capacity);

  if (defects.length === 0) {
    Logger.log("Status: ALL OK");
  } else {
    Logger.log("Status: DEFECT FOUND");

    for (var j = 0; j < defects.length; j++) {
      Logger.log("Defect: " + defects[j]);
    }
  }

  Logger.log("Remarks: " + remarks);
  Logger.log("Checked By: " + checkedBy);
}
function generateDefectID() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName("Fire Extinguisher Monthly Inspection");

  var lastRow = sheet.getLastRow();

  var data = sheet.getRange(lastRow, 1, 1, 17).getValues()[0];

  var assetNo = data[3];
  var defects = [];

  var parameters = [
    "Safety Pin",
    "Discharge Hose",
    "Discharge Nozzle / Horn",
    "Pressure / Weight",
    "Cap Assembly",
    "Inspection Sticker"
  ];

  for (var i = 0; i < 6; i++) {
    var value = String(data[i + 6]).trim().toUpperCase();

    if (value !== "OK") {
      defects.push(parameters[i] + " = " + data[i + 6]);
    }
  }

  Logger.log("----- DEFECT ID TEST -----");

  if (defects.length === 0) {
    Logger.log("Status: ALL OK");
    Logger.log("No Defect ID required.");
    return;
  }

  var year = new Date().getFullYear();

  var defectID = "DEF-" + year + "-" + String(lastRow - 1).padStart(3, "0");

  Logger.log("Status: DEFECT FOUND");
  Logger.log("Asset No.: " + assetNo);
  Logger.log("Defect ID: " + defectID);

  for (var j = 0; j < defects.length; j++) {
    Logger.log("Defect: " + defects[j]);
  }
}
function createDefectTracker() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var tracker = ss.getSheetByName("Defect Action Tracker");

  if (!tracker) {
    tracker = ss.insertSheet("Defect Action Tracker");
  }

  var headers = [
    "Defect ID",
    "Inspection Date",
    "Asset No.",
    "Location",
    "Extinguisher Type",
    "Capacity",
    "Defect",
    "Remarks",
    "Checked By",
    "Status",
    "Due Date",
    "Action Taken",
    "Closed Date"
  ];

  tracker.getRange(1, 1, 1, headers.length).setValues([headers]);

  tracker.getRange(1, 1, 1, headers.length)
      .setFontWeight("bold");

  tracker.setFrozenRows(1);

  Logger.log("----- DEFECT TRACKER -----");
  Logger.log("Defect Action Tracker created successfully.");
  Logger.log("Total Columns: " + headers.length);
}
function testAddDefectToTracker() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var tracker = ss.getSheetByName("Defect Action Tracker");

  if (!tracker) {
    Logger.log("Defect Action Tracker not found.");
    return;
  }

  var testDefectID = "TEST-001";

  var row = [
    testDefectID,
    new Date(),
    "TEST-306",
    "New workshop Fire stand",
    "CO₂",
    "4.5 Kg",
    "Pressure / Weight = NOT OK",
    "TEST DEFECT - This is only a test entry",
    "TEST USER",
    "Open",
    new Date(),
    "",
    ""
  ];

  tracker.appendRow(row);

  Logger.log("----- TEST DEFECT -----");
  Logger.log("Test defect added successfully.");
  Logger.log("Defect ID: " + testDefectID);
}
function addLatestRealDefectToTracker() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var inspectionSheet = ss.getSheetByName("Fire Extinguisher Monthly Inspection");
  var tracker = ss.getSheetByName("Defect Action Tracker");

  if (!inspectionSheet || !tracker) {
    Logger.log("Inspection Sheet or Defect Tracker not found.");
    return;
  }

  var lastRow = inspectionSheet.getLastRow();
  var data = inspectionSheet.getRange(lastRow, 1, 1, 17).getValues()[0];

  var defects = [];

  var parameters = [
    "Safety Pin",
    "Discharge Hose",
    "Discharge Nozzle / Horn",
    "Pressure / Weight",
    "Cap Assembly",
    "Inspection Sticker"
  ];

  for (var i = 0; i < 6; i++) {
    var value = String(data[i + 6]).trim();

    if (value.toUpperCase() !== "OK") {
      defects.push(parameters[i] + " = " + value);
    }
  }

  if (defects.length === 0) {
    Logger.log("ALL OK - No defect added.");
    return;
  }

  var year = new Date().getFullYear();
  var defectID = "DEF-" + year + "-" + String(lastRow - 1).padStart(3, "0");

  var inspectionDate = data[1];
  var assetNo = data[3];
  var location = data[2];
  var extinguisherType = data[4];
  var capacity = data[5];
  var remarks = data[15];
  var checkedBy = data[16];

  var defectText = defects.join(" | ");

  var dueDate = new Date();
  dueDate.setDate(dueDate.getDate() + 7);

  var row = [
    defectID,
    inspectionDate,
    assetNo,
    location,
    extinguisherType,
    capacity,
    defectText,
    remarks,
    checkedBy,
    "Open",
    dueDate,
    "",
    ""
  ];

  tracker.appendRow(row);

  Logger.log("----- REAL DEFECT ADDED -----");
  Logger.log("Defect ID: " + defectID);
  Logger.log("Asset No.: " + assetNo);
  Logger.log("Location: " + location);
  Logger.log("Defect: " + defectText);
  Logger.log("Status: Open");
  Logger.log("Due Date: " + dueDate);
}
function handleExtinguisherInspection(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var inspectionSheet = e.range.getSheet();
  var tracker = ss.getSheetByName("Defect Action Tracker");

  if (!tracker) {
    Logger.log("Defect Action Tracker not found.");
    return;
  }

  var rowNumber = e.range.getRow();

  // Header row skip
  if (rowNumber <= 1) {
    return;
  }

  var data = inspectionSheet
    .getRange(rowNumber, 1, 1, 17)
    .getValues()[0];

    // ==============================
// UPDATE ASSET MASTER
// ==============================

var assetMaster = ss.getSheetByName("Contractor Fire Extinguisher Master");

if (assetMaster) {

  var assetNo = String(data[3]).trim();
  var location = data[2];
  var extinguisherType = data[4];
  var capacity = data[5];
  var inspectionDate = data[1];

  if (assetNo) {

    var masterLastRow = assetMaster.getLastRow();

    if (masterLastRow >= 2) {

      var masterData = assetMaster
        .getRange(2, 1, masterLastRow - 1, 20)
        .getValues();

      for (var m = 0; m < masterData.length; m++) {

        var masterAssetNo = String(masterData[m][0]).trim();
var masterType = String(masterData[m][3]).trim().toUpperCase();
var masterCapacity = String(masterData[m][4]).trim().toUpperCase();

var formType = String(extinguisherType).trim().toUpperCase();
var formCapacity = String(capacity).trim().toUpperCase();

// ABC (DCP) and ABC should be treated as same type
if (formType === "ABC (DCP)") {
  formType = "ABC";
}

if (masterType === "ABC (DCP)") {
  masterType = "ABC";
}
Logger.log("FORM: Asset=" + assetNo + " | Type=" + formType + " | Capacity=" + formCapacity);
Logger.log("MASTER: Asset=" + masterAssetNo + " | Type=" + masterType + " | Capacity=" + masterCapacity);
Logger.log("MATCH=" + (masterAssetNo === assetNo && masterType === formType && masterCapacity === formCapacity));

if (
  masterAssetNo === assetNo &&
  masterType === formType &&
  masterCapacity === formCapacity
) {

          // Column H = Last Inspection Date
          assetMaster
            .getRange(m + 2, 8)
            .setValue(inspectionDate);

          // Column I = Inspection Status
          assetMaster
            .getRange(m + 2, 9)
            .setValue("Inspected");

          Logger.log("Asset Master updated: " + assetNo);

          break;
        }
      }
    }
  }
}

  var parameters = [
    "Safety Pin",
    "Discharge Hose",
    "Discharge Nozzle / Horn",
    "Pressure / Weight",
    "Cap Assembly",
    "Inspection Sticker"
  ];

  var defects = [];

  for (var i = 0; i < 6; i++) {
    var value = String(data[i + 6]).trim();

    if (value.toUpperCase() !== "OK") {
      defects.push(parameters[i] + " = " + value);
    }
  }

  // If everything is OK, stop
  if (defects.length === 0) {
    Logger.log("ALL OK - No defect added.");
    return;
  }

  var year = new Date().getFullYear();

  var defectID =
    "DEF-" +
    year +
    "-" +
    String(rowNumber - 1).padStart(3, "0");

  var inspectionDate = data[1];
  var assetNo = data[3];
  var location = data[2];
  var extinguisherType = data[4];
  var capacity = data[5];
  var remarks = data[15];
  var checkedBy = data[16];

  var defectText = defects.join(" | ");

  var dueDate = new Date();
  dueDate.setDate(dueDate.getDate() + 7);

  var trackerRow = [
    defectID,
    inspectionDate,
    assetNo,
    location,
    extinguisherType,
    capacity,
    defectText,
    remarks,
    checkedBy,
    "Open",
    dueDate,
    "",
    ""
  ];

  tracker.appendRow(trackerRow);



  Logger.log("----- AUTOMATIC DEFECT ADDED -----");
  Logger.log("Defect ID: " + defectID);
  Logger.log("Asset No.: " + assetNo);
  Logger.log("Location: " + location);
  Logger.log("Defect: " + defectText);
  Logger.log("Status: Open");
}
function handleDefectTrackerEdit(e) {

  var sheet = e.range.getSheet();

  if (sheet.getName() !== "Defect Action Tracker") {
    return;
  }

  var row = e.range.getRow();
  var col = e.range.getColumn();

  if (row === 1) {
    return;
  }

  if (col !== 12) {
    return;
  }

  var actionTaken = sheet.getRange(row, 12).getValue();

  var statusCell = sheet.getRange(row, 10);
  var closedDateCell = sheet.getRange(row, 13);

  if (actionTaken !== "") {

    statusCell.setValue("Closed");
    closedDateCell.setValue(new Date());

    var defectID = sheet.getRange(row, 1).getValue();

    var properties = PropertiesService.getScriptProperties();

    var mailKey = "CLOSED_MAIL_SENT_" + defectID;

    var alreadySent = properties.getProperty(mailKey);

    if (!alreadySent) {

      sendDefectClosedEmail(row);

      properties.setProperty(mailKey, "YES");
    }

  } else {

    statusCell.setValue("Open");
    closedDateCell.clearContent();
  }
}
function sendDefectClosedEmail(row) {

  var sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("Defect Action Tracker");

  var email = "yuvrajmadavi9@gmail.com";

  var defectID = sheet.getRange(row, 1).getValue();
  var inspectionDate = sheet.getRange(row, 2).getValue();
  var assetNo = sheet.getRange(row, 3).getValue();
  var location = sheet.getRange(row, 4).getValue();
  var extinguisherType = sheet.getRange(row, 5).getValue();
  var capacity = sheet.getRange(row, 6).getValue();
  var defect = sheet.getRange(row, 7).getValue();
  var status = sheet.getRange(row, 10).getValue();
  var dueDate = sheet.getRange(row, 11).getValue();
  var actionTaken = sheet.getRange(row, 12).getValue();
  var closedDate = sheet.getRange(row, 13).getValue();

  var subject = "✅ FIRE SAFETY DEFECT CLOSED - " + defectID;

  var htmlTable =
    '<table border="1" cellpadding="7" cellspacing="0" ' +
    'style="border-collapse:collapse;font-family:Arial;font-size:12px;">';

  htmlTable += '<tr>';
  htmlTable += '<th style="background:#2e7d32;color:white;">Defect ID</th>';
  htmlTable += '<th style="background:#2e7d32;color:white;">Inspection Date</th>';
  htmlTable += '<th style="background:#2e7d32;color:white;">Asset No.</th>';
  htmlTable += '<th style="background:#2e7d32;color:white;">Location</th>';
  htmlTable += '<th style="background:#2e7d32;color:white;">Extinguisher Type</th>';
  htmlTable += '<th style="background:#2e7d32;color:white;">Capacity</th>';
  htmlTable += '<th style="background:#2e7d32;color:white;">Defect</th>';
  htmlTable += '<th style="background:#2e7d32;color:white;">Status</th>';
  htmlTable += '<th style="background:#2e7d32;color:white;">Due Date</th>';
  htmlTable += '<th style="background:#2e7d32;color:white;">Action Taken</th>';
  htmlTable += '<th style="background:#2e7d32;color:white;">Closed Date</th>';
  htmlTable += '</tr>';

  htmlTable += '<tr>';
  htmlTable += '<td>' + defectID + '</td>';
  htmlTable += '<td>' + inspectionDate + '</td>';
  htmlTable += '<td>' + assetNo + '</td>';
  htmlTable += '<td>' + location + '</td>';
  htmlTable += '<td>' + extinguisherType + '</td>';
  htmlTable += '<td>' + capacity + '</td>';
  htmlTable += '<td>' + defect + '</td>';
  htmlTable += '<td style="color:#2e7d32;font-weight:bold;">CLOSED</td>';
  htmlTable += '<td>' + dueDate + '</td>';
  htmlTable += '<td>' + actionTaken + '</td>';
  htmlTable += '<td>' + closedDate + '</td>';
  htmlTable += '</tr>';

  htmlTable += '</table>';

  var htmlBody =
    '<h2>✅ FIRE SAFETY DEFECT CLOSED</h2>' +
    '<p><b>Defect has been closed successfully.</b></p>' +
    htmlTable +
    '<br>' +
    '<p><b>Corrective action has been recorded in the Defect Action Tracker.</b></p>';

  var plainBody =
    "FIRE SAFETY DEFECT CLOSED\n\n" +
    "Defect ID: " + defectID + "\n" +
    "Inspection Date: " + inspectionDate + "\n" +
    "Asset No.: " + assetNo + "\n" +
    "Location: " + location + "\n" +
    "Extinguisher Type: " + extinguisherType + "\n" +
    "Capacity: " + capacity + "\n" +
    "Defect: " + defect + "\n" +
    "Status: CLOSED\n" +
    "Due Date: " + dueDate + "\n" +
    "Action Taken: " + actionTaken + "\n" +
    "Closed Date: " + closedDate + "\n\n" +
    "Corrective action has been recorded in the Defect Action Tracker.";

  MailApp.sendEmail({
    to: email,
    subject: subject,
    body: plainBody,
    htmlBody: htmlBody
  });

}
function checkDefectOverdueNotifications() {

  var sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("Defect Action Tracker");

  var lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return;
  }

  var data = sheet.getRange(2, 1, lastRow - 1, 13).getValues();

  var today = new Date();
  today.setHours(0, 0, 0, 0);

  var properties = PropertiesService.getScriptProperties();

  for (var i = 0; i < data.length; i++) {

    var row = i + 2;

    var defectID = data[i][0];
    var status = data[i][9];
    var dueDate = data[i][10];

    if (!defectID || !dueDate) {
      continue;
    }

    if (status !== "Open") {
      continue;
    }

    var due = new Date(dueDate);
    due.setHours(0, 0, 0, 0);

    if (due >= today) {
      continue;
    }

    var mailKey = "OVERDUE_MAIL_SENT_" + defectID;

    var alreadySent = properties.getProperty(mailKey);

    if (alreadySent) {
      continue;
    }

    sendDefectOverdueEmail(row);

    properties.setProperty(mailKey, "YES");
  }
}
function sendDefectOverdueEmail(row) {

  var sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("Defect Action Tracker");

  var toEmail = "jrg@lloyds.in";

  var ccEmail =
    "sdz@lloyds.in," +
    "yuvrajmadavi9@gmail.com," +
    "sheikhamaan611@gmail.com," +
    "bhoyarmukteshwar123@gmail.com," +
    "pravinkawde209@gmail.com," +
    "sanjaysidam17@gmail.com," +
    "prasaddaradmare99@gmail.com";

  var defectID = sheet.getRange(row, 1).getValue();
  var inspectionDate = sheet.getRange(row, 2).getValue();
  var assetNo = sheet.getRange(row, 3).getValue();
  var location = sheet.getRange(row, 4).getValue();
  var extinguisherType = sheet.getRange(row, 5).getValue();
  var capacity = sheet.getRange(row, 6).getValue();
  var defect = sheet.getRange(row, 7).getValue();
  var status = sheet.getRange(row, 10).getValue();
  var dueDate = sheet.getRange(row, 11).getValue();

  var subject = "⚠️ FIRE SAFETY DEFECT OVERDUE - " + defectID;

  var htmlTable =
    '<table border="1" cellpadding="7" cellspacing="0" ' +
    'style="border-collapse:collapse;font-family:Arial;font-size:12px;">';

  htmlTable += '<tr>';
  htmlTable += '<th style="background:#e65100;color:white;">Defect ID</th>';
  htmlTable += '<th style="background:#e65100;color:white;">Inspection Date</th>';
  htmlTable += '<th style="background:#e65100;color:white;">Asset No.</th>';
  htmlTable += '<th style="background:#e65100;color:white;">Location</th>';
  htmlTable += '<th style="background:#e65100;color:white;">Extinguisher Type</th>';
  htmlTable += '<th style="background:#e65100;color:white;">Capacity</th>';
  htmlTable += '<th style="background:#e65100;color:white;">Defect</th>';
  htmlTable += '<th style="background:#e65100;color:white;">Status</th>';
  htmlTable += '<th style="background:#e65100;color:white;">Due Date</th>';
  htmlTable += '</tr>';

  htmlTable += '<tr>';
  htmlTable += '<td>' + defectID + '</td>';
  htmlTable += '<td>' + inspectionDate + '</td>';
  htmlTable += '<td>' + assetNo + '</td>';
  htmlTable += '<td>' + location + '</td>';
  htmlTable += '<td>' + extinguisherType + '</td>';
  htmlTable += '<td>' + capacity + '</td>';
  htmlTable += '<td>' + defect + '</td>';
  htmlTable += '<td style="color:#e65100;font-weight:bold;">OVERDUE</td>';
  htmlTable += '<td>' + dueDate + '</td>';
  htmlTable += '</tr>';

  htmlTable += '</table>';

  var htmlBody =
    '<h2>⚠️ FIRE SAFETY DEFECT OVERDUE</h2>' +
    '<p><b>Corrective action is overdue.</b></p>' +
    htmlTable +
    '<br>' +
    '<p><b>Please complete the corrective action at the earliest.</b></p>';

  var plainBody =
    "FIRE SAFETY DEFECT OVERDUE\n\n" +
    "Defect ID: " + defectID + "\n" +
    "Inspection Date: " + inspectionDate + "\n" +
    "Asset No.: " + assetNo + "\n" +
    "Location: " + location + "\n" +
    "Extinguisher Type: " + extinguisherType + "\n" +
    "Capacity: " + capacity + "\n" +
    "Defect: " + defect + "\n" +
    "Status: OVERDUE\n" +
    "Due Date: " + dueDate + "\n\n" +
    "Please complete the corrective action at the earliest.";

  MailApp.sendEmail({
    to: toEmail,
    cc: ccEmail,
    subject: subject,
    body: plainBody,
    htmlBody: htmlBody
  });
}
function sendDefectEmail(defectID, inspectionDate, assetNo, location, extinguisherType, capacity, defect, dueDate) {

  var email = "yuvrajmadavi9@gmail.com";

  var subject = "🔥 FIRE SAFETY DEFECT ALERT - " + defectID;

  var htmlTable =
    '<table border="1" cellpadding="7" cellspacing="0" ' +
    'style="border-collapse:collapse;font-family:Arial;font-size:12px;">';

  htmlTable += '<tr>';
  htmlTable += '<th style="background:#d32f2f;color:white;">Defect ID</th>';
  htmlTable += '<th style="background:#d32f2f;color:white;">Inspection Date</th>';
  htmlTable += '<th style="background:#d32f2f;color:white;">Asset No.</th>';
  htmlTable += '<th style="background:#d32f2f;color:white;">Location</th>';
  htmlTable += '<th style="background:#d32f2f;color:white;">Extinguisher Type</th>';
  htmlTable += '<th style="background:#d32f2f;color:white;">Capacity</th>';
  htmlTable += '<th style="background:#d32f2f;color:white;">Defect</th>';
  htmlTable += '<th style="background:#d32f2f;color:white;">Status</th>';
  htmlTable += '<th style="background:#d32f2f;color:white;">Due Date</th>';
  htmlTable += '</tr>';

  htmlTable += '<tr>';
  htmlTable += '<td>' + defectID + '</td>';
  htmlTable += '<td>' + inspectionDate + '</td>';
  htmlTable += '<td>' + assetNo + '</td>';
  htmlTable += '<td>' + location + '</td>';
  htmlTable += '<td>' + extinguisherType + '</td>';
  htmlTable += '<td>' + capacity + '</td>';
  htmlTable += '<td>' + defect + '</td>';
  htmlTable += '<td>OPEN</td>';
  htmlTable += '<td>' + dueDate + '</td>';
  htmlTable += '</tr>';

  htmlTable += '</table>';

  var htmlBody =
    '<h2>🔥 FIRE SAFETY DEFECT ALERT</h2>' +
    '<p><b>New Defect Identified</b></p>' +
    htmlTable +
    '<br><p><b>Please take necessary corrective action.</b></p>';

  var plainBody =
    "FIRE SAFETY DEFECT ALERT\n\n" +
    "Defect ID: " + defectID + "\n" +
    "Inspection Date: " + inspectionDate + "\n" +
    "Asset No.: " + assetNo + "\n" +
    "Location: " + location + "\n" +
    "Extinguisher Type: " + extinguisherType + "\n" +
    "Capacity: " + capacity + "\n" +
    "Defect: " + defect + "\n" +
    "Status: OPEN\n" +
    "Due Date: " + dueDate + "\n\n" +
    "Please take necessary corrective action.";

  MailApp.sendEmail({
    to: email,
    subject: subject,
    body: plainBody,
    htmlBody: htmlBody
  });
}
function testEmailPermission() {
  MailApp.sendEmail(
    "yuvrajmadavi9@gmail.com",
    "Fire Safety Automation - Email Test",
    "Email permission test successful."
  );
}
function sendResponseNotification(e) {

  var recipients = "sdz@lloyds.in,yuvrajmadavi9@gmail.com,sheikhamaan611@gmail.com,jrg@lloyds.in,hra@lloyds.in,bhoyarmukteshwar123@gmail.com,pravinkawde209@gmail.com,sanjaysidam17@gmail.com,prasaddaradmare99@gmail.com";

  var subject = "🔥 Fire Extinguisher Monthly Inspection - New Response";

  var sheet = e.range.getSheet();
  var row = e.range.getRow();
  var lastColumn = sheet.getLastColumn();

  // Get headers and response data from the same row
  var headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  var values = sheet.getRange(row, 1, 1, lastColumn).getValues()[0];

  var htmlTable =
    '<table border="1" cellpadding="6" cellspacing="0" ' +
    'style="border-collapse:collapse;font-family:Arial;font-size:12px;">';

  // Header row
  htmlTable += '<tr>';

  for (var i = 0; i < headers.length; i++) {

    // Hide Zone from email
    if (String(headers[i]).trim().toLowerCase().indexOf("zone wise inspection") !== -1) {
      continue;
    }

    htmlTable +=
      '<th style="background:#4b3cc4;color:white;white-space:nowrap;">' +
      headers[i] +
      '</th>';
  }

  htmlTable += '</tr>';

  // Data row
  htmlTable += '<tr>';

  for (var j = 0; j < headers.length; j++) {

    // Hide Zone from email
    if (String(headers[j]).trim().toLowerCase().indexOf("zone wise inspection") !== -1) {
      continue;
    }

    var value = values[j];

    // Make Inspection Photo clickable
    if (
      headers[j].toLowerCase().indexOf("photo") !== -1 &&
      String(value).indexOf("http") === 0
    ) {
      value = '<a href="' + value + '" target="_blank">View Photo</a>';
    }

    htmlTable +=
      '<td style="white-space:nowrap;">' +
      value +
      '</td>';
  }

  htmlTable += '</tr>';
  htmlTable += '</table>';

  var htmlBody =
    '<h2>🔥 FIRE EXTINGUISHER MONTHLY INSPECTION</h2>' +
    '<p><b>New Response Received</b></p>' +
    htmlTable +
    '<br><p>This is an automatic notification from Fire Safety Automation.</p>';

  var plainBody =
    "Fire Extinguisher Monthly Inspection - New Response\n\n";

  for (var k = 0; k < headers.length; k++) {

    if (String(headers[k]).trim().toLowerCase().indexOf("zone wise inspection") !== -1) {
      continue;
    }

    plainBody += headers[k] + ": " + values[k] + "\n";
  }

  MailApp.sendEmail({
    to: recipients,
    subject: subject,
    body: plainBody,
    htmlBody: htmlBody
  });
}
function sendDailyInspectionSummary() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var inspectionSheet =
    ss.getSheetByName("Fire Extinguisher Monthly Inspection");

  var defectSheet =
    ss.getSheetByName("Defect Action Tracker");

  if (!inspectionSheet || !defectSheet) {
    Logger.log("Inspection Sheet or Defect Action Tracker not found.");
    return;
  }

  var toEmail = "jrg@lloyds.in";

  var ccEmail =
    "sdz@lloyds.in," +
    "yuvrajmadavi9@gmail.com," +
    "sheikhamaan611@gmail.com," +
    "hra@lloyds.in," +
    "bhoyarmukteshwar123@gmail.com," +
    "pravinkawde209@gmail.com," +
    "sanjaysidam17@gmail.com," +
    "prasaddaradmare99@gmail.com";

  var timezone = ss.getSpreadsheetTimeZone();

  var yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);

  var reportDate = Utilities.formatDate(
    yesterday,
    timezone,
    "yyyy-MM-dd"
  );

  var reportDateDisplay = Utilities.formatDate(
    yesterday,
    timezone,
    "dd-MM-yyyy"
  );
    // ==============================
  // GET PREVIOUS DAY INSPECTIONS
  // ==============================

  var inspectionLastRow = inspectionSheet.getLastRow();
  var inspectionLastColumn = inspectionSheet.getLastColumn();

  var inspectionHeaders =
    inspectionSheet
      .getRange(1, 1, 1, inspectionLastColumn)
      .getValues()[0];

  var inspectionData = [];

  if (inspectionLastRow > 1) {
    inspectionData =
      inspectionSheet
        .getRange(
          2,
          1,
          inspectionLastRow - 1,
          inspectionLastColumn
        )
        .getValues();
  }

  var yesterdayInspections = [];

  for (var i = 0; i < inspectionData.length; i++) {

    var timestamp = inspectionData[i][0];

    if (!(timestamp instanceof Date)) {
      continue;
    }

    var rowDate = Utilities.formatDate(
      timestamp,
      timezone,
      "yyyy-MM-dd"
    );

    if (rowDate === reportDate) {
      yesterdayInspections.push(inspectionData[i]);
    }
  }

  // ==============================
  // GET PREVIOUS DAY DEFECTS
  // ==============================

  var defectLastRow = defectSheet.getLastRow();

  var defectData = [];

  if (defectLastRow > 1) {
    defectData =
      defectSheet
        .getRange(
          2,
          1,
          defectLastRow - 1,
          13
        )
        .getValues();
  }

  var yesterdayDefects = [];

  for (var d = 0; d < defectData.length; d++) {

    var defectDate = defectData[d][1];

    if (!(defectDate instanceof Date)) {
      continue;
    }

    var defectRowDate = Utilities.formatDate(
      defectDate,
      timezone,
      "yyyy-MM-dd"
    );

    if (defectRowDate === reportDate) {
      yesterdayDefects.push(defectData[d]);
    }
  }
    // ==============================
  // INSPECTION SUMMARY
  // ==============================

  var inspectionSummary = {};

  for (var i = 0; i < yesterdayInspections.length; i++) {

    var type =
      String(yesterdayInspections[i][4]).trim();

    var capacity =
      String(yesterdayInspections[i][5]).trim();

    var key =
      type + "||" + capacity;

    if (!inspectionSummary[key]) {

      inspectionSummary[key] = {
        type: type,
        capacity: capacity,
        total: 0,
        defects: 0
      };
    }

    inspectionSummary[key].total++;
  }

  // ==============================
  // COUNT DEFECTS BY TYPE + CAPACITY
  // ==============================

  for (var d = 0; d < yesterdayDefects.length; d++) {

    var defectType =
      String(yesterdayDefects[d][4]).trim();

    var defectCapacity =
      String(yesterdayDefects[d][5]).trim();

    var defectKey =
      defectType + "||" + defectCapacity;

    if (!inspectionSummary[defectKey]) {

      inspectionSummary[defectKey] = {
        type: defectType,
        capacity: defectCapacity,
        total: 0,
        defects: 0
      };
    }

    inspectionSummary[defectKey].defects++;
  }

  // ==============================
  // INSPECTION SUMMARY TABLE
  // ==============================

  var inspectionSummaryTable =
    '<table border="1" cellpadding="7" cellspacing="0" ' +
    'style="border-collapse:collapse;font-family:Arial;font-size:12px;">';

  inspectionSummaryTable += '<tr>';

  inspectionSummaryTable +=
    '<th style="background:#4b3cc4;color:white;">Extinguisher Type</th>';

  inspectionSummaryTable +=
    '<th style="background:#4b3cc4;color:white;">Capacity</th>';

  inspectionSummaryTable +=
    '<th style="background:#4b3cc4;color:white;">Total Inspected</th>';

  inspectionSummaryTable +=
    '<th style="background:#d32f2f;color:white;">Defective Extinguisher Qty</th>';
  inspectionSummaryTable +=
    '<th style="background:#4b3cc4;color:white;">OK</th>';

  inspectionSummaryTable += '</tr>';
    // ==============================
  // INSPECTION SUMMARY ROWS
  // ==============================

  for (var key in inspectionSummary) {

    var total = inspectionSummary[key].total;
    var defects = inspectionSummary[key].defects;
    var ok = total - defects;

    inspectionSummaryTable += '<tr>';

    inspectionSummaryTable +=
      '<td>' + inspectionSummary[key].type + '</td>';

    inspectionSummaryTable +=
      '<td>' + inspectionSummary[key].capacity + '</td>';

    inspectionSummaryTable +=
      '<td style="text-align:center;">' +
      total +
      '</td>';

    inspectionSummaryTable +=
      '<td style="text-align:center;background:#ffebee;color:#c62828;font-weight:bold;">' +
      defects +
      '</td>';

    inspectionSummaryTable +=
      '<td style="text-align:center;">' +
      ok +
      '</td>';

    inspectionSummaryTable += '</tr>';
  }

  inspectionSummaryTable += '<tr>';

  inspectionSummaryTable +=
    '<td colspan="2"><b>TOTAL INSPECTIONS</b></td>';

  inspectionSummaryTable +=
    '<td style="text-align:center;"><b>' +
    yesterdayInspections.length +
    '</b></td>';

  inspectionSummaryTable +=
    '<td style="text-align:center;background:#ffebee;color:#c62828;font-weight:bold;">' +
    yesterdayDefects.length +
    '</td>';

  inspectionSummaryTable +=
    '<td style="text-align:center;"><b>' +
    (yesterdayInspections.length - yesterdayDefects.length) +
    '</b></td>';

  inspectionSummaryTable += '</tr>';

  inspectionSummaryTable += '</table>';

  // ==============================
  // DEFECT SUMMARY TABLE
  // ==============================

  var defectSummary = {};

  for (var d = 0; d < yesterdayDefects.length; d++) {

    var type =
      String(yesterdayDefects[d][4]).trim();

    var capacity =
      String(yesterdayDefects[d][5]).trim();

    var key =
      type + "||" + capacity;

    if (!defectSummary[key]) {

      defectSummary[key] = {
        type: type,
        capacity: capacity,
        qty: 0
      };
    }

    defectSummary[key].qty++;
  }

  var defectSummaryTable =
    '<table border="1" cellpadding="7" cellspacing="0" ' +
    'style="border-collapse:collapse;font-family:Arial;font-size:12px;">';

  defectSummaryTable += '<tr>';

  defectSummaryTable +=
    '<th style="background:#d32f2f;color:white;">Extinguisher Type</th>';

  defectSummaryTable +=
    '<th style="background:#d32f2f;color:white;">Capacity</th>';

  defectSummaryTable +=
    '<th style="background:#d32f2f;color:white;">Defective Extinguisher Qty</th>';

  defectSummaryTable += '</tr>';
    // ==============================
  // DEFECT SUMMARY ROWS
  // ==============================

  for (var key in defectSummary) {

    defectSummaryTable += '<tr>';

    defectSummaryTable +=
      '<td>' + defectSummary[key].type + '</td>';

    defectSummaryTable +=
      '<td>' + defectSummary[key].capacity + '</td>';

    defectSummaryTable +=
      '<td style="text-align:center;background:#ffebee;color:#c62828;font-weight:bold;">' +
      defectSummary[key].qty +
      '</td>';

    defectSummaryTable += '</tr>';
  }

  defectSummaryTable += '<tr>';

  defectSummaryTable +=
    '<td colspan="2"><b>TOTAL DEFECTIVE EXTINGUISHERS</b></td>';
  defectSummaryTable +=
    '<td style="text-align:center;background:#ffebee;color:#c62828;font-weight:bold;">' +
    yesterdayDefects.length +
    '</td>';

  defectSummaryTable += '</tr>';

  defectSummaryTable += '</table>';
    // ==============================
  // DEFECT DETAILS TABLE
  // ==============================

  var defectTable =
    '<table border="1" cellpadding="6" cellspacing="0" ' +
    'style="border-collapse:collapse;font-family:Arial;font-size:11px;">';

  defectTable += '<tr>';

  var defectHeaders = [
    "Defect ID",
    "Inspection Date",
    "Asset No.",
    "Location",
    "Extinguisher Type",
    "Capacity",
    "Defect",
    "Remarks",
    "Checked By",
    "Status",
    "Due Date",
    "Action Taken",
    "Closed Date"
  ];

  for (var h = 0; h < defectHeaders.length; h++) {

    defectTable +=
      '<th style="background:#d32f2f;color:white;white-space:nowrap;">' +
      defectHeaders[h] +
      '</th>';
  }

  defectTable += '</tr>';

  for (var d = 0; d < yesterdayDefects.length; d++) {

    defectTable += '<tr>';

    for (var c = 0; c < 13; c++) {

      var value = yesterdayDefects[d][c];

      if (value instanceof Date) {

        value = Utilities.formatDate(
          value,
          timezone,
          "dd-MM-yyyy HH:mm"
        );
      }

      defectTable +=
        '<td style="white-space:nowrap;">' +
        value +
        '</td>';
    }

    defectTable += '</tr>';
  }

  defectTable += '</table>';
    // ==============================
  // COMPLETE INSPECTION TABLE
  // ==============================

  var htmlTable =
    '<table border="1" cellpadding="6" cellspacing="0" ' +
    'style="border-collapse:collapse;font-family:Arial;font-size:11px;">';

  htmlTable += '<tr>';

  for (var h = 0; h < inspectionHeaders.length; h++) {

    if (
      String(inspectionHeaders[h])
        .trim()
        .toLowerCase()
        .indexOf("zone wise inspection") !== -1
    ) {
      continue;
    }

    htmlTable +=
      '<th style="background:#4b3cc4;color:white;white-space:nowrap;">' +
      inspectionHeaders[h] +
      '</th>';
  }

  htmlTable += '</tr>';
    // ==============================
  // INSPECTION DATA ROWS
  // ==============================

  for (var r = 0; r < yesterdayInspections.length; r++) {

    htmlTable += '<tr>';

    for (var c = 0; c < inspectionHeaders.length; c++) {

      if (
        String(inspectionHeaders[c])
          .trim()
          .toLowerCase()
          .indexOf("zone wise inspection") !== -1
      ) {
        continue;
      }

      var value = yesterdayInspections[r][c];

      if (value instanceof Date) {
        value = Utilities.formatDate(
          value,
          timezone,
          "dd-MM-yyyy HH:mm"
        );
      }

      if (
        String(inspectionHeaders[c])
          .toLowerCase()
          .indexOf("photo") !== -1 &&
        String(value).indexOf("http") === 0
      ) {
        value =
          '<a href="' +
          value +
          '" target="_blank">View Photo</a>';
      }

      htmlTable +=
        '<td style="white-space:nowrap;">' +
        value +
        '</td>';
    }

    htmlTable += '</tr>';
  }

  htmlTable += '</table>';
    // ==============================
  // FINAL DAILY REPORT EMAIL
  // ==============================

  var subject =
    "🔥 Fire Extinguisher Daily Report - " +
    reportDateDisplay;

  var htmlBody =
    '<h2>🔥 FIRE EXTINGUISHER DAILY REPORT</h2>' +

    '<p><b>Date:</b> ' +
    reportDateDisplay +
    '</p>' +

    '<h3 style="color:#4b3cc4;">INSPECTION SUMMARY</h3>' +
    inspectionSummaryTable +

    '<br>' +

    '<h3 style="color:#d32f2f;">🔴 DEFECT SUMMARY</h3>' +
    defectSummaryTable +

    '<br>' +

    '<h3 style="color:#d32f2f;">🔴 DEFECT DETAILS</h3>' +
    defectTable +

    '<br>' +

    '<h3 style="color:#4b3cc4;">COMPLETE INSPECTION DETAILS</h3>' +
    htmlTable +

    '<br>' +

    '<p><b>Total Inspections:</b> ' +
    yesterdayInspections.length +
    '</p>' +

    '<p><b style="color:#d32f2f;">Total Defective Extinguishers:</b> ' +
    yesterdayDefects.length +
    '</p>' +

    '<br>' +

    '<p>This is an automatic daily notification from Fire Safety Automation.</p>';

  var plainBody =
    "Fire Extinguisher Daily Report\n\n" +
    "Date: " +
    reportDateDisplay +
    "\n\n" +
    "Total Inspections: " +
    yesterdayInspections.length +
    "\n" +
    "Total Defective Extinguishers: " +
    yesterdayDefects.length +
    "\n\n" +
    "Please refer to the email for complete inspection and defect details.";

  MailApp.sendEmail({
    to: toEmail,
    cc: ccEmail,
    subject: subject,
    body: plainBody,
    htmlBody: htmlBody
  });

  Logger.log(
    "Daily Fire Extinguisher Report sent. " +
    "Inspections: " +
    yesterdayInspections.length +
    " | Defects: " +
    yesterdayDefects.length
  );

}
function sendMonthlyFireExtinguisherKPI() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var inspectionSheet =
    ss.getSheetByName("Fire Extinguisher Monthly Inspection");

  var defectSheet =
    ss.getSheetByName("Defect Action Tracker");

  if (!inspectionSheet || !defectSheet) {
    return;
  }

  var today = new Date();

var lastDayNumber = new Date(
  today.getFullYear(),
  today.getMonth() + 1,
  0
).getDate();

if (today.getDate() !== lastDayNumber) {
  return;
}

var monthKey = Utilities.formatDate(
  today,
  Session.getScriptTimeZone(),
  "yyyy-MM"
);

var properties = PropertiesService.getScriptProperties();

var mailKey = "MONTHLY_KPI_SENT_" + monthKey;

if (properties.getProperty(mailKey)) {
  return;
}

var monthStart = new Date(
  today.getFullYear(),
  today.getMonth(),
  1
);

var nextMonthStart = new Date(
  today.getFullYear(),
  today.getMonth() + 1,
  1
);
  var monthName =
    Utilities.formatDate(
      today,
      Session.getScriptTimeZone(),
      "MMMM yyyy"
    );

  var inspectionData =
    inspectionSheet.getDataRange().getValues();

  var inspectedCount = 0;
    // ==========================================
  // INSPECTION COMPLIANCE - ASSET MASTER
  // ==========================================

  var assetMasterSheet =
    ss.getSheetByName("Fire Extinguisher Asset Master");

  var activeAssetKeys = {};
  var inspectedAssetKeys = {};
  var totalActiveAssets = 0;

  if (assetMasterSheet) {

    var assetMasterLastRow =
      assetMasterSheet.getLastRow();

    if (assetMasterLastRow >= 2) {

      var assetMasterData =
        assetMasterSheet
          .getRange(2, 1, assetMasterLastRow - 1, 5)
          .getValues();

      for (var a = 0; a < assetMasterData.length; a++) {

        var masterAssetNo =
          String(assetMasterData[a][0]).trim();

        var masterType =
          String(assetMasterData[a][2]).trim().toUpperCase();

        var masterCapacity =
          String(assetMasterData[a][3]).trim().toUpperCase();

        var masterStatus =
          String(assetMasterData[a][4]).trim();

        if (!masterAssetNo || masterStatus !== "Active") {
          continue;
        }

        if (masterType === "ABC (DCP)") {
          masterType = "ABC";
        }

        var masterKey =
          masterAssetNo +
          " | " +
          masterType +
          " | " +
          masterCapacity;

        activeAssetKeys[masterKey] = true;
        totalActiveAssets++;
      }
    }
  }

  


// ==========================================
// COUNT UNIQUE ASSETS INSPECTED THIS MONTH
// ==========================================

for (var i = 1; i < inspectionData.length; i++) {

  var inspectionDate =
    new Date(inspectionData[i][1]);

  if (
    inspectionDate >= monthStart &&
    inspectionDate < nextMonthStart
  ) {

    var inspectionAssetNo =
      String(inspectionData[i][3]).trim();

    var inspectionType =
      String(inspectionData[i][4]).trim().toUpperCase();

    var inspectionCapacity =
      String(inspectionData[i][5]).trim().toUpperCase();

    if (inspectionType === "ABC (DCP)") {
      inspectionType = "ABC";
    }

    var inspectionAssetKey =
      inspectionAssetNo +
      " | " +
      inspectionType +
      " | " +
      inspectionCapacity;

    if (activeAssetKeys[inspectionAssetKey]) {
      inspectedAssetKeys[inspectionAssetKey] = true;
    }
  }
}


// Unique active assets inspected this month
inspectedCount =
  Object.keys(inspectedAssetKeys).length;
  var defectData =
  defectSheet.getDataRange().getValues();
  

  var defectiveCount = 0;
  var closedCount = 0;
  var openCount = 0;
  var overdueCount = 0;

  var todayDate = new Date();
  todayDate.setHours(0, 0, 0, 0);

  for (var j = 1; j < defectData.length; j++) {

    var defectInspectionDate =
      new Date(defectData[j][1]);

    if (
      defectInspectionDate >= monthStart &&
      defectInspectionDate < nextMonthStart
    ) {

      defectiveCount++;

      var status = defectData[j][9];
      var dueDate = defectData[j][10];

      if (status === "Closed") {
        closedCount++;
      }

      if (status === "Open") {
        openCount++;

        if (dueDate) {

          var due = new Date(dueDate);
          due.setHours(0, 0, 0, 0);

          if (due < todayDate) {
            overdueCount++;
          }
        }
      }
    }
  }

  var defectRate = 0;

  if (inspectedCount > 0) {
    defectRate =
      (defectiveCount / inspectedCount) * 100;
  }

  defectRate = defectRate.toFixed(1);

  var toEmail = "jrg@lloyds.in";

  var ccEmail =
    "sdz@lloyds.in," +
    "yuvrajmadavi9@gmail.com," +
    "sheikhamaan611@gmail.com," +
    "bhoyarmukteshwar123@gmail.com," +
    "pravinkawde209@gmail.com," +
    "sanjaysidam17@gmail.com," +
    "prasaddaradmare99@gmail.com," +
    "hra@lloyds.in";

  var subject =
    "📊 Fire Extinguisher Monthly KPI - " +
    monthName;

  var htmlBody =
    "<h2>📊 FIRE EXTINGUISHER MONTHLY KPI</h2>" +
    "<p><b>Reporting Month:</b> " +
    monthName +
    "</p>" +

    "<table border='1' cellpadding='8' cellspacing='0' " +
    "style='border-collapse:collapse;font-family:Arial;font-size:13px;'>" +

    "<tr>" +
    "<th>Parameter</th>" +
    "<th>Result</th>" +
    "</tr>" +

    "<tr><td>Total Inspected</td><td><b>" +
    inspectedCount +
    "</b></td></tr>" +

    "<tr><td>Defective Extinguishers</td><td><b>" +
    defectiveCount +
    "</b></td></tr>" +

    "<tr><td>Defect Rate</td><td><b>" +
    defectRate +
    "%</b></td></tr>" +

    "<tr><td>Closed Defects</td><td><b>" +
    closedCount +
    "</b></td></tr>" +

    "<tr><td>Open Defects</td><td><b>" +
    openCount +
    "</b></td></tr>" +

    "<tr><td>Overdue Defects</td><td><b>" +
    overdueCount +
    "</b></td></tr>" +

    "</table>" +

    "<br>" +
    "<p><b>This is an automatic Fire Safety KPI report.</b></p>";

  var plainBody =
    "FIRE EXTINGUISHER MONTHLY KPI\n\n" +
    "Reporting Month: " + monthName + "\n" +
    "Total Inspected: " + inspectedCount + "\n" +
    "Defective Extinguishers: " + defectiveCount + "\n" +
    "Defect Rate: " + defectRate + "%\n" +
    "Closed Defects: " + closedCount + "\n" +
    "Open Defects: " + openCount + "\n" +
    "Overdue Defects: " + overdueCount;

  MailApp.sendEmail({
    to: toEmail,
    cc: ccEmail,
    subject: subject,
    body: plainBody,
    htmlBody: htmlBody
  });
  properties.setProperty(mailKey, "YES");
}
// ==========================================
// FIRE INSPECTION APP - WEB APP BASE
// ==========================================

function doGet(e) {

  // GitHub se Asset List request
  if (e && e.parameter && e.parameter.action === "assets") {

    var assets = getInspectionAssets();

    var callback = e.parameter.callback;

    if (callback) {

      var output =
        callback +
        "(" +
        JSON.stringify(assets) +
        ");";

      return ContentService
        .createTextOutput(output)
        .setMimeType(ContentService.MimeType.JAVASCRIPT);

    }

    return ContentService
      .createTextOutput(JSON.stringify(assets))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Original Apps Script app
  return HtmlService
    .createHtmlOutputFromFile("Index")
    .setTitle("Fire Extinguisher Inspection");
}
// ==========================================
// FIRE INSPECTION APP - GET ASSET DATA
// ==========================================

// ==========================================
// FIRE INSPECTION APP - ASSET LIST
// ==========================================

// ==========================================
// FIRE INSPECTION APP - ASSET LIST
// ==========================================

// ==========================================
// FIRE INSPECTION APP - ASSET LIST
// ==========================================

// ==========================================
// FIRE INSPECTION APP - ASSET LIST
// AVAILABLE + NOT AVAILABLE + DEFECTIVE
// ==========================================

function getInspectionAssets() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var sheet =
    ss.getSheetByName("Fire Extinguisher Asset Master");

  if (!sheet) {
    throw new Error(
      "Fire Extinguisher Asset Master sheet not found."
    );
  }

  var lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return [];
  }

  // A:H
  // A = Asset No.
  // B = Location
  // C = Extinguisher Type
  // D = Capacity
  // E = Status
  // F = Extinguisher ID
  // G = Inspection Status
  // H = Status Reason

  var data =
    sheet
      .getRange(2, 1, lastRow - 1, 8)
      .getValues();

  var assets = [];

  for (var i = 0; i < data.length; i++) {

    var assetNo =
      String(data[i][0]).trim();

    var location =
      String(data[i][1]).trim();

    var type =
      String(data[i][2]).trim();

    var capacity =
      String(data[i][3]).trim();

    var status =
      String(data[i][4]).trim();

    var extinguisherId =
      String(data[i][5]).trim();

    var inspectionStatus =
      String(data[i][6]).trim();

    var statusReason =
      String(data[i][7]).trim();

    if (!assetNo) {
      continue;
    }

    if (!type || !capacity || !extinguisherId) {
      continue;
    }

    // Blank status will be treated as Available
    if (!inspectionStatus) {
      inspectionStatus = "Available";
    }

    assets.push({

      assetNo: assetNo,

      extinguisherId: extinguisherId,

      location: location,

      type: type,

      capacity: capacity,

      status: status,

      inspectionStatus: inspectionStatus,

      statusReason: statusReason

    });

  }

  return assets;
}

// ==========================================
// FIRE INSPECTION APP - INITIAL ASSET MASTER SETUP
// ==========================================

function setupFireExtinguisherAssetMaster() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var inspectionSheet =
    ss.getSheetByName("Fire Extinguisher Monthly Inspection");

  var masterSheet =
    ss.getSheetByName("Fire Extinguisher Asset Master");

  if (!inspectionSheet) {
    throw new Error("Fire Extinguisher Monthly Inspection sheet not found.");
  }

  if (!masterSheet) {
    throw new Error("Fire Extinguisher Asset Master sheet not found.");
  }

  var lastRow = inspectionSheet.getLastRow();

  if (lastRow < 2) {
    Logger.log("No inspection data found.");
    return;
  }

  var data =
    inspectionSheet
      .getRange(2, 1, lastRow - 1, 17)
      .getValues();

  var existing = {};

  var masterLastRow = masterSheet.getLastRow();

  if (masterLastRow >= 2) {

    var masterData =
      masterSheet
        .getRange(2, 1, masterLastRow - 1, 5)
        .getValues();

    for (var i = 0; i < masterData.length; i++) {

      var existingAsset =
        String(masterData[i][0]).trim().toUpperCase();

      var existingType =
        String(masterData[i][2]).trim().toUpperCase();

      var existingCapacity =
        String(masterData[i][3]).trim().toUpperCase();

      if (existingAsset && existingType && existingCapacity) {

        if (existingType === "ABC (DCP)") {
          existingType = "ABC";
        }

        var existingKey =
          existingAsset +
          "|" +
          existingType +
          "|" +
          existingCapacity;

        existing[existingKey] = true;
      }
    }
  }

  var newRows = [];

  for (var j = 0; j < data.length; j++) {

    var assetNo =
      String(data[j][3]).trim();

    var location =
      String(data[j][2]).trim();

    var type =
      String(data[j][4]).trim();

    var capacity =
      String(data[j][5]).trim();

    if (
      !assetNo ||
      assetNo.toLowerCase() === "no asset number"
    ) {
      continue;
    }

    if (!type || !capacity) {
      continue;
    }

    var normalizedType =
      type.toUpperCase();

    if (normalizedType === "ABC (DCP)") {
      normalizedType = "ABC";
    }

    var key =
      assetNo.toUpperCase() +
      "|" +
      normalizedType +
      "|" +
      capacity.toUpperCase();

    if (existing[key]) {
      continue;
    }

    newRows.push([
      assetNo,
      location,
      type,
      capacity,
      "Active"
    ]);

    existing[key] = true;
  }

  if (newRows.length > 0) {

    masterSheet
      .getRange(
        masterSheet.getLastRow() + 1,
        1,
        newRows.length,
        5
      )
      .setValues(newRows);
  }

  Logger.log(
    "Asset Master setup completed. New assets added: " +
    newRows.length
  );
}
function saveInspectionFromWebApp(data) {

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var sheet = ss.getSheetByName(
    "Fire Extinguisher Monthly Inspection"
  );

  if (!sheet) {
    throw new Error(
      "Fire Extinguisher Monthly Inspection sheet not found."
    );
  }

  // ==========================================
  // GET ASSET DETAILS
  // ==========================================

  var selectedAsset = String(data.assetNo || "").trim();

  if (!selectedAsset) {
    throw new Error("Please select an Asset No.");
  }

  var parts = selectedAsset.split(" | ");

  var assetNo = String(parts[0] || "").trim();

  if (!assetNo) {
    throw new Error("Invalid Asset No.");
  }
    // ==========================================
  // DUPLICATE INSPECTION ID PROTECTION
  // ==========================================

  var inspectionId =
    String(data.inspectionId || "").trim();

  if (!inspectionId) {
    throw new Error("Inspection ID missing.");
  }

  var lastRow = sheet.getLastRow();

  if (lastRow >= 2) {

    var idRange =
      sheet.getRange(
        2,
        19,
        lastRow - 1,
        1
      );

    var existingId =
      idRange
        .createTextFinder(inspectionId)
        .matchEntireCell(true)
        .findNext();

    if (existingId) {

      return (
        "Inspection already saved: " +
        inspectionId
      );

    }
  }

  // ==========================================
  // USE AUTO-FILLED DETAILS
  // ==========================================

  var location =
    String(data.location || "").trim();

  var type =
    String(data.extinguisherType || "").trim();

  var capacity =
    String(data.capacity || "").trim();

  // ==========================================
  // UPLOAD PHOTO TO GOOGLE DRIVE
  // ==========================================

  var photoUrl = "";

  if (data.photo) {

    try {

      var folderName = "Fire Inspection Photos";

      var folders = DriveApp.getFoldersByName(folderName);

      var folder;

      if (folders.hasNext()) {

        folder = folders.next();

      } else {

        folder = DriveApp.createFolder(folderName);

      }

      var photoData = String(data.photo);

      var parts = photoData.split(",");

      if (parts.length < 2) {
        throw new Error("Invalid photo data.");
      }

      var mimeType =
        photoData.match(/data:(.*?);base64/)[1];

      var decoded =
        Utilities.base64Decode(parts[1]);

      var blob =
        Utilities.newBlob(
          decoded,
          mimeType,
          "Inspection_" +
          assetNo +
          "_" +
          new Date().getTime() +
          ".jpg"
        );

      var file =
        folder.createFile(blob);

      file.setSharing(
        DriveApp.Access.ANYONE_WITH_LINK,
        DriveApp.Permission.VIEW
      );

      photoUrl = file.getUrl();

    } catch (photoError) {

      throw new Error(
        "Photo upload failed: " +
        photoError.message
      );

    }
  }
  // ==========================================
// CALCULATE CURRENT INSPECTION STATUS
// BEFORE SAVING MONTHLY INSPECTION
// ==========================================

var currentParameters = [

  "Safety Pin",
  "Discharge Hose",
  "Discharge Nozzle / Horn",
  "Pressure / Weight",
  "Cap Assembly",
  "Inspection Sticker"

];

var currentValues = [

  data.safetyPin,
  data.dischargeHose,
  data.dischargeNozzle,
  data.pressureWeight,
  data.capAssembly,
  data.inspectionSticker

];

var currentDefects = [];

for (var cs = 0; cs < currentValues.length; cs++) {

  var currentValue =
    String(currentValues[cs] || "").trim();

  if (
    currentValue &&
    currentValue.toUpperCase() !== "OK"
  ) {

    currentDefects.push(
      currentParameters[cs] +
      " = " +
      currentValue
    );

  }

}

// ==========================================
// CURRENT INSPECTION RESULT
// ==========================================

if (currentDefects.length > 0) {

  data.inspectionStatus = "Defective";

  data.statusReason =
    currentDefects.join(" | ");

}

// ==========================================
// MANUAL NOT AVAILABLE
// ==========================================

else if (
  String(data.inspectionStatus || "").trim() ===
  "Not Available"
) {

  data.inspectionStatus = "Not Available";

  data.statusReason =
    String(data.statusReason || "").trim();

}

// ==========================================
// ALL SIX POINTS OK
// ==========================================

else {

  data.inspectionStatus = "Available";

  data.statusReason = "";

}

  // ==========================================
  // SAVE INSPECTION
  // ==========================================

 sheet.appendRow([

  new Date(),                    // A - Timestamp
  data.inspectionDate,           // B - Inspection Date
  location,                      // C - Location
  assetNo,                       // D - Asset No.
  type,                          // E - Extinguisher Type
  capacity,                      // F - Capacity

  data.safetyPin,                // G - Safety Pin
  data.dischargeHose,            // H - Discharge Hose
  data.dischargeNozzle,          // I - Discharge Nozzle / Horn
  data.pressureWeight,           // J - Pressure / Weight
  data.capAssembly,              // K - Cap Assembly
  data.inspectionSticker,        // L - Inspection Sticker

  data.mfgYear,                  // M - Mfg. Year
  data.hydroTestDueYear,         // N - Hydro Test Due Year
  photoUrl,                      // O - Inspection Photo
  data.remarks,                  // P - Remarks
  data.checkedBy,                // Q - Checked By
  "",                            // R - Zone Wise Inspection
  data.inspectionId,             // S - Inspection ID
  data.extinguisherId,           // T - Extinguisher ID
  data.inspectionStatus,         // U - Inspection Status
  data.statusReason              // V - Status / Defect Reason

]);

// ==========================================
  // RUN EXISTING INSPECTION AUTOMATION
  // ==========================================

  var newRow = sheet.getLastRow();

  handleExtinguisherInspection({
  range: sheet.getRange(newRow, 1)
});

// AUTO UPDATE ASSET MASTER
setupFireExtinguisherAssetMaster();

updateAssetMasterInspectionStatus(data);

return "Inspection saved successfully.";
}
function doPost(e) {

  try {

    if (!e || !e.postData || !e.postData.contents) {
      throw new Error("No POST data received.");
    }

    var data = JSON.parse(e.postData.contents);

    var message = saveInspectionFromWebApp(data);

    return ContentService
      .createTextOutput(JSON.stringify({
        success: true,
        message: message
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {

    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        message: error.message
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
// ==========================================
// CREATE UNIQUE EXTINGUISHER IDs
// ONE-TIME SETUP
// ==========================================

function createExtinguisherIDs() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var sheet = ss.getSheetByName(
    "Fire Extinguisher Asset Master"
  );

  if (!sheet) {
    throw new Error(
      "Fire Extinguisher Asset Master sheet not found."
    );
  }

  var lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    Logger.log("No Asset Master data found.");
    return;
  }

  var data = sheet
    .getRange(2, 1, lastRow - 1, 5)
    .getValues();

  var counters = {};

  var output = [];

  for (var i = 0; i < data.length; i++) {

    var assetNo =
      String(data[i][0]).trim();

    if (!assetNo) {
      output.push([""]);
      continue;
    }

    var key =
      assetNo.toUpperCase();

    if (!counters[key]) {
      counters[key] = 0;
    }

    counters[key]++;

    var extinguisherID =
      assetNo +
      "-" +
      String(counters[key]).padStart(2, "0");

    output.push([extinguisherID]);
  }

  sheet
    .getRange(2, 6, output.length, 1)
    .setValues(output);

  Logger.log(
    "Extinguisher IDs created successfully."
  );

  Logger.log(
    "Total records processed: " +
    output.length
  );
}
// ==========================================
// AUTO SET INSPECTION STATUS
// ==========================================

// ==========================================
// ASSET MASTER - INSPECTION STATUS SETUP
// AVAILABLE + NOT AVAILABLE + DEFECTIVE
// ==========================================

function setupInspectionStatus() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var sheet =
    ss.getSheetByName("Fire Extinguisher Asset Master");

  if (!sheet) {
    throw new Error(
      "Fire Extinguisher Asset Master sheet not found."
    );
  }

  var lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return;
  }

  // ==========================================
  // G = Inspection Status
  // H = Status / Defect Reason
  // ==========================================

  var statusRange =
    sheet.getRange(
      2,
      7,
      lastRow - 1,
      1
    );

  // ==========================================
  // BLANK STATUS = AVAILABLE
  // ==========================================

  var statusValues =
    statusRange.getValues();

  for (var i = 0; i < statusValues.length; i++) {

    if (!String(statusValues[i][0]).trim()) {
      statusValues[i][0] = "Available";
    }

  }

  statusRange.setValues(statusValues);

  // ==========================================
  // 3-OPTION DROPDOWN
  // ==========================================

  var rule =
    SpreadsheetApp.newDataValidation()
      .requireValueInList(
        [
          "Available",
          "Not Available",
          "Defective"
        ],
        true
      )
      .setAllowInvalid(false)
      .build();

  statusRange.setDataValidation(rule);

  // ==========================================
  // H HEADER
  // ==========================================

  sheet.getRange("H1")
    .setValue("Status / Defect Reason");

}
function testInspectionAssets() {

  var assets = getInspectionAssets();

  Logger.log("TOTAL ASSETS = " + assets.length);

  Logger.log(JSON.stringify(assets.slice(0, 10), null, 2));

}
// ==========================================
// AUTO UPDATE FIRE EXTINGUISHER ASSET MASTER
// MONTHLY INSPECTION + MANUAL STATUS
// ==========================================

// ==========================================
// AUTO UPDATE FIRE EXTINGUISHER ASSET MASTER
// INSPECTION RESULT BASED STATUS
// ==========================================

function updateAssetMasterInspectionStatus(data) {

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var masterSheet =
    ss.getSheetByName("Fire Extinguisher Asset Master");

  if (!masterSheet) {
    throw new Error(
      "Fire Extinguisher Asset Master sheet not found."
    );
  }

  var extinguisherId =
    String(data.extinguisherId || "").trim();

  if (!extinguisherId) {
    Logger.log("Extinguisher ID missing.");
    return;
  }

  var lastRow =
    masterSheet.getLastRow();

  if (lastRow < 2) {
    return;
  }

  // ==========================================
  // READ ASSET MASTER A:H
  // ==========================================

  var masterData =
    masterSheet
      .getRange(
        2,
        1,
        lastRow - 1,
        8
      )
      .getValues();

  // ==========================================
  // FIND EXTinguisher ID
  // COLUMN F = EXTinguisher ID
  // ==========================================

  for (var i = 0; i < masterData.length; i++) {

    var masterExtinguisherId =
      String(masterData[i][5]).trim();

    if (
      masterExtinguisherId !==
      extinguisherId
    ) {
      continue;
    }

    var row = i + 2;

    // ==========================================
    // FORM STATUS / REASON
    // ==========================================

    var formStatus =
      String(
        data.inspectionStatus || ""
      ).trim();

    var formReason =
      String(
        data.statusReason || ""
      ).trim();

    // ==========================================
    // SIX INSPECTION PARAMETERS
    // ==========================================

    var parameters = [

      "Safety Pin",

      "Discharge Hose",

      "Discharge Nozzle / Horn",

      "Pressure / Weight",

      "Cap Assembly",

      "Inspection Sticker"

    ];

    var values = [

      data.safetyPin,

      data.dischargeHose,

      data.dischargeNozzle,

      data.pressureWeight,

      data.capAssembly,

      data.inspectionSticker

    ];

    // ==========================================
    // FIND DEFECTS
    // ==========================================

    var defects = [];

    for (
      var d = 0;
      d < values.length;
      d++
    ) {

      var value =
        String(
          values[d] || ""
        ).trim();

      if (
        value &&
        value.toUpperCase() !== "OK"
      ) {

        defects.push(
          parameters[d] +
          " = " +
          value
        );

      }

    }

    // ==========================================
    // FINAL STATUS
    // ==========================================

    var finalStatus = "";
    var finalReason = "";

    // ==========================================
    // 1. ANY DEFECT FOUND
    // ==========================================

    if (
      defects.length > 0
    ) {

      finalStatus =
        "Defective";

      finalReason =
        defects.join(" | ");

    }

    // ==========================================
    // 2. MANUAL NOT AVAILABLE
    // ==========================================

    else if (
      formStatus ===
      "Not Available"
    ) {

      finalStatus =
        "Not Available";

      finalReason =
        formReason;

    }

    // ==========================================
    // 3. ALL SIX POINTS OK
    // ==========================================

    else {

      finalStatus =
        "Available";

      finalReason =
        "";

    }

    // ==========================================
    // UPDATE COLUMN G
    // INSPECTION STATUS
    // ==========================================

    masterSheet
      .getRange(row, 7)
      .setValue(finalStatus);

    // ==========================================
    // UPDATE COLUMN H
    // STATUS / DEFECT REASON
    // ==========================================

    masterSheet
      .getRange(row, 8)
      .setValue(finalReason);

    // ==========================================
    // LOG
    // ==========================================

    Logger.log(
      "Asset Master Updated: " +
      extinguisherId +
      " → " +
      finalStatus
    );

    Logger.log(
      "Reason: " +
      finalReason
    );

    break;
  }
}