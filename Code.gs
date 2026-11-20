/**
 * CHRIST APOSTOLIC CHURCH — OKE IYANU DISTRICT, OGIJO
 * Bio-data form backend.
 *
 * SETUP: paste this whole file into Extensions > Apps Script in your
 * Google Sheet (replacing the default code), then follow the setup guide
 * to deploy it as a Web App and connect it to index.html.
 */

var SHEET_NAME = "Responses";

var COLUMNS = [
  "Submitted At",
  "Full Name",
  "Date of Birth",
  "Sex",
  "Marital Status",
  "State of Origin",
  "LGA",
  "Residential Address",
  "Phone/WhatsApp No.",
  "Email",
  "Occupation/Profession",
  "Spouse's Name",
  "Number of Children",
  "Next of Kin Full Name",
  "Next of Kin Relationship",
  "Next of Kin Phone",
  "Next of Kin Address",
  "Date Joined Church",
  "Date of Baptism",
  "Ministry/Department",
  "Previous Church",
  "Emergency Contact Name",
  "Emergency Contact Relationship",
  "Emergency Contact Phone",
  "Signature"
];

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(COLUMNS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Handles the form submission from index.html
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = getSheet_();
    var row = COLUMNS.map(function (col) {
      return data[col] !== undefined ? data[col] : "";
    });
    sheet.appendRow(row);
    return ContentService.createTextOutput(
      JSON.stringify({ status: "ok" })
    ).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(
      JSON.stringify({ status: "error", message: err.message })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

// Adds a menu inside the Google Sheet, visible only to people who can open
// and edit this Sheet (i.e. whoever you've shared it with).
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Church admin tools")
    .addItem("Generate biodata PDF for selected row", "generatePdfForSelectedRow")
    .addToUi();
}

// Select any cell in a member's row in the Responses sheet, then run this
// from the "Church admin tools" menu to get a formatted, printable PDF.
function generatePdfForSelectedRow() {
  var ui = SpreadsheetApp.getUi();
  var sheet = getSheet_();
  var activeRange = sheet.getActiveRange();
  var rowIndex = activeRange.getRow();

  if (rowIndex < 2) {
    ui.alert("Please click on a cell inside a member's row (not the header row) and try again.");
    return;
  }

  var values = sheet.getRange(rowIndex, 1, 1, COLUMNS.length).getValues()[0];
  var record = {};
  COLUMNS.forEach(function (col, i) {
    record[col] = values[i];
  });

  if (!record["Full Name"]) {
    ui.alert("That row doesn't have a name in it. Please select a row with a member's data.");
    return;
  }

  var doc = DocumentApp.create("Biodata - " + record["Full Name"] + " - TEMP");
  var body = doc.getBody();
  body.setMarginTop(40).setMarginBottom(40).setMarginLeft(50).setMarginRight(50);

  var title = body.appendParagraph("CHRIST APOSTOLIC CHURCH");
  title.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  title.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  var sub = body.appendParagraph("Oke Iyanu, Mountain of Wonders District HQRT. Ogijo");
  sub.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  var sub2 = body.appendParagraph("Church Member Bio-Data");
  sub2.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  sub2.setBold(true);
  body.appendParagraph("");

  function section(heading, pairs) {
    var h = body.appendParagraph(heading);
    h.setHeading(DocumentApp.ParagraphHeading.HEADING3);
    pairs.forEach(function (pair) {
      var label = pair[0], value = pair[1];
      var p = body.appendParagraph("");
      p.appendText(label + ": ").setBold(true);
      p.appendText(String(value || "-"));
    });
    body.appendParagraph("");
  }

  section("Personal details", [
    ["Full name", record["Full Name"]],
    ["Date of birth", record["Date of Birth"]],
    ["Sex", record["Sex"]],
    ["Marital status", record["Marital Status"]],
    ["State of origin", record["State of Origin"]],
    ["LGA", record["LGA"]],
    ["Residential address", record["Residential Address"]],
    ["Phone/WhatsApp", record["Phone/WhatsApp No."]],
    ["Email", record["Email"]],
    ["Occupation", record["Occupation/Profession"]]
  ]);

  section("Family details", [
    ["Spouse's name", record["Spouse's Name"]],
    ["Number of children", record["Number of Children"]]
  ]);

  section("Next of kin", [
    ["Full name", record["Next of Kin Full Name"]],
    ["Relationship", record["Next of Kin Relationship"]],
    ["Phone", record["Next of Kin Phone"]],
    ["Address", record["Next of Kin Address"]]
  ]);

  section("Church details", [
    ["Date joined church", record["Date Joined Church"]],
    ["Date of baptism", record["Date of Baptism"]],
    ["Ministry/department", record["Ministry/Department"]],
    ["Previous church", record["Previous Church"]]
  ]);

  section("Emergency contact", [
    ["Name", record["Emergency Contact Name"]],
    ["Relationship", record["Emergency Contact Relationship"]],
    ["Phone", record["Emergency Contact Phone"]]
  ]);

  section("Declaration", [
    ["Signature", record["Signature"]],
    ["Submitted at", record["Submitted At"]]
  ]);

  doc.saveAndClose();

  var pdfBlob = DriveApp.getFileById(doc.getId()).getAs("application/pdf");
  var folder = getOrCreatePdfFolder_();
  var pdfFile = folder.createFile(pdfBlob).setName("Biodata - " + record["Full Name"] + ".pdf");

  DriveApp.getFileById(doc.getId()).setTrashed(true); // remove the temp Doc, keep only the PDF

  var htmlOutput = HtmlService
    .createHtmlOutput('<p>PDF ready for <b>' + record["Full Name"] + '</b>.</p>' +
      '<p><a href="' + pdfFile.getUrl() + '" target="_blank">Open / download the PDF</a></p>')
    .setWidth(350)
    .setHeight(120);
  ui.showModalDialog(htmlOutput, "Biodata PDF generated");
}

function getOrCreatePdfFolder_() {
  var folderName = "Biodata PDFs";
  var folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(folderName);
}
